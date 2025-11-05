import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.folders import Inbox
from exchangelib.items import Message
from exchangelib import Q
import urllib3
import pytz
from dotenv import load_dotenv
from cryptography.fernet import Fernet
import io

from outlookHelp import (
    get_riyadh_datetime,
    is_due_soon,
    add_sent_category,
    format_due_date_for_email
)

# ================================================
# 🔐 Secure Configuration (Encrypted)
# ================================================
def load_encrypted_env():
    """Load and decrypt the .env file."""
    try:
        encryption_key = os.getenv("ENV_ENCRYPTION_KEY")
        if not encryption_key:
            if os.path.exists('.env.key'):
                with open('.env.key', 'rb') as f:
                    encryption_key = f.read().decode()
            else:
                raise ValueError("Encryption key not found! Set ENV_ENCRYPTION_KEY environment variable.")
        
        cipher = Fernet(encryption_key.encode())
        with open('.env.encrypted', 'rb') as f:
            encrypted_data = f.read()
        decrypted_data = cipher.decrypt(encrypted_data)
        load_dotenv(stream=io.StringIO(decrypted_data.decode()))
        print("✅ Loaded encrypted environment variables")
    except FileNotFoundError:
        print("⚠️  .env.encrypted not found, trying regular .env file...")
        load_dotenv()
    except Exception as e:
        print(f"❌ Error loading encrypted environment: {e}")
        print("Falling back to regular .env file...")
        load_dotenv()

# Load encrypted environment variables
load_encrypted_env()

EXCHANGE_USERNAME = os.getenv("EXCHANGE_USERNAME")
EXCHANGE_PASSWORD = os.getenv("EXCHANGE_PASSWORD")
EXCHANGE_EMAIL = os.getenv("EXCHANGE_EMAIL")
EXCHANGE_URL = os.getenv("EXCHANGE_URL")

FOLDER_NAME = "Flag"
SENT_CATEGORY = "AutoReminderSent"

if not all([EXCHANGE_USERNAME, EXCHANGE_PASSWORD, EXCHANGE_EMAIL, EXCHANGE_URL]):
    raise ValueError(
        "Missing Exchange environment variables. Please set EXCHANGE_USERNAME, "
        "EXCHANGE_PASSWORD, EXCHANGE_EMAIL, and EXCHANGE_URL."
    )

# ================================================
# ⚙️ Logging & Security
# ================================================
logging.basicConfig(level=logging.WARNING)
urllib3.disable_warnings()

# ================================================
# 🔧 Exchange Connection
# ================================================
def get_exchange_account():
    """Establish connection to the Exchange account."""
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
    credentials = Credentials(EXCHANGE_USERNAME, EXCHANGE_PASSWORD)
    config = Configuration(credentials=credentials, service_endpoint=EXCHANGE_URL)
    account = Account(
        primary_smtp_address=EXCHANGE_EMAIL,
        config=config,
        autodiscover=False,
        access_type=DELEGATE
    )
    return account


# ================================================
# 📧 Email Processing Helpers
# ================================================
def get_all_recipients(msg):
    """Extract all unique email addresses from To and CC fields."""
    recipients = set()
    
    if hasattr(msg, 'to_recipients') and msg.to_recipients:
        for recipient in msg.to_recipients:
            if hasattr(recipient, 'email_address'):
                recipients.add(recipient.email_address.lower())
    
    if hasattr(msg, 'cc_recipients') and msg.cc_recipients:
        for recipient in msg.cc_recipients:
            if hasattr(recipient, 'email_address'):
                recipients.add(recipient.email_address.lower())
    
    return recipients


def get_responders_to_message(account, original_msg):
    """
    Find all email addresses that have replied to the original message.
    Searches the entire mailbox for replies based on conversation ID or subject.
    """
    responders = set()
    try:
        # Method 1: Use conversation_id
        if hasattr(original_msg, 'conversation_id') and original_msg.conversation_id:
            conversation_id = original_msg.conversation_id
            replies = account.inbox.filter(conversation_id=conversation_id)
            
            for reply in replies:
                if reply.id == original_msg.id:
                    continue
                if hasattr(reply, 'sender') and reply.sender and hasattr(reply.sender, 'email_address'):
                    responders.add(reply.sender.email_address.lower())

        # Method 2: Fallback - search by subject
        if not responders and hasattr(original_msg, 'subject') and original_msg.subject:
            subject = original_msg.subject
            replies = account.inbox.filter(subject__contains=subject)
            for reply in replies:
                if reply.id == original_msg.id:
                    continue
                if hasattr(reply, 'sender') and reply.sender and hasattr(reply.sender, 'email_address'):
                    responders.add(reply.sender.email_address.lower())

        print(f"  📊 Found {len(responders)} responders: {responders}")

    except Exception as e:
        print(f"  ⚠️ Error finding responders: {e}")

    return responders


def get_non_responders(original_msg, account):
    """Get list of recipients who haven't responded."""
    all_recipients = get_all_recipients(original_msg)
    responders = get_responders_to_message(account, original_msg)
    
    if hasattr(original_msg, 'sender') and original_msg.sender and hasattr(original_msg.sender, 'email_address'):
        all_recipients.discard(original_msg.sender.email_address.lower())
    
    non_responders = all_recipients - responders
    
    print(f"  👥 All recipients: {len(all_recipients)}")
    print(f"  ✅ Responders: {len(responders)}")
    print(f"  ⏰ Non-responders: {len(non_responders)}")
    
    return non_responders


def email_should_be_processed(msg):
    """Check if an email meets all the criteria for processing."""
    try:
        categories = [c.lower() for c in (msg.categories or [])]
        return SENT_CATEGORY.lower() not in categories
    except Exception as e:
        print(f"Error checking email criteria for '{getattr(msg, 'subject', 'unknown')}': {e}")
        return False


def send_reminder_to_non_responders(msg, non_responders, account):
    """Send reminder email as a reply only to non-responders."""
    try:
        if not non_responders:
            print("  ℹ️ All recipients have responded. No reminder needed.")
            return True
        
        # Format due date in Riyadh time
        due_date_str = None
        if hasattr(msg, "reminder_due_by") and msg.reminder_due_by:
            due_date_str = format_due_date_for_email(msg.reminder_due_by)

        subject = f"🔔 تذكير بالمتابعة: {msg.subject}"
        body = (
            f"السلام عليكم ورحمة الله وبركاته،\n\n"
            f"نود تذكيركم بأن الرسالة التالية بلغت موعدها المحدد للمتابعة:\n\n"
            f"📩 العنوان: {msg.subject}\n"
            f"📅 الموعد (بتوقيت الرياض): {due_date_str or 'غير محدد'}\n\n"
            f"يرجى اتخاذ اللازم.\n\n"
            f"قسم المتابعة - هيئة الغذاء والدواء"
        )

        # Create reply
        reply = msg.create_reply_all(subject=subject, body=body)
        reply.to_recipients = list(non_responders)
        reply.cc_recipients = []
        reply.send()

        print(f"  ✅ Sent reminder to {len(non_responders)} non-responders")
        print(f"  📧 Recipients: {', '.join(non_responders)}")
        msg.refresh()
        return True

    except Exception as e:
        print(f"  ❌ Error sending reminder for '{msg.subject}': {e}")
        import traceback
        traceback.print_exc()
        return False


# ================================================
# 🚀 Main Logic
# ================================================
def main():
    """Main function to check flagged emails and send reminders."""
    print("🔄 Starting Exchange reminder process...")

    try:
        account = get_exchange_account()
        now_riyadh = get_riyadh_datetime()

        try:
            target_folder = account.inbox / FOLDER_NAME
            print(f"📁 Using folder: {target_folder.name}")
        except Exception as e:
            print(f"❌ Could not find '{FOLDER_NAME}' folder: {e}")
            return

        messages = list(target_folder.all())
        print(f"📬 Found {len(messages)} messages in '{FOLDER_NAME}'.")

        flagged_count = 0
        reminder_count = 0

        for msg in messages:
            try:
                print(f"\n{'='*60}")
                print(f"Processing: {msg.subject}")
                print(f"{'='*60}")

                reminder_is_set = getattr(msg, 'reminder_is_set', None)
                reminder_due_by = getattr(msg, 'reminder_due_by', None)

                print(f"  reminder_is_set: {reminder_is_set}")
                print(f"  reminder_due_by (UTC): {reminder_due_by}")

                if not reminder_is_set or not reminder_due_by:
                    print("  ⚠️ Skipping: No reminder set.")
                    continue

                flagged_count += 1

                # Show Riyadh time in logs
                try:
                    riyadh_time_str = format_due_date_for_email(reminder_due_by)
                    print(f"  reminder_due_by (Riyadh): {riyadh_time_str}")
                except Exception as e:
                    print(f"  ⚠️ Could not format Riyadh time: {e}")

                if not email_should_be_processed(msg):
                    print(f"  ⚠️ Already processed ({SENT_CATEGORY}). Skipping.")
                    continue

                if not is_due_soon(reminder_due_by, now_riyadh):
                    print("  ℹ️ Not due yet. Skipping.")
                    continue

                non_responders = get_non_responders(msg, account)
                if send_reminder_to_non_responders(msg, non_responders, account):
                    msg.refresh()
                    msg.categories = add_sent_category(msg.categories or [], SENT_CATEGORY)
                    msg.save(update_fields=['categories'])
                    reminder_count += 1
                    print("  ✅ Reminder marked as sent.")

            except Exception as e:
                print(f"❌ Error processing '{getattr(msg, 'subject', 'unknown')}': {e}")
                import traceback
                traceback.print_exc()
                continue

        print(f"\n📊 Summary:")
        print(f"  - Total messages: {len(messages)}")
        print(f"  - Flagged with due dates: {flagged_count}")
        print(f"  - Reminders sent: {reminder_count}")

    except Exception as e:
        print(f"❌ Error in main process: {e}")
        import traceback
        traceback.print_exc()


# ================================================
# 🎯 Entry Point
# ================================================
if __name__ == "__main__":
    main()