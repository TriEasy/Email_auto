import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE, Message, Q
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
import urllib3
import pytz
from dotenv import load_dotenv
from cryptography.fernet import Fernet
import io

from reply.outlookHelp import (
    get_riyadh_datetime,
    is_due_soon,
    add_sent_category
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
# 📬 Email Processing Helpers
# ================================================
def email_should_be_processed(msg):
    """Check if an email meets all the criteria for processing."""
    try:
        categories = [c.lower() for c in (msg.categories or [])]
        return SENT_CATEGORY.lower() not in categories
    except Exception as e:
        print(f"Error checking email criteria for '{getattr(msg, 'subject', 'unknown')}': {e}")
        return False


# ================================================
# 🧠 Helper: Identify Non-Responders (TO only)
# ================================================
def get_non_responders(account, original_msg):
    """Return list of recipients (TO only) who have not replied to the original email."""
    try:
        # ✅ Only include 'To' recipients — ignore CC
        all_recipients = set()
        for r in (original_msg.to_recipients or []):
            if hasattr(r, "email_address") and r.email_address:
                all_recipients.add(r.email_address.lower())

        conversation_id = getattr(original_msg, 'conversation_id', None)
        subject = original_msg.subject or ""

        replied_senders = set()
        if conversation_id:
            related_msgs = account.inbox.filter(conversation_id=conversation_id)
        else:
            related_msgs = account.inbox.filter(subject__contains=subject)

        for m in related_msgs:
            if m.id != original_msg.id and m.sender:
                replied_senders.add(m.sender.email_address.lower())

        non_responders = [r for r in all_recipients if r not in replied_senders]
        return non_responders

    except Exception as e:
        print(f"⚠️ Error determining non-responders: {e}")
        return []


# ================================================
# 📤 Send Reminder to Non-Responders
# ================================================
def send_reminder_to_non_responders(account, msg):
    """Send reminder email only to TO recipients who have not replied."""
    try:
        non_responders = get_non_responders(account, msg)
        if not non_responders:
            print("✅ Everyone in TO replied — skipping reminder.")
            return False

        due_date_str = msg.reminder_due_by.strftime('%Y-%m-%d %H:%M') if msg.reminder_due_by else "غير محدد"
        subject = f"🔔 تذكير بالمتابعة: {msg.subject}"
        body = (
            f"السلام عليكم ورحمة الله وبركاته،\n\n"
            f"نود تذكيركم بأن الرسالة التالية بلغت موعدها المحدد للمتابعة:\n\n"
            f"📩 العنوان: {msg.subject}\n"
            f"📅 الموعد: {due_date_str}\n\n"
            f"يرجى اتخاذ اللازم.\n\n"
            f"قسم المتابعة - هيئة الغذاء والدواء"
        )

        new_msg = Message(
            account=account,
            subject=subject,
            body=body,
            to_recipients=list(non_responders)
        )
        new_msg.send_and_save()
        print(f"✅ Sent reminder to non-responders (TO only): {', '.join(non_responders)}")
        return True

    except Exception as e:
        print(f"❌ Error sending reminder to non-responders for '{msg.subject}': {e}")
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
            print(f"❌ Could not find '{FOLDER_NAME}' folder under Inbox: {e}")
            return

        messages = list(target_folder.all())
        total_count = len(messages)
        print(f"📬 Found {total_count} messages in '{FOLDER_NAME}' folder.")

        if total_count == 0:
            print("ℹ️ No messages found. Exiting.")
            return

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
                print(f"  reminder_due_by: {reminder_due_by}")

                if not reminder_is_set or not reminder_due_by:
                    print("  ⚠️ Skipping: No reminder or due date.")
                    continue

                flagged_count += 1

                categories = msg.categories or []
                if not email_should_be_processed(msg):
                    print(f"  ⚠️ Already processed ({SENT_CATEGORY} exists). Skipping.")
                    continue

                is_due = is_due_soon(reminder_due_by, now_riyadh)
                print(f"  Is due within 2 days? {is_due}")

                if not is_due:
                    print("  ℹ️ Not due yet. Skipping.")
                    continue

                # ✅ Send reminder only to TO non-responders
                if send_reminder_to_non_responders(account, msg):
                    try:
                        msg.refresh()
                        msg.categories = add_sent_category(categories, SENT_CATEGORY)
                        msg.save(update_fields=['categories'])
                        reminder_count += 1
                        print("  ✅ Reminder marked as sent.")
                    except Exception as e:
                        print(f"⚠️ Could not save category due to ChangeKey issue: {e}")

            except Exception as e:
                print(f"❌ Error processing '{getattr(msg, 'subject', 'unknown')}': {e}")
                import traceback
                traceback.print_exc()
                continue

        print(f"\n📊 Summary:")
        print(f"  - Total messages: {total_count}")
        print(f"  - Flagged with due dates: {flagged_count}")
        print(f"  - Reminders sent: {reminder_count}")

    except Exception as e:
        print(f"❌ Error in main process: {e}")
        import traceback
        traceback.print_exc()


# ================================================
# 🏁 Entry Point
# ================================================
if __name__ == "__main__":
    main()