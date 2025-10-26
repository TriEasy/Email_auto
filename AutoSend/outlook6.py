import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.folders import Calendar, Contacts, Tasks, Inbox
from exchangelib.items import Message
import pytz
# Assuming outlookhelpers2.py contains the original helper functions
from outlookhelpers2 import (
    get_riyadh_datetime,
    is_due_soon,
    format_due_date_for_email,
    get_reminder_subject,
    get_reminder_body,
    add_sent_category
)

# Configuration
EXCHANGE_USERNAME = "example"
EXCHANGE_PASSWORD = "example"
EXCHANGE_EMAIL = "example@sfda.gov.sa"
EXCHANGE_URL = 'example'

FOLDER_NAME = "Flag"
SENT_CATEGORY = "AutoReminderSent"

logging.basicConfig(level=logging.WARNING)
import urllib3
urllib3.disable_warnings()


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


def email_should_be_processed(msg):
    """Check if an email meets all the criteria for processing."""
    try:
        categories = msg.categories or []
        if SENT_CATEGORY in categories:
            return False
        return True
    except Exception as e:
        print(f"Error checking email criteria for '{msg.subject if hasattr(msg, 'subject') else 'unknown'}': {e}")
        return False


def get_original_recipients(msg):
    """Extract all original recipients from the email."""
    recipients = []
    for recipient_field in [msg.to_recipients, msg.cc_recipients, msg.bcc_recipients]:
        if recipient_field:
            for recipient in recipient_field:
                if recipient.email_address:
                    recipients.append(recipient.email_address)
    
    seen = set()
    unique_recipients = []
    for email in recipients:
        if email.lower() not in seen:
            seen.add(email.lower())
            unique_recipients.append(email)
    
    return "; ".join(unique_recipients)


def send_reminder_email(account, recipients, subject, body):
    """Create and send a new email using exchangelib."""
    try:
        mail = Message(
            account=account,
            subject=subject,
            body=body
        )
        mail.to_recipients = recipients.split('; ')
        mail.send()
        print(f"Reminder sent successfully to: {recipients}")
        return True
    except Exception as e:
        print(f"Error sending email via EWS: {e}")
        return False


def process_and_send_reminder(msg, now_riyadh, account):
    """Process a single email and send a reminder if due soon."""
    try:
        due_date = msg.flag.due_date
        
        if not is_due_soon(due_date, now_riyadh):
            return False
        
        original_recipients = get_original_recipients(msg)
        
        if not original_recipients:
            print(f"⚠️ No recipients found for '{msg.subject}'. Skipping reminder.")
            return False
        
        print(f"Email '{msg.subject}' is due soon. Sending reminder to: {original_recipients}")
        
        subject = get_reminder_subject(msg.subject)
        due_date_str = format_due_date_for_email(due_date, now_riyadh.tzinfo)
        body = get_reminder_body(msg.subject, due_date_str)
        
        email_sent = send_reminder_email(account, original_recipients, subject, body)
        
        if not email_sent:
            print(f"❌ Failed to send reminder for '{msg.subject}'")
            return False
        
        current_categories = msg.categories or []
        new_categories_list = add_sent_category(current_categories, SENT_CATEGORY)
        
        msg.categories = new_categories_list
        msg.save(update_fields=['categories'])
        
        print(f"✅ تم إرسال تذكير لـ: {msg.subject}")
        return True
    
    except Exception as e:
        print(f"Error processing reminder for '{msg.subject if hasattr(msg, 'subject') else 'unknown'}': {e}")
        return False


def main():
    """Main function to check flagged emails and send reminders."""
    print("Starting Exchange reminder process...")
    
    try:
        account = get_exchange_account()
        now_riyadh = get_riyadh_datetime()
        
        try:
            target_folder = account.inbox / "Flag"
            print(f"Using folder: {target_folder.name}")
        except Exception as e:
            print(f"Could not find Flag folder under Inbox: {e}")
            return
        
        # ✅ FIXED: Added 'flag' to the .only() call
        messages = target_folder.all().only(
            'subject', 'to_recipients', 'cc_recipients', 'bcc_recipients',
            'categories', 'flag'  # This was missing!
        )
        
        total_count = messages.count()
        print(f"Found {total_count} messages in '{FOLDER_NAME}' folder.")
        
        # Count messages with due dates
        flagged_count = 0
        reminder_count = 0
        
        for msg in messages:
            flag = getattr(msg, 'flag', None)
            if not flag or getattr(flag, 'status', None) != 'Flagged':
                continue
            
            if not getattr(flag, 'due_date', None):
                continue
            
            flagged_count += 1
            print(f"  → '{msg.subject}' - Due: {flag.due_date}")
            
            if email_should_be_processed(msg):
                if process_and_send_reminder(msg, now_riyadh, account):
                    reminder_count += 1
        
        print(f"\n📊 Summary:")
        print(f"  - Total messages in folder: {total_count}")
        print(f"  - Flagged with due dates: {flagged_count}")
        print(f"  - Reminders sent: {reminder_count}")
    
    except Exception as e:
        print(f"Error in main process: {e}")
        raise


if __name__ == "__main__":
    main()