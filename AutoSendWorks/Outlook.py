import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.folders import Calendar, Contacts, Tasks, Inbox
from exchangelib.items import Message
import pytz
# Assuming outlookhelpers2.py contains the original helper functions
from OutlookHelper import (
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
        
        # ✅ FIXED: Use .all() without field restrictions to get all message data
        # We'll retrieve messages and let exchangelib load all fields automatically
        messages = list(target_folder.all())
        
        total_count = len(messages)
        print(f"Found {total_count} messages in '{FOLDER_NAME}' folder.")
        
        if total_count == 0:
            print("No messages found in folder. Exiting.")
            return
        
        # Count messages with due dates
        flagged_count = 0
        reminder_count = 0
        
        # Process each message - check for reminder properties
        for msg in messages:
            try:
                # Debug: Print what we're seeing
                print(f"\n{'='*60}")
                print(f"Processing: {msg.subject}")
                print(f"{'='*60}")
                
                # Check for reminder properties (the actual properties used by Exchange)
                reminder_is_set = getattr(msg, 'reminder_is_set', None)
                reminder_due_by = getattr(msg, 'reminder_due_by', None)
                
                print(f"  reminder_is_set: {reminder_is_set}")
                print(f"  reminder_due_by: {reminder_due_by}")
                
                # Skip if no reminder is set
                if not reminder_is_set:
                    print(f"  ⚠️ No reminder set - SKIPPING")
                    continue
                
                # Skip if no due date
                if not reminder_due_by:
                    print(f"  ⚠️ Reminder set but no due date - SKIPPING")
                    continue
                
                flagged_count += 1
                print(f"  ✅ Has reminder with due date: {reminder_due_by}")
                
                # Check categories
                categories = msg.categories or []
                print(f"  Current categories: {categories}")
                print(f"  Has {SENT_CATEGORY}? {SENT_CATEGORY in categories}")
                
                # Check if already processed
                if email_should_be_processed(msg):
                    print(f"  → Proceeding to check if due soon and send reminder...")
                    print(f"  Current Riyadh time: {now_riyadh}")
                    
                    # Check if due soon (within 2 days)
                    is_due = is_due_soon(reminder_due_by, now_riyadh)
                    print(f"  Is due within 2 days? {is_due}")
                    
                    if is_due:
                        # Get recipients
                        original_recipients = get_original_recipients(msg)
                        
                        if not original_recipients:
                            print(f"  ⚠️ No recipients found. Skipping.")
                            continue
                        
                        print(f"  Recipients: {original_recipients}")
                        
                        # Send reminder
                        subject = get_reminder_subject(msg.subject)
                        due_date_str = format_due_date_for_email(reminder_due_by, now_riyadh.tzinfo)
                        body = get_reminder_body(msg.subject, due_date_str)
                        
                        email_sent = send_reminder_email(account, original_recipients, subject, body)
                        
                        if email_sent:
                            # Mark as sent
                            current_categories = msg.categories or []
                            new_categories_list = add_sent_category(current_categories, SENT_CATEGORY)
                            msg.categories = new_categories_list
                            msg.save(update_fields=['categories'])
                            
                            reminder_count += 1
                            print(f"  ✅ REMINDER SENT SUCCESSFULLY")
                        else:
                            print(f"  ❌ Failed to send reminder")
                    else:
                        print(f"  ℹ️ Not due within 2 days yet")
                else:
                    print(f"  ⚠️ Already processed (has {SENT_CATEGORY} category)")
                        
            except Exception as e:
                print(f"❌ Error processing message '{getattr(msg, 'subject', 'unknown')}': {e}")
                import traceback
                traceback.print_exc()
                continue
        
        print(f"\n📊 Summary:")
        print(f"  - Total messages in folder: {total_count}")
        print(f"  - Flagged with due dates: {flagged_count}")
        print(f"  - Reminders sent: {reminder_count}")
    
    except Exception as e:
        print(f"Error in main process: {e}")
        import traceback
        traceback.print_exc()
        raise


if __name__ == "__main__":
    main()