import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.folders import Calendar, Contacts, Tasks, Inbox
from exchangelib.items import Message
import pytz
# Assuming OutlookHelpers.py contains the original helper functions
from OutlookHelpers import (
    get_riyadh_datetime,
    is_due_soon,
    format_due_date_for_email,
    get_reminder_subject,
    get_reminder_body,
    add_sent_category
)

# Configuration
# !! REPLACE WITH YOUR ACTUAL CREDENTIALS AND SERVER INFO !!
EXCHANGE_USERNAME = "your_username"  # e.g., 'user.name'
EXCHANGE_PASSWORD = "your_password"
EXCHANGE_EMAIL = f"{EXCHANGE_USERNAME}@sfda.gov.sa" # e.g., 'user.name@sfda.gov.sa'
EXCHANGE_URL = 'example.com' # Replace with your EWS service endpoint URL

# Folder configuration
# In EWS, the "Flag" folder is usually the Tasks folder, 
# or a specific search is needed. We will search the Inbox for flagged items.
# The original script targeted "Inbox/Flag", which is a MAPI view, 
# but in EWS we typically filter the main folder (Inbox).
FOLDER_NAME = "Inbox"
SENT_CATEGORY = "AutoReminderSent"

# Configure logging to suppress verbose exchangelib output if needed
logging.basicConfig(level=logging.WARNING)
# Suppress SSL warnings (as in your snippet) - use with caution
import urllib3
urllib3.disable_warnings()


def get_exchange_account():
    """
    Establish connection to the Exchange account.
    """
    # Use NoVerifyHTTPAdapter to handle potential SSL certificate issues (like in your snippet)
    BaseProtocol.HTTP_ADAPTER_CLS = NoVerifyHTTPAdapter
    
    # Setup credentials and configuration
    credentials = Credentials(EXCHANGE_USERNAME, EXCHANGE_PASSWORD)
    config = Configuration(credentials=credentials, service_endpoint=EXCHANGE_URL)
    
    # Initialize the Account object
    account = Account(
        primary_smtp_address=EXCHANGE_EMAIL,
        config=config,
        autodiscover=False,
        access_type=DELEGATE  # Assuming delegate access as per your snippet
    )
    return account


def email_should_be_processed(msg):
    """
    Check if an email meets all the criteria for processing, adapted for exchangelib.
    
    Args:
        msg: exchangelib Message item
    
    Returns:
        bool: True if should be processed, False otherwise
    """
    try:
        # Check 1: Is it marked as a task (IsMarkedAsTask in win32com corresponds to a TaskRequest item, 
        # but here we rely on flag status which is a property of the Message item itself)
        # We'll filter for flagged items during the folder search, 
        # but check for an actual due date here.
        
        # Check 2: Does it have a due date (TaskDueDate in win32com is Task_due_date in exchangelib)?
        if not hasattr(msg, 'task_due_date') or msg.task_due_date is None:
            return False
            
        # Check 3: Has a reminder already been sent?
        categories = msg.categories or []
        if SENT_CATEGORY in categories:
            return False
        
        return True
    
    except Exception as e:
        # Catch exchangelib-specific errors or general exceptions
        print(f"Error checking email criteria for '{msg.subject if hasattr(msg, 'subject') else 'unknown'}': {e}")
        return False


def get_original_recipients(msg):
    """
    Extract all original recipients from the email (To, CC, BCC) from exchangelib item.
    
    Args:
        msg: exchangelib Message item
    
    Returns:
        str: Semicolon-separated list of email addresses
    """
    recipients = []
    
    # Exchangelib stores recipients in explicit fields, which contain EmailAddress objects
    # We combine To, Cc, and Bcc fields
    for recipient_field in [msg.to_recipients, msg.cc_recipients, msg.bcc_recipients]:
        if recipient_field:
            for recipient in recipient_field:
                if recipient.email_address:
                    recipients.append(recipient.email_address)
    
    # Remove duplicates
    seen = set()
    unique_recipients = []
    for email in recipients:
        if email.lower() not in seen:
            seen.add(email.lower())
            unique_recipients.append(email)
    
    return "; ".join(unique_recipients)


def send_reminder_email(account, recipients, subject, body):
    """
    Create and send a new email using exchangelib.
    
    Args:
        account: exchangelib Account object
        recipients: Email addresses to send to (semicolon-separated)
        subject: Email subject
        body: Email body
    """
    try:
        # Create a new Message item
        mail = Message(
            account=account,
            subject=subject,
            body=body
        )
        
        # Set recipients. exchangelib expects a list of email strings for .to_recipients
        mail.to_recipients = recipients.split('; ') # Convert semicolon-separated string to list
        
        # Send the email
        mail.send()
        print(f"Reminder sent successfully to: {recipients}")
        return True
        
    except Exception as e:
        print(f"Error sending email via EWS: {e}")
        return False


def process_and_send_reminder(msg, now_riyadh, account):
    """
    Process a single email and send a reminder if due soon.
    
    Args:
        msg: exchangelib Message item
        now_riyadh: Current datetime in Riyadh (pytz-aware)
        account: exchangelib Account object
    
    Returns:
        bool: True if reminder was sent, False otherwise
    """
    try:
        # exchangelib stores due date in task_due_date
        due_date = msg.task_due_date
        
        # Check if it's due within the 2-day window
        # The is_due_soon helper must handle datetime objects from exchangelib
        if not is_due_soon(due_date, now_riyadh):
            return False
        
        # Email is due soon - get original recipients
        original_recipients = get_original_recipients(msg)
        
        if not original_recipients:
            print(f"⚠️ No recipients found for '{msg.subject}'. Skipping reminder.")
            return False
        
        # Prepare reminder
        print(f"Email '{msg.subject}' is due soon. Sending reminder to: {original_recipients}")
        
        # Get formatted text
        subject = get_reminder_subject(msg.subject)
        # Pass the timezone from now_riyadh to format the date correctly
        due_date_str = format_due_date_for_email(due_date, now_riyadh.tzinfo) 
        body = get_reminder_body(msg.subject, due_date_str)
        
        # Send the reminder email to original recipients
        email_sent = send_reminder_email(account, original_recipients, subject, body)
        
        if not email_sent:
            print(f"❌ Failed to send reminder for '{msg.subject}'")
            return False
        
        # Mark the original email as "Sent" by updating categories and saving
        current_categories = msg.categories or []
        new_categories_list = add_sent_category(current_categories, SENT_CATEGORY)
        
        # Update the item properties and save it back to Exchange
        msg.categories = new_categories_list
        # We need to explicitly call save on the object to update it on the server
        msg.save(update_fields=['categories']) 
        
        print(f"✅ تم إرسال تذكير لـ: {msg.subject}")
        return True
    
    except Exception as e:
        print(f"Error processing reminder for '{msg.subject if hasattr(msg, 'subject') else 'unknown'}': {e}")
        return False


def main():
    """Main function to check flagged emails and send reminders using exchangelib."""
    print("Starting Exchange reminder process...")
    
    try:
        # 1. Connect to Exchange
        account = get_exchange_account()
        
        # 2. Get current time in Riyadh (pytz-aware)
        now_riyadh = get_riyadh_datetime()
        
        # 3. Get the target folder (Inbox)
        target_folder = account.inbox
        
        # 4. Filter for flagged messages.
        # exchangelib uses Item.flag to represent the flag status.
        # We filter for items that are 'Flagged' and have a due date.
        # Note: 'flag__ne=None' generally means "any flag status except None/clear".
        # We use flag.status='Flagged' which is more explicit for an active flag.
        # exchangelib does not have a direct property for 'IsMarkedAsTask', 
        # so we rely on having a flag and a task_due_date.
        
        # The filter checks for:
        # - Items with a flag status of 'Flagged' (active)
        # - Items that have a task_due_date set (otherwise they aren't 'tasks')
        messages = target_folder.all().filter(
            flag__status='Flagged',
            task_due_date__isnull=False
        ).only('subject', 'to_recipients', 'cc_recipients', 'bcc_recipients', 
               'categories', 'task_due_date')

        total_count = messages.count()
        print(f"Found {total_count} flagged and due messages in '{FOLDER_NAME}'.")
        
        reminder_count = 0
        
        # 5. Loop through each message
        for msg in messages:
            # Check if email should be processed (e.g., check for 'AutoReminderSent')
            if email_should_be_processed(msg):
                # Process and send reminder if due soon
                if process_and_send_reminder(msg, now_riyadh, account):
                    reminder_count += 1
        
        print(f"\n📨 تم إرسال {reminder_count} تذكير من مجلد '{FOLDER_NAME}'.")
    
    except Exception as e:
        print(f"Error in main process: {e}")
        raise


if __name__ == "__main__":
    # Ensure your OutlookHelpers.py functions are adjusted to handle 
    # exchangelib datetime objects (which are timezone-aware)
    main()