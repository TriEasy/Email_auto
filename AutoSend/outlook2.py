"""
This script checks a specific Outlook folder for flagged emails
due soon and sends a reminder using win32com.client.
"""
import win32com.client
from OutlookHelpers import (
    get_riyadh_datetime,
    is_due_soon,
    format_due_date_for_email,
    get_reminder_subject,
    get_reminder_body,
    add_sent_category
)

# Configuration
FOLDER_NAME = "Flag"
TARGET_FOLDER_PATH = f"Inbox/{FOLDER_NAME}"


def get_target_folder(outlook, folder_path):
    """
    Navigate to the target folder in Outlook.
    
    Args:
        outlook: Outlook application object
        folder_path: Path like "Inbox/Flag"
    
    Returns:
        The target folder object
    """
    namespace = outlook.GetNamespace("MAPI")
    
    # Split the path and navigate through folders
    folder_names = folder_path.split("/")
    current_folder = namespace.GetDefaultFolder(6)  # 6 = olFolderInbox
    
    # Navigate through subfolders if needed
    for folder_name in folder_names[1:]:  # Skip 'Inbox' as we already have it
        current_folder = current_folder.Folders[folder_name]
    
    return current_folder


def email_should_be_processed(msg):
    """
    Check if an email meets all the criteria for processing.
    
    Args:
        msg: Outlook mail item
    
    Returns:
        bool: True if should be processed, False otherwise
    """
    try:
        # Check 1: Is it a mail item (Class 43)?
        if msg.Class != 43:
            return False
        
        # Check 2: Is it marked as a task?
        if not msg.IsMarkedAsTask:
            return False
        
        # Check 3: Does it have a due date?
        if msg.TaskDueDate is None:
            return False
        
        # Check 4: Has a reminder already been sent?
        categories = msg.Categories or ""
        if "AutoReminderSent" in categories:
            return False
        
        return True
    
    except Exception as e:
        print(f"Error checking email criteria: {e}")
        return False


def send_reminder_email(outlook, recipient, subject, body):
    """
    Create and send a new email.
    
    Args:
        outlook: Outlook application object
        recipient: Email address to send to
        subject: Email subject
        body: Email body
    """
    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.To = recipient
        mail.Subject = subject
        mail.Body = body
        mail.Send()
    except Exception as e:
        print(f"Error sending email: {e}")
        raise


def process_and_send_reminder(msg, now_riyadh, outlook):
    """
    Process a single email and send a reminder if due soon.
    
    Args:
        msg: Outlook mail item
        now_riyadh: Current datetime in Riyadh
        outlook: Outlook application object
    
    Returns:
        bool: True if reminder was sent, False otherwise
    """
    try:
        due_date = msg.TaskDueDate
        
        # Check if it's due within the 2-day window
        if not is_due_soon(due_date, now_riyadh):
            return False
        
        # Email is due soon - prepare reminder
        print(f"Email '{msg.Subject}' is due soon. Sending reminder.")
        
        # Get formatted text
        subject = get_reminder_subject(msg.Subject)
        due_date_str = format_due_date_for_email(due_date, now_riyadh.tzinfo)
        body = get_reminder_body(msg.Subject, due_date_str)
        
        # Send the reminder email
        send_reminder_email(outlook, RECIPIENT_EMAIL, subject, body)
        
        # Mark the original email as "Sent"
        new_categories = add_sent_category(msg.Categories)
        msg.Categories = new_categories
        msg.Save()
        
        print(f"✅ تم إرسال تذكير لـ: {msg.Subject}")
        return True
    
    except Exception as e:
        print(f"Error processing reminder for '{msg.Subject}': {e}")
        return False


def main():
    """Main function to check flagged emails and send reminders."""
    print("Starting Outlook reminder process...")
    
    try:
        # Initialize Outlook
        outlook = win32com.client.Dispatch("Outlook.Application")
        
        # Get current time in Riyadh
        now_riyadh = get_riyadh_datetime()
        
        # Get the target folder
        target_folder = get_target_folder(outlook, TARGET_FOLDER_PATH)
        
        # Get all messages in the folder
        messages = target_folder.Items
        total_count = messages.Count
        print(f"Found {total_count} messages in '{TARGET_FOLDER_PATH}'.")
        
        reminder_count = 0
        
        # Loop through each message
        for msg in messages:
            # Check if email should be processed
            if email_should_be_processed(msg):
                # Process and send reminder if due soon
                if process_and_send_reminder(msg, now_riyadh, outlook):
                    reminder_count += 1
        
        print(f"\n📨 تم إرسال {reminder_count} تذكير من مجلد '{FOLDER_NAME}'.")
    
    except Exception as e:
        print(f"Error in main process: {e}")
        raise


if __name__ == "__main__":
    main()