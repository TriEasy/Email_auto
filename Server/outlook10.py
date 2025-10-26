import os
import logging
from datetime import timedelta
from exchangelib import Credentials, Account, Configuration, DELEGATE
from exchangelib.protocol import BaseProtocol, NoVerifyHTTPAdapter
from exchangelib.folders import Inbox
from exchangelib.items import Message
import urllib3
import pytz
from dotenv import load_dotenv

from outlookHelp import (
    get_riyadh_datetime,
    is_due_soon,
    format_due_date_for_email,
    get_reminder_subject,
    get_reminder_body,
    add_sent_category
)

# ================================================
# 🔐 Secure Configuration
# ================================================
load_dotenv()  # Load .env file if present (optional)

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
# 📧 Exchange Connection
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
        categories = msg.categories or []
        return SENT_CATEGORY not in categories
    except Exception as e:
        print(f"Error checking email criteria for '{getattr(msg, 'subject', 'unknown')}': {e}")
        return False


def get_original_recipients(msg):
    """Extract all original recipients from the email."""
    recipients = []
    for field in [msg.to_recipients, msg.cc_recipients, msg.bcc_recipients]:
        if field:
            for recipient in field:
                if recipient.email_address:
                    recipients.append(recipient.email_address)
    # Remove duplicates
    unique = list({r.lower(): r for r in recipients}.values())
    return "; ".join(unique)


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
        print(f"✅ Reminder sent successfully to: {recipients}")
        return True
    except Exception as e:
        print(f"❌ Error sending email via EWS: {e}")
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
            print(f"📁 Using folder: {target_folder.name}"_
