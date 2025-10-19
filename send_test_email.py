import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- STEP 1: CONFIGURE YOUR EMAIL DETAILS ---
# --- Replace the placeholder text with your actual information ---

# Email Credentials
# IMPORTANT: For Gmail, this is NOT your regular password.
# It's a 16-character "App Password" you generate from your Google Account.
SENDER_EMAIL = "fahudeah@gmail.com"
APP_PASSWORD = "ncnaqxseyxktxwzf"

# Recipient's Email
RECIPIENT_EMAIL = "aldakheel.career@outlook.com"

# Email Content
SUBJECT = "Python Email Test"
BODY = "http://10.245.1.129:5666/"

# SMTP Server Settings (this is standard for Gmail)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# --- STEP 2: RUN THE SCRIPT (No changes needed below this line) ---

def send_email():
    """Connects to the SMTP server and sends the email."""
    print("Attempting to send email...")

    # Create the email message object
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = RECIPIENT_EMAIL
    msg['Subject'] = SUBJECT

    # Attach the body of the email
    msg.attach(MIMEText(BODY, 'plain'))

    try:
        # Establish a secure connection with the SMTP server
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()  # Upgrade the connection to be secure
        
        # Log in to your email account
        server.login(SENDER_EMAIL, APP_PASSWORD)
        
        # Send the email
        text = msg.as_string()
        server.sendmail(SENDER_EMAIL, RECIPIENT_EMAIL, text)
        
        # Disconnect from the server
        server.quit()
        
        print(f"✅ Email sent successfully to {RECIPIENT_EMAIL}")

    except smtplib.SMTPAuthenticationError:
        print("\n❌ Authentication Error: Failed to send email.")
        print("   Please check the following:")
        print("   1. Is your SENDER_EMAIL correct?")
        print("   2. Is your APP_PASSWORD correct? (It should be a 16-character App Password, not your regular password).")
        print("   3. Is 2-Step Verification enabled for your Google Account?")
        
    except Exception as e:
        print(f"\n❌ An error occurred: {e}")

# This part of the script runs the function when you execute the file
if __name__ == "__main__":
    send_email()