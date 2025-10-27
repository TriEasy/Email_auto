# ================================================
# 🔐 Secure Configuration (Encrypted)
# ================================================
import subprocess
from cryptography.fernet import Fernet
import io

def load_encrypted_env():
    """Load and decrypt the .env file."""
    try:
        # Get encryption key from Windows Credential Manager
        result = subprocess.run(
            ['cmdkey', '/list:exchange_app_key'],
            capture_output=True,
            text=True
        )
        
        # If key not found in credential manager, try .env.key file
        if 'Target: exchange_app_key' not in result.stdout:
            if os.path.exists('.env.key'):
                with open('.env.key', 'rb') as f:
                    encryption_key = f.read()
            else:
                raise ValueError("Encryption key not found! Please store it in Windows Credential Manager.")
        else:
            # Extract key from credential manager (Windows-specific)
            # For simplicity, we'll use the .env.key file method
            # You can enhance this to actually retrieve from credential manager
            if os.path.exists('.env.key'):
                with open('.env.key', 'rb') as f:
                    encryption_key = f.read()
            else:
                encryption_key = os.getenv("ENV_ENCRYPTION_KEY")
                if not encryption_key:
                    raise ValueError("Encryption key not found!")
                encryption_key = encryption_key.encode()
        
        # Decrypt .env.encrypted file
        cipher = Fernet(encryption_key)
        with open('.env.encrypted', 'rb') as f:
            encrypted_data = f.read()
        
        decrypted_data = cipher.decrypt(encrypted_data)
        
        # Load into environment
        load_dotenv(stream=io.StringIO(decrypted_data.decode()))
        
    except FileNotFoundError:
        print("⚠️  .env.encrypted not found, trying regular .env file...")
        load_dotenv()
    except Exception as e:
        print(f"❌ Error loading encrypted environment: {e}")
        print("Falling back to regular .env file...")
        load_dotenv()

# Load encrypted configuration
load_encrypted_env()

EXCHANGE_USERNAME = os.getenv("EXCHANGE_USERNAME")
EXCHANGE_PASSWORD = os.getenv("EXCHANGE_PASSWORD")
EXCHANGE_EMAIL = os.getenv("EXCHANGE_EMAIL")
EXCHANGE_URL = os.getenv("EXCHANGE_URL")