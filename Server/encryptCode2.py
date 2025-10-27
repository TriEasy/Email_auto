def load_encrypted_env():
    """Load and decrypt the .env file."""
    try:
        # Get encryption key from environment variable
        encryption_key = os.getenv("ENV_ENCRYPTION_KEY")
        
        if not encryption_key:
            # Try .env.key file as fallback
            if os.path.exists('.env.key'):
                with open('.env.key', 'rb') as f:
                    encryption_key = f.read().decode()
            else:
                raise ValueError("Encryption key not found! Set ENV_ENCRYPTION_KEY environment variable.")
        
        # Decrypt .env.encrypted file
        cipher = Fernet(encryption_key.encode())
        with open('.env.encrypted', 'rb') as f:
            encrypted_data = f.read()
        
        decrypted_data = cipher.decrypt(encrypted_data)
        
        # Load into environment
        load_dotenv(stream=io.StringIO(decrypted_data.decode()))
        print("✅ Loaded encrypted environment variables")
        
    except FileNotFoundError:
        print("⚠️  .env.encrypted not found, trying regular .env file...")
        load_dotenv()
    except Exception as e:
        print(f"❌ Error loading encrypted environment: {e}")
        print("Falling back to regular .env file...")
        load_dotenv()