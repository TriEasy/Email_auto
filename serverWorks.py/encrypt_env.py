from cryptography.fernet import Fernet
import os

def encrypt_env_file():
    """Encrypts the .env file and generates an encryption key."""
    
    # Check if .env exists
    if not os.path.exists('.env'):
        print("❌ Error: .env file not found!")
        return
    
    # Generate encryption key
    key = Fernet.generate_key()
    cipher = Fernet(key)
    
    # Read and encrypt .env file
    print("🔒 Encrypting .env file...")
    with open('.env', 'rb') as f:
        env_data = f.read()
    
    encrypted_data = cipher.encrypt(env_data)
    
    # Save encrypted file
    with open('.env.encrypted', 'wb') as f:
        f.write(encrypted_data)
    
    # Save key to a separate file
    with open('.env.key', 'wb') as f:
        f.write(key)
    
    print("✅ Encryption complete!")
    print(f"✅ Created: .env.encrypted")
    print(f"✅ Created: .env.key")
    print("\n" + "="*60)
    print("🔑 YOUR ENCRYPTION KEY (save this securely!):")
    print("="*60)
    print(key.decode())
    print("="*60)
    print("\n⚠️  IMPORTANT SECURITY STEPS:")
    print("1. Copy the key above and store it securely")
    print("2. Delete the original .env file: del .env (Windows) or rm .env (Mac/Linux)")
    print("3. Delete .env.key file: del .env.key (Windows) or rm .env.key (Mac/Linux)")
    print("4. Add the key to Windows Credential Manager (see next steps)")

if __name__ == "__main__":
    encrypt_env_file()