import bcrypt

# Replace these passwords with the ones you want to hash
passwords = ["password123", "password456"]

# Hash the passwords manually
hashed_passwords = [bcrypt.hashpw(password.encode(), bcrypt.gensalt()).decode() for password in passwords]

# Print the hashed passwords
print("Hashed Passwords:", hashed_passwords)
