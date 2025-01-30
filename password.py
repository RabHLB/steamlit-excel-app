import bcrypt

# Hashed passwords
hashed_passwords = [
    "$2b$12$WDhf5n2M1fe7R2rBBhn7e.jT0kE5umEx608XD7dLFUr55dB8zZk3a",  # johndoe
    "$2b$12$iLipEEh.OarddPx5qXPw1eatZS4TV7CRR1ang2ZkBminzWa8TQNcy"   # janesmith
]

# Corresponding plaintext passwords
plaintext_passwords = ["password123", "password456"]

# Verify passwords
for plain, hashed in zip(plaintext_passwords, hashed_passwords):
    result = bcrypt.checkpw(plain.encode(), hashed.encode())
    print(f"Password: {plain}, Match: {result}")
