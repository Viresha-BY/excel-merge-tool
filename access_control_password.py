import yaml
import bcrypt

def load_users():
    with open("users.yaml", "r") as f:
        return yaml.safe_load(f)["users"]

def verify_user(username, password):
    users = load_users()
    for user in users:
        if user["username"] == username and bcrypt.checkpw(password.encode(), user["password_hash"].encode()):
            return user["role"]
    return None
