# Excel Merge Tool – Secure Setup & Usage Guide

## 1. Clone the Repository

```sh
git clone https://github.com/Viresha-BY/excel-merge-tool.git
cd excel-merge-tool
```

---

## 2. Install Requirements

Create a virtual environment (recommended):

```sh
python3 -m venv venv
source venv/bin/activate
```

Install dependencies:

```sh
pip install -r requirements.txt
```

---

## 3. Set Up User Authentication

### a) Generate Password Hashes

Open Python:

```sh
python
```
Then run:

```python
import bcrypt
print(bcrypt.hashpw(b"yourpassword", bcrypt.gensalt()).decode())
```
Copy the output string (starts with `$2b$...`).

---

### b) Create Your `users.yaml` File

**Do NOT commit this file to git.**

Create a file called `users.yaml` in the project folder:

```yaml
users:
  - username: yourname
    password_hash: "$2b$12$examplehashgoeshere"
    role: admin
```

Add more users as needed. Use the hash you generated above.

---

### c) Confirm `.gitignore` Protects Secrets

Ensure these lines are in your `.gitignore` file:

```
users.yaml
.secrets.toml
.env
```

---

## 4. Run the App

Start Streamlit:

```sh
streamlit run app.py
```

---

## 5. Login

- Open the link shown in your terminal (usually http://localhost:8501).
- Use your username and password from the `users.yaml` file.

---

## 6. Adding/Removing Users

- To add a user: Generate a new hash, add a new entry in `users.yaml`.
- To remove a user: Delete their entry.
- **Restart the app** for changes to take effect.

---

## 7. Deploying (Streamlit Cloud or other)

- Ensure your `users.yaml` is uploaded **securely** (use Streamlit Cloud secrets or upload after deploy).
- Never commit real secrets to git!

---

## 8. Troubleshooting

- **ModuleNotFoundError:** Install missing libraries (`pip install pyyaml bcrypt`).
- **Login fails:** Double-check the username and hash in `users.yaml`.
- **Permission errors:** Make sure your `users.yaml` is in the same directory as `app.py`.

---

## Example `users.yaml`

```yaml
users:
  - username: admin
    password_hash: "$2b$12$abc123..."
    role: admin
  - username: bob
    password_hash: "$2b$12$xyz456..."
    role: operator
```

---

## Example `.gitignore`

```
users.yaml
.secrets.toml
.env
```

---

**Enjoy your secure Excel Merge Tool!**
