import os
import hashlib
from passlib.context import CryptContext

pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

EMAIL_PEPPER = os.getenv("APP_EMAIL_PEPPER", "")
PWD_PEPPER = os.getenv("APP_PWD_PEPPER", "")


def normalize_email(email: str) -> str:
    return (email or "").strip().lower()


def email_hash(email: str) -> str:
    e = normalize_email(email)
    raw = (e + "|" + EMAIL_PEPPER).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def hash_password(password: str) -> str:
    p = (password or "") + PWD_PEPPER
    return pwd_context.hash(p)


def verify_password(password: str, hashed: str) -> bool:
    p = (password or "") + PWD_PEPPER
    return pwd_context.verify(p, hashed)
