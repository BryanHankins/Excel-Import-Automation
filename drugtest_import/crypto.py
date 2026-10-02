"""Encryption at rest for record fields and uploaded images, plus password hashing."""
import hashlib
import hmac
import os
import secrets

from cryptography.fernet import Fernet, InvalidToken

KEY_ENV = "DRUGTEST_KEY"


class ConfigError(Exception):
    pass


def generate_key() -> str:
    return Fernet.generate_key().decode("ascii")


class Cipher:
    """Fernet (AES-128-CBC + HMAC-SHA256) for values, and keyed HMACs for lookups.

    One master key is configured; separate sub-keys are derived for the
    searchable index and the session cookie so a leak of one doesn't expose
    the others.
    """

    def __init__(self, key: str):
        try:
            self._fernet = Fernet(key.encode("ascii"))
        except (ValueError, TypeError) as e:
            raise ConfigError(f"{KEY_ENV} is not a valid key. Generate one with: python -m drugtest_import.manage gen-key") from e
        raw = key.encode("ascii")
        self._index_key = hmac.new(raw, b"drugtest-index-v1", hashlib.sha256).digest()
        self.session_secret = hmac.new(raw, b"drugtest-session-v1", hashlib.sha256).hexdigest()

    @classmethod
    def from_env(cls) -> "Cipher":
        key = os.environ.get(KEY_ENV)
        if not key:
            raise ConfigError(f"Set {KEY_ENV}. Generate one with: python -m drugtest_import.manage gen-key")
        return cls(key)

    def encrypt(self, value: str | None) -> bytes | None:
        return None if value is None else self._fernet.encrypt(value.encode("utf-8"))

    def decrypt(self, token: bytes | None) -> str | None:
        if token is None:
            return None
        try:
            return self._fernet.decrypt(token).decode("utf-8")
        except InvalidToken as e:
            raise ConfigError("Data can't be decrypted with this key. Is DRUGTEST_KEY correct?") from e

    def encrypt_bytes(self, data: bytes) -> bytes:
        return self._fernet.encrypt(data)

    def decrypt_bytes(self, token: bytes) -> bytes:
        return self._fernet.decrypt(token)

    def index(self, value: str | None) -> str | None:
        """Deterministic keyed hash so encrypted values can still be matched exactly."""
        if not value:
            return None
        normalized = " ".join(value.lower().split())
        return hmac.new(self._index_key, normalized.encode("utf-8"), hashlib.sha256).hexdigest()


# scrypt parameters: ~16 MB memory, well above OWASP's minimum recommendation
SCRYPT = {"n": 2**14, "r": 8, "p": 1}


def hash_password(password: str) -> str:
    salt = secrets.token_bytes(16)
    digest = hashlib.scrypt(password.encode("utf-8"), salt=salt, dklen=32, **SCRYPT)
    return f"scrypt${salt.hex()}${digest.hex()}"


def verify_password(password: str, stored: str) -> bool:
    try:
        scheme, salt_hex, digest_hex = stored.split("$")
    except ValueError:
        return False
    if scheme != "scrypt":
        return False
    digest = hashlib.scrypt(password.encode("utf-8"), salt=bytes.fromhex(salt_hex), dklen=32, **SCRYPT)
    return hmac.compare_digest(digest.hex(), digest_hex)
