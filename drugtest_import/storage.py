"""SQLite storage: encrypted records, uploads awaiting review, users and the audit log."""
import csv
import json
import os
import sqlite3
import tempfile
import threading
import uuid
from datetime import date, datetime, timezone

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Font, PatternFill

from .crypto import Cipher, ConfigError, hash_password
from .schema import FIELDS, normalize, validate

ROLES = ["admin", "reviewer", "viewer"]
UPLOAD_STATES = ("reading", "ready", "failed")

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value BLOB);

CREATE TABLE IF NOT EXISTS records (
    id INTEGER PRIMARY KEY,
    employee_id BLOB,           -- encrypted
    employee_index TEXT,        -- keyed hash for exact matching
    name BLOB NOT NULL,         -- encrypted
    name_index TEXT NOT NULL,
    department TEXT,
    test_date TEXT NOT NULL,
    test_type TEXT NOT NULL,
    result BLOB NOT NULL,       -- encrypted
    notes BLOB,                 -- encrypted
    source_file BLOB,           -- encrypted (file names often contain names)
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS records_lookup ON records (name_index, test_date, test_type);

CREATE TABLE IF NOT EXISTS uploads (
    id TEXT PRIMARY KEY,
    filename BLOB NOT NULL,     -- encrypted
    status TEXT NOT NULL,
    extraction BLOB,            -- encrypted JSON
    error TEXT,
    uploaded_by TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY,
    username TEXT NOT NULL UNIQUE COLLATE NOCASE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL,
    last_login TEXT
);

CREATE TABLE IF NOT EXISTS audit_log (
    id INTEGER PRIMARY KEY,
    at TEXT NOT NULL,
    actor TEXT NOT NULL,
    action TEXT NOT NULL,
    target TEXT,
    detail TEXT
);
CREATE TRIGGER IF NOT EXISTS audit_log_no_update BEFORE UPDATE ON audit_log
BEGIN SELECT RAISE(ABORT, 'audit log is append-only'); END;
CREATE TRIGGER IF NOT EXISTS audit_log_no_delete BEFORE DELETE ON audit_log
BEGIN SELECT RAISE(ABORT, 'audit log is append-only'); END;
"""

EXPORT_HEADERS = FIELDS + ["SourceFile", "EnteredAt", "EnteredBy"]
COLUMN_WIDTHS = [12, 24, 16, 12, 10, 13, 40, 24, 20, 14]
RESULT_FILLS = {"Positive": "FECACA", "Pending": "FEF08A", "Inconclusive": "FEF08A", "Refused": "FED7AA"}


def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def atomic_write(path: str, write) -> None:
    """Call write(tmp_path), then move the result into place so readers never see a partial file."""
    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp_path = tempfile.mkstemp(dir=directory, suffix=".tmp")
    os.close(fd)
    try:
        write(tmp_path)
        os.replace(tmp_path, path)
    except BaseException:
        os.unlink(tmp_path)
        raise


class Database:
    """All persistent state. Safe to share between threads."""

    def __init__(self, path: str, cipher: Cipher, uploads_dir: str | None = None):
        self.path = path
        self.cipher = cipher
        self.uploads_dir = uploads_dir or os.path.join(os.path.dirname(os.path.abspath(path)), "uploads")
        os.makedirs(self.uploads_dir, mode=0o700, exist_ok=True)
        self._lock = threading.RLock()
        self.conn = sqlite3.connect(path, check_same_thread=False)
        self.conn.row_factory = sqlite3.Row
        self.conn.execute("PRAGMA foreign_keys = ON")
        self.conn.executescript(SCHEMA)
        self._check_key()

    def _check_key(self):
        """Fail fast if the configured key isn't the one this database was created with."""
        row = self._fetchone("SELECT value FROM meta WHERE key = 'key_check'")
        if row is None:
            self._execute("INSERT INTO meta (key, value) VALUES ('key_check', ?)", (self.cipher.encrypt("ok"),))
        elif self.cipher.decrypt(row["value"]) != "ok":
            raise ConfigError("DRUGTEST_KEY doesn't match this database.")

    def close(self):
        with self._lock:
            self.conn.close()

    def _execute(self, sql, params=()):
        with self._lock, self.conn:
            return self.conn.execute(sql, params)

    def _fetchall(self, sql, params=()):
        with self._lock:
            return self.conn.execute(sql, params).fetchall()

    def _fetchone(self, sql, params=()):
        with self._lock:
            return self.conn.execute(sql, params).fetchone()

    # --- audit ------------------------------------------------------------

    def log(self, actor: str, action: str, target: str | None = None, detail: str | None = None) -> None:
        self._execute(
            "INSERT INTO audit_log (at, actor, action, target, detail) VALUES (?, ?, ?, ?, ?)",
            (now(), actor, action, target, detail),
        )

    def audit_entries(self, limit: int = 500) -> list[dict]:
        rows = self._fetchall("SELECT * FROM audit_log ORDER BY id DESC LIMIT ?", (limit,))
        return [dict(row) for row in rows]

    # --- records ----------------------------------------------------------

    def add_record(self, record: dict, actor: str, source_file: str | None = None) -> int:
        """Insert a validated record. Only the file's base name is kept, not its full path."""
        c = self.cipher
        values = {
            "employee_id": c.encrypt(record.get("EmployeeID")),
            "employee_index": c.index(record.get("EmployeeID")),
            "name": c.encrypt(record["Name"]),
            "name_index": c.index(record["Name"]),
            "department": record.get("Department"),
            "test_date": record["TestDate"],
            "test_type": record["TestType"],
            "result": c.encrypt(record["Result"]),
            "notes": c.encrypt(record.get("Notes")),
            "source_file": c.encrypt(os.path.basename(source_file)) if source_file else None,
            "created_at": now(),
            "created_by": actor,
        }
        with self._lock, self.conn:
            cursor = self.conn.execute(
                f"INSERT INTO records ({', '.join(values)}) VALUES ({', '.join('?' * len(values))})",
                list(values.values()),
            )
            self.conn.execute(
                "INSERT INTO audit_log (at, actor, action, target) VALUES (?, ?, 'record.create', ?)",
                (now(), actor, str(cursor.lastrowid)),
            )
        return cursor.lastrowid

    def find_duplicate(self, record: dict) -> dict | None:
        """Return an existing record for the same person, date and test type, if any.

        Different employee IDs mean different people; a missing ID on either side still matches.
        """
        employee_index = self.cipher.index(record.get("EmployeeID"))
        row = self._fetchone(
            "SELECT * FROM records WHERE name_index = ? AND test_date = ? AND test_type = ? "
            "AND (employee_index IS ? OR ? IS NULL OR employee_index IS NULL) LIMIT 1",
            (self.cipher.index(record.get("Name")), record.get("TestDate"), record.get("TestType"),
             employee_index, employee_index),
        )
        return self._to_record(row) if row else None

    def records(self) -> list[dict]:
        rows = self._fetchall("SELECT * FROM records ORDER BY test_date DESC, id DESC")
        return [self._to_record(row) for row in rows]

    def count(self) -> int:
        return self._fetchone("SELECT count(*) FROM records")[0]

    def import_csv(self, path: str, actor: str) -> tuple[int, int]:
        """Load records from a CSV produced by earlier versions.

        Returns (imported, skipped); rows that fail validation are skipped.
        """
        imported = skipped = 0
        with open(path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                record = normalize(row)
                if validate(record):
                    skipped += 1
                    continue
                self.add_record(record, actor, source_file=path)
                imported += 1
        self.log(actor, "records.import_csv", os.path.basename(path), f"{imported} imported, {skipped} skipped")
        return imported, skipped

    def _to_record(self, row: sqlite3.Row) -> dict:
        c = self.cipher
        return {
            "id": row["id"],
            "EmployeeID": c.decrypt(row["employee_id"]),
            "Name": c.decrypt(row["name"]),
            "Department": row["department"],
            "TestDate": row["test_date"],
            "TestType": row["test_type"],
            "Result": c.decrypt(row["result"]),
            "Notes": c.decrypt(row["notes"]),
            "SourceFile": c.decrypt(row["source_file"]),
            "EnteredAt": row["created_at"],
            "EnteredBy": row["created_by"],
        }

    # --- uploads awaiting review -----------------------------------------

    def _upload_path(self, upload_id: str) -> str:
        return os.path.join(self.uploads_dir, f"{uuid.UUID(upload_id)}.bin")

    def create_upload(self, filename: str, data: bytes, actor: str) -> str:
        """Store an uploaded image encrypted on disk until it's reviewed."""
        upload_id = str(uuid.uuid4())
        path = self._upload_path(upload_id)
        with open(os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), "wb") as f:
            f.write(self.cipher.encrypt_bytes(data))
        self._execute(
            "INSERT INTO uploads (id, filename, status, uploaded_by, created_at) VALUES (?, ?, 'reading', ?, ?)",
            (upload_id, self.cipher.encrypt(os.path.basename(filename)), actor, now()),
        )
        self.log(actor, "upload.create", upload_id)
        return upload_id

    def upload_image(self, upload_id: str) -> bytes:
        with open(self._upload_path(upload_id), "rb") as f:
            return self.cipher.decrypt_bytes(f.read())

    def set_extraction(self, upload_id: str, extraction: dict | None, error: str | None = None) -> None:
        status = "failed" if error else "ready"
        payload = self.cipher.encrypt(json.dumps(extraction)) if extraction is not None else None
        self._execute(
            "UPDATE uploads SET status = ?, extraction = ?, error = ? WHERE id = ?",
            (status, payload, error, upload_id),
        )

    def get_upload(self, upload_id: str) -> dict | None:
        row = self._fetchone("SELECT * FROM uploads WHERE id = ?", (upload_id,))
        return self._to_upload(row) if row else None

    def pending_uploads(self) -> list[dict]:
        rows = self._fetchall("SELECT * FROM uploads ORDER BY created_at, rowid")
        return [self._to_upload(row) for row in rows]

    def delete_upload(self, upload_id: str) -> None:
        """Remove the stored image and its row once it's been saved or skipped."""
        try:
            os.unlink(self._upload_path(upload_id))
        except FileNotFoundError:
            pass
        self._execute("DELETE FROM uploads WHERE id = ?", (upload_id,))

    def _to_upload(self, row: sqlite3.Row) -> dict:
        extraction = row["extraction"]
        return {
            "id": row["id"],
            "filename": self.cipher.decrypt(row["filename"]),
            "status": row["status"],
            "extraction": json.loads(self.cipher.decrypt(extraction)) if extraction else None,
            "error": row["error"],
            "uploaded_by": row["uploaded_by"],
            "created_at": row["created_at"],
        }

    # --- users ------------------------------------------------------------

    def create_user(self, username: str, password: str, role: str, actor: str) -> int:
        if role not in ROLES:
            raise ValueError(f"Role must be one of {ROLES}")
        if len(password) < 12:
            raise ValueError("Password must be at least 12 characters")
        username = username.strip()
        if not username:
            raise ValueError("Username is required")
        try:
            cursor = self._execute(
                "INSERT INTO users (username, password_hash, role, created_at) VALUES (?, ?, ?, ?)",
                (username, hash_password(password), role, now()),
            )
        except sqlite3.IntegrityError as e:
            raise ValueError(f"User {username!r} already exists") from e
        self.log(actor, "user.create", username, f"role={role}")
        return cursor.lastrowid

    def get_user(self, user_id: int) -> dict | None:
        row = self._fetchone("SELECT * FROM users WHERE id = ?", (user_id,))
        return dict(row) if row else None

    def get_user_by_name(self, username: str) -> dict | None:
        row = self._fetchone("SELECT * FROM users WHERE username = ?", (username.strip(),))
        return dict(row) if row else None

    def users(self) -> list[dict]:
        return [dict(row) for row in self._fetchall("SELECT * FROM users ORDER BY username")]

    def user_count(self) -> int:
        return self._fetchone("SELECT count(*) FROM users")[0]

    def update_user(self, user_id: int, actor: str, *, role: str | None = None, active: bool | None = None,
                    password: str | None = None) -> None:
        user = self.get_user(user_id)
        if not user:
            raise ValueError("No such user")
        if role is not None:
            if role not in ROLES:
                raise ValueError(f"Role must be one of {ROLES}")
            self._execute("UPDATE users SET role = ? WHERE id = ?", (role, user_id))
            self.log(actor, "user.role", user["username"], role)
        if active is not None:
            self._execute("UPDATE users SET active = ? WHERE id = ?", (int(active), user_id))
            self.log(actor, "user.enable" if active else "user.disable", user["username"])
        if password is not None:
            if len(password) < 12:
                raise ValueError("Password must be at least 12 characters")
            self._execute("UPDATE users SET password_hash = ? WHERE id = ?", (hash_password(password), user_id))
            self.log(actor, "user.password", user["username"])

    def active_admin_count(self) -> int:
        return self._fetchone("SELECT count(*) FROM users WHERE role = 'admin' AND active = 1")[0]

    def record_login(self, user_id: int) -> None:
        self._execute("UPDATE users SET last_login = ? WHERE id = ?", (now(), user_id))

    # --- export -----------------------------------------------------------

    def export_xlsx(self, path: str, records: list[dict] | None = None) -> int:
        """Write records (default: all) to a formatted Excel workbook. Returns rows written.

        The workbook is NOT encrypted - it's for handing to people who need it.
        """
        records = self.records() if records is None else records
        wb = Workbook()
        ws = wb.active
        ws.title = "Drug Tests"
        ws.append(EXPORT_HEADERS)
        for record in records:
            row = [record[header] for header in EXPORT_HEADERS]
            try:
                row[FIELDS.index("TestDate")] = date.fromisoformat(record["TestDate"])
            except (TypeError, ValueError):
                pass
            ws.append(row)

        for cell in ws[1]:
            cell.font = Font(bold=True)
        for index, width in enumerate(COLUMN_WIDTHS, start=1):
            ws.column_dimensions[ws.cell(row=1, column=index).column_letter].width = width
        for cell in ws["D"][1:]:
            cell.number_format = "yyyy-mm-dd"
        # Employee IDs stay text so leading zeros survive
        for cell in ws["A"][1:]:
            cell.number_format = "@"
        ws.freeze_panes = "A2"
        last_col = ws.cell(row=1, column=len(EXPORT_HEADERS)).column_letter
        last_row = max(len(records) + 1, 2)
        ws.auto_filter.ref = f"A1:{last_col}{last_row}"
        result_col = ws.cell(row=1, column=FIELDS.index("Result") + 1).column_letter
        for result, color in RESULT_FILLS.items():
            ws.conditional_formatting.add(
                f"A2:{last_col}{last_row}",
                FormulaRule(formula=[f'${result_col}2="{result}"'], fill=PatternFill("solid", fgColor=color)),
            )
        atomic_write(path, wb.save)
        return len(records)
