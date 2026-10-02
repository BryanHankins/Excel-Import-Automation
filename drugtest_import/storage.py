"""SQLite storage: form types, encrypted entries, uploads awaiting review, users and the audit log."""
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
from .forms import BUILT_IN_FORMS, DRUG_TEST, LEGACY_DRUG_TEST_COLUMNS, FormTemplate, normalize, validate

ROLES = ["admin", "reviewer", "viewer"]

SCHEMA = """
CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value BLOB);

CREATE TABLE IF NOT EXISTS forms (
    key TEXT PRIMARY KEY,
    definition TEXT NOT NULL,   -- FormTemplate JSON (field names, not data - not encrypted)
    updated_at TEXT NOT NULL,
    updated_by TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS entries (
    id INTEGER PRIMARY KEY,
    form_key TEXT NOT NULL REFERENCES forms (key),
    data BLOB NOT NULL,         -- encrypted JSON of field values
    duplicate_index TEXT,       -- keyed hash of the form's duplicate-check fields
    record_date TEXT,           -- the form's date field, for sorting and date filters
    source_file BLOB,           -- encrypted (file names often contain names)
    created_at TEXT NOT NULL,
    created_by TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS entries_form ON entries (form_key, record_date);
CREATE INDEX IF NOT EXISTS entries_duplicates ON entries (form_key, duplicate_index);

CREATE TABLE IF NOT EXISTS uploads (
    id TEXT PRIMARY KEY,
    form_key TEXT NOT NULL REFERENCES forms (key),
    filename BLOB NOT NULL,     -- encrypted
    status TEXT NOT NULL,       -- reading | ready | failed
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

COLUMN_WIDTH = {"longtext": 40, "text": 22, "id": 14, "date": 12, "number": 10, "choice": 14}
FLAG_FILL = "FECACA"


class FormInUse(ValueError):
    pass


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
        self._check_key()
        legacy_records = self._columns("records")
        legacy_uploads = self._columns("uploads") and "form_key" not in self._columns("uploads")
        if legacy_uploads:
            self.conn.execute("ALTER TABLE uploads RENAME TO uploads_v1")
        self.conn.executescript(SCHEMA)
        if not self._fetchone("SELECT 1 FROM forms LIMIT 1"):
            for form in BUILT_IN_FORMS:
                self._write_form(form, "system")
        if legacy_records or legacy_uploads:
            self._migrate_v1(legacy_records, legacy_uploads)

    def _columns(self, table: str) -> list[str]:
        return [row[1] for row in self.conn.execute(f"PRAGMA table_info({table})")]

    def _check_key(self):
        """Fail fast if the configured key isn't the one this database was created with."""
        self.conn.execute("CREATE TABLE IF NOT EXISTS meta (key TEXT PRIMARY KEY, value BLOB)")
        row = self._fetchone("SELECT value FROM meta WHERE key = 'key_check'")
        if row is None:
            self._execute("INSERT INTO meta (key, value) VALUES ('key_check', ?)", (self.cipher.encrypt("ok"),))
        elif self.cipher.decrypt(row["value"]) != "ok":
            raise ConfigError("DRUGTEST_KEY doesn't match this database.")

    def _migrate_v1(self, legacy_records: bool, legacy_uploads: bool) -> None:
        """Move data from the drug-test-only tables of earlier versions into the form-type tables."""
        c = self.cipher
        with self._lock, self.conn:
            if legacy_records:
                for row in self.conn.execute("SELECT * FROM records ORDER BY id").fetchall():
                    values = {
                        "employee_id": c.decrypt(row["employee_id"]), "name": c.decrypt(row["name"]),
                        "department": row["department"], "test_date": row["test_date"],
                        "test_type": row["test_type"], "result": c.decrypt(row["result"]),
                        "notes": c.decrypt(row["notes"]),
                    }
                    self._insert_entry(DRUG_TEST, values, row["created_by"], row["source_file"], row["created_at"])
                self.conn.execute("DROP TABLE records")
            if legacy_uploads:
                self.conn.execute(
                    "INSERT INTO uploads (id, form_key, filename, status, extraction, error, uploaded_by, created_at) "
                    "SELECT id, 'drug-test', filename, 'reading', NULL, NULL, uploaded_by, created_at FROM uploads_v1"
                )
                self.conn.execute("DROP TABLE uploads_v1")

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

    # --- form types -------------------------------------------------------

    def forms(self) -> list[FormTemplate]:
        rows = self._fetchall("SELECT definition FROM forms ORDER BY key")
        forms = [FormTemplate.model_validate_json(row["definition"]) for row in rows]
        return sorted(forms, key=lambda f: f.name.lower())

    def get_form(self, key: str) -> FormTemplate | None:
        row = self._fetchone("SELECT definition FROM forms WHERE key = ?", (key,))
        return FormTemplate.model_validate_json(row["definition"]) if row else None

    def _write_form(self, form: FormTemplate, actor: str) -> None:
        self.conn.execute(
            "INSERT INTO forms (key, definition, updated_at, updated_by) VALUES (?, ?, ?, ?) "
            "ON CONFLICT (key) DO UPDATE SET definition = excluded.definition, "
            "updated_at = excluded.updated_at, updated_by = excluded.updated_by",
            (form.key, form.model_dump_json(), now(), actor),
        )

    def save_form(self, form: FormTemplate, actor: str) -> None:
        """Create or update a form type.

        Once a form type has entries its fields can't be removed or change type,
        so stored data always matches a field. Labels, choices, hints and
        required/duplicate flags can still change.
        """
        existing = self.get_form(form.key)
        if existing and self.count(form.key):
            new = {f.key: f for f in form.fields}
            for old in existing.fields:
                if old.key not in new:
                    raise FormInUse(f"'{old.label}' can't be removed: entries already use it.")
                if new[old.key].type != old.type:
                    raise FormInUse(f"'{old.label}' can't change type: entries already use it.")
        with self._lock, self.conn:
            self._write_form(form, actor)
            if existing and existing.duplicate_keys != form.duplicate_keys:
                self._reindex(form)
            self.conn.execute(
                "INSERT INTO audit_log (at, actor, action, target) VALUES (?, ?, ?, ?)",
                (now(), actor, "form.update" if existing else "form.create", form.key),
            )

    def delete_form(self, key: str, actor: str) -> None:
        if self.count(key) or self._fetchone("SELECT 1 FROM uploads WHERE form_key = ?", (key,)):
            raise FormInUse("This form type has entries or uploads, so it can't be deleted.")
        if len(self.forms()) <= 1:
            raise FormInUse("At least one form type is needed.")
        with self._lock, self.conn:
            self.conn.execute("DELETE FROM forms WHERE key = ?", (key,))
            self.conn.execute(
                "INSERT INTO audit_log (at, actor, action, target) VALUES (?, ?, 'form.delete', ?)",
                (now(), actor, key),
            )

    # --- entries ----------------------------------------------------------

    def _duplicate_index(self, form: FormTemplate, values: dict) -> str | None:
        keys = form.duplicate_keys
        if not keys or not all(values.get(k) for k in keys):
            return None
        return self.cipher.index(json.dumps([values[k] for k in keys]))

    def _insert_entry(self, form, values, actor, source_file_token, created_at) -> int:
        cursor = self.conn.execute(
            "INSERT INTO entries (form_key, data, duplicate_index, record_date, source_file, created_at, created_by) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (form.key, self.cipher.encrypt(json.dumps(values)), self._duplicate_index(form, values),
             values.get(form.date_field) if form.date_field else None, source_file_token, created_at, actor),
        )
        return cursor.lastrowid

    def _reindex(self, form: FormTemplate) -> None:
        for row in self.conn.execute("SELECT id, data FROM entries WHERE form_key = ?", (form.key,)).fetchall():
            values = json.loads(self.cipher.decrypt(row["data"]))
            self.conn.execute("UPDATE entries SET duplicate_index = ? WHERE id = ?",
                              (self._duplicate_index(form, values), row["id"]))

    def add_entry(self, form: FormTemplate, values: dict, actor: str, source_file: str | None = None) -> int:
        """Insert validated values. Only the file's base name is kept, not its full path."""
        values = {key: values.get(key) for key in form.keys}
        token = self.cipher.encrypt(os.path.basename(source_file)) if source_file else None
        with self._lock, self.conn:
            entry_id = self._insert_entry(form, values, actor, token, now())
            self.conn.execute(
                "INSERT INTO audit_log (at, actor, action, target) VALUES (?, ?, 'entry.create', ?)",
                (now(), actor, f"{form.key}/{entry_id}"),
            )
        return entry_id

    def find_duplicate(self, form: FormTemplate, values: dict) -> dict | None:
        """Return an existing entry of this form type whose duplicate-check fields all match."""
        index = self._duplicate_index(form, values)
        if not index:
            return None
        row = self._fetchone("SELECT * FROM entries WHERE form_key = ? AND duplicate_index = ? LIMIT 1",
                             (form.key, index))
        return self._to_entry(row) if row else None

    def entries(self, form_key: str) -> list[dict]:
        rows = self._fetchall(
            "SELECT * FROM entries WHERE form_key = ? ORDER BY record_date DESC, id DESC", (form_key,))
        return [self._to_entry(row) for row in rows]

    def count(self, form_key: str | None = None) -> int:
        if form_key is None:
            return self._fetchone("SELECT count(*) FROM entries")[0]
        return self._fetchone("SELECT count(*) FROM entries WHERE form_key = ?", (form_key,))[0]

    def import_csv(self, path: str, actor: str, form: FormTemplate = DRUG_TEST) -> tuple[int, int]:
        """Load entries from a CSV whose headers are field keys or labels (or the old drug-test headers).

        Returns (imported, skipped); rows that fail validation are skipped.
        """
        form = self.get_form(form.key) or form
        by_header = {f.key: f.key for f in form.fields} | {f.label.lower(): f.key for f in form.fields}
        if form.key == DRUG_TEST.key:
            by_header |= {h.lower(): k for h, k in LEGACY_DRUG_TEST_COLUMNS.items()}
        imported = skipped = 0
        with open(path, newline="", encoding="utf-8-sig") as f:
            for row in csv.DictReader(f):
                mapped = {by_header.get(h, by_header.get((h or "").lower())): v for h, v in row.items()}
                values = normalize(form, mapped)
                if validate(form, values):
                    skipped += 1
                    continue
                self.add_entry(form, values, actor, source_file=path)
                imported += 1
        self.log(actor, "entries.import_csv", form.key, f"{os.path.basename(path)}: {imported} imported, {skipped} skipped")
        return imported, skipped

    def _to_entry(self, row: sqlite3.Row) -> dict:
        return {
            "id": row["id"],
            "form_key": row["form_key"],
            "values": json.loads(self.cipher.decrypt(row["data"])),
            "source_file": self.cipher.decrypt(row["source_file"]),
            "entered_at": row["created_at"],
            "entered_by": row["created_by"],
        }

    # --- uploads awaiting review -----------------------------------------

    def _upload_path(self, upload_id: str) -> str:
        return os.path.join(self.uploads_dir, f"{uuid.UUID(upload_id)}.bin")

    def create_upload(self, filename: str, data: bytes, form_key: str, actor: str) -> str:
        """Store an uploaded image encrypted on disk until it's reviewed."""
        upload_id = str(uuid.uuid4())
        path = self._upload_path(upload_id)
        with open(os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), "wb") as f:
            f.write(self.cipher.encrypt_bytes(data))
        self._execute(
            "INSERT INTO uploads (id, form_key, filename, status, uploaded_by, created_at) "
            "VALUES (?, ?, ?, 'reading', ?, ?)",
            (upload_id, form_key, self.cipher.encrypt(os.path.basename(filename)), actor, now()),
        )
        self.log(actor, "upload.create", upload_id, form_key)
        return upload_id

    def upload_image(self, upload_id: str) -> bytes:
        with open(self._upload_path(upload_id), "rb") as f:
            return self.cipher.decrypt_bytes(f.read())

    def set_extraction(self, upload_id: str, extraction: dict | None, error: str | None = None) -> None:
        """Store {"values": {...}, "uncertain": [...]} from the AI, or the error if reading failed."""
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
            "form_key": row["form_key"],
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

    def export_xlsx(self, path: str, form: FormTemplate, entries: list[dict] | None = None) -> int:
        """Write entries (default: all of this form type) to a formatted workbook. Returns rows written.

        The workbook is NOT encrypted - it's for handing to people who need it.
        """
        entries = self.entries(form.key) if entries is None else entries
        wb = Workbook()
        ws = wb.active
        ws.title = form.name[:31]
        ws.append([f.label for f in form.fields] + ["Source file", "Entered at", "Entered by"])
        for entry in entries:
            row = []
            for field in form.fields:
                value = entry["values"].get(field.key)
                if value and field.type == "date":
                    try:
                        value = date.fromisoformat(value)
                    except ValueError:
                        pass
                elif value and field.type == "number":
                    try:
                        value = float(value)
                    except ValueError:
                        pass
                row.append(value)
            ws.append(row + [entry["source_file"], entry["entered_at"], entry["entered_by"]])

        for cell in ws[1]:
            cell.font = Font(bold=True)
        widths = [COLUMN_WIDTH[f.type] for f in form.fields] + [24, 20, 14]
        for index, width in enumerate(widths, start=1):
            letter = ws.cell(row=1, column=index).column_letter
            ws.column_dimensions[letter].width = width
            field = form.fields[index - 1] if index <= len(form.fields) else None
            for cell in ws[letter][1:]:
                if field and field.type == "date":
                    cell.number_format = "yyyy-mm-dd"
                elif field and field.type == "id":
                    cell.number_format = "@"  # keep leading zeros
        ws.freeze_panes = "A2"
        last_col = ws.cell(row=1, column=len(widths)).column_letter
        last_row = max(len(entries) + 1, 2)
        ws.auto_filter.ref = f"A1:{last_col}{last_row}"
        if form.flag_field:
            flag_col = ws.cell(row=1, column=form.keys.index(form.flag_field) + 1).column_letter
            for value in form.flag_values:
                quoted = value.replace('"', '""')
                ws.conditional_formatting.add(
                    f"A2:{last_col}{last_row}",
                    FormulaRule(formula=[f'${flag_col}2="{quoted}"'], fill=PatternFill("solid", fgColor=FLAG_FILL)),
                )
        atomic_write(path, wb.save)
        return len(entries)
