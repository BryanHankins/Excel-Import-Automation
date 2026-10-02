# Drug Test Import

Turn photos of handwritten drug-test notes into clean, searchable records.

1. Upload one or more photos of notes.
2. Claude's vision model reads the handwriting and fills in the fields. Fields it found hard to read are flagged.
3. A reviewer checks each field against the photo, fixes anything wrong, and saves or skips. Nothing is stored until a person confirms.
4. Search, filter and export records to a formatted Excel workbook.

There are two ways to run it, sharing the same database format:

- **Web app** (multi-user, for an office or clinic): sign-in, roles, audit log.
- **Desktop app** (single user on one PC).

## Setup

Requires Python 3.10+.

```
pip install -r requirements.txt
python -m drugtest_import.manage gen-key        # prints a key - store it in a password manager
export DRUGTEST_KEY=<that key>                  # Windows: set DRUGTEST_KEY=<that key>
export ANTHROPIC_API_KEY=sk-ant-...
```

**Keep `DRUGTEST_KEY` safe and backed up separately from the database.** All records are encrypted with it; without it they can't be recovered.

### Web app

```
python -m drugtest_import.manage create-user maria --role admin
python -m drugtest_import.web                   # http://127.0.0.1:8000
```

Roles:

| Role | Can |
|---|---|
| viewer | Search and view records |
| reviewer | + upload photos, review/save records, export to Excel |
| admin | + add/disable users, change roles, reset passwords, view the audit log |

### Desktop app

```
python -m drugtest_import.app
```

Requires Tk (included with the python.org installers). Records from the old `DrugTestingOrganizer.csv` are imported automatically on first run.

### Other admin commands

```
python -m drugtest_import.manage set-password maria
python -m drugtest_import.manage import-csv old-records.csv
```

## Configuration

| Variable | Default | Purpose |
|---|---|---|
| `DRUGTEST_KEY` | — (required) | Encryption key from `manage gen-key` |
| `ANTHROPIC_API_KEY` | — (required) | Used to read photos |
| `DRUGTEST_DB` | `drugtest.db` | Database file |
| `DRUGTEST_UPLOADS` | `uploads/` next to the database | Encrypted photos awaiting review |
| `DRUGTEST_MODEL` | `claude-opus-5-5` | Claude model used to read images |
| `DRUGTEST_HOST` / `DRUGTEST_PORT` | `127.0.0.1` / `8000` | Web server address |
| `DRUGTEST_INSECURE_COOKIES` | unset | Set to `1` only to test over plain HTTP on a non-localhost address |
| `DRUGTEST_CSV` | `DrugTestingOrganizer.csv` | Old CSV log the desktop app imports on first run |

## Security

- **Encryption at rest:** names, employee IDs, results, notes and photo file names are encrypted in the database (Fernet: AES-128-CBC + HMAC-SHA256). Uploaded photos are stored encrypted and deleted once saved or skipped. Exact-match lookups use keyed hashes, so duplicates can be found without decrypting.
- **Sign-in:** passwords are hashed with scrypt and must be at least 12 characters. An account locks for 15 minutes after 5 failed attempts. Sessions end after 30 minutes idle or 8 hours total.
- **Web hardening:** CSRF tokens on every form, `SameSite=Strict` secure cookies, a strict Content-Security-Policy with no JavaScript, and `no-store` caching.
- **Audit log:** sign-ins (including failures), uploads, saves, skips, record views, exports and user changes are logged with who and when. The log is append-only in the database.
- **Not encrypted:** Excel exports are plain workbooks for handing to people who need them; every export is logged.
- **Third party:** photos are sent to the Anthropic API to be read.

### Deploying the web app

- Serve it over HTTPS only, behind a reverse proxy such as Caddy or nginx; cookies are marked secure. For example, the Caddyfile `records.example.com { reverse_proxy 127.0.0.1:8000 }` is enough.
- Run one instance per customer, with its own database and key.
- Back up the database file and the `uploads/` folder. Store the key separately.

## Development

```
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

Sample data lives in `examples/`.
