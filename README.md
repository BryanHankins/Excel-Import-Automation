# Form Import

Turn photos of handwritten or printed paper forms into clean, searchable records.

1. Pick a form type and upload one or more photos.
2. Claude's vision model reads each form and fills in that form type's fields. Fields it found hard to read are flagged.
3. A reviewer checks each field against the photo, fixes anything wrong, and saves or skips. Nothing is stored until a person confirms.
4. Search, filter and export records to a formatted Excel workbook.

## Form types

Built in: **Drug test**, **Safety incident**, **Training sign-off** and **Equipment inspection**. Admins can edit these or create their own under **Form types** in the web app, with no code changes. A form type has:

- **Fields**, each with:
  - a label;
  - a type: short text, long text, ID/code, date, number, or a choice from a list;
  - optional choices, a hint, and whether the field is required.
- **Duplicate check fields:** if every ticked field matches an existing record, the reviewer is warned before saving.
- **A date field** used to sort and filter records.
- **A highlight rule**, e.g. flag Drug test records whose Result is Positive or Pending.

The AI prompt and answer format, the review screen, validation, the records table and the Excel export all follow the form type.

Once a form type has records, its existing fields can't be removed or change type, so stored data always matches a field. Labels, choices, hints and flags can still change, and new fields can be added.

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
| admin | + manage form types and users, view the audit log |

### Desktop app

```
python -m drugtest_import.app
```

Requires Tk (included with the python.org installers). Choose the form type in the toolbar before opening images. Records from the old `DrugTestingOrganizer.csv` are imported as drug tests on first run.

### Other admin commands

```
python -m drugtest_import.manage set-password maria
python -m drugtest_import.manage list-forms
python -m drugtest_import.manage import-csv old-records.csv --form drug-test
```

CSV headers can be the field labels (e.g. `Inspection date`) or keys (e.g. `inspection_date`). Rows that fail validation are skipped and counted.

Databases created by earlier versions (drug tests only) are upgraded automatically the first time they're opened.

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

- **Encryption at rest:** every field value and photo file name is encrypted in the database (Fernet: AES-128-CBC + HMAC-SHA256); only each record's sort date is stored in plain text. Uploaded photos are stored encrypted and deleted once saved or skipped. Exact-match lookups use keyed hashes, so duplicates can be found without decrypting.
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
