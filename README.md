# Form Import

Turn photos of handwritten or printed paper forms into clean, searchable records.

1. Pick a form type and upload one or more photos.
2. A vision AI model reads each form and fills in that form type's fields. By default it's a free, open-source model running on your own computer, so photos never leave the machine and there are no per-scan fees. Fields it found hard to read are flagged.
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

**1. Install the app**

```
python -m pip install -r requirements.txt
```

**2. Install the free AI model**

Install Ollama from https://ollama.com (Windows, Mac or Linux). Then, in a terminal:

```
ollama pull qwen3-vl:8b                         # one-time download, several GB
python -m drugtest_import.manage check-ai       # confirms everything is ready
```

A computer with 16 GB of RAM is comfortable. A graphics card (NVIDIA, or an Apple Silicon Mac) makes reading much faster; on a CPU alone, expect up to a minute or more per photo. Other models work too. Try a few with the accuracy tool below and set `DRUGTEST_MODEL` to the one that reads your forms best.

**3. Create the encryption key**

```
python -m drugtest_import.manage gen-key        # prints a key - store it in a password manager
```

Set it before running the app:

- Mac/Linux: `export DRUGTEST_KEY=<that key>`
- Windows PowerShell: `$env:DRUGTEST_KEY = "<that key>"`

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

### Fixing and deleting records

Click a record's first column on the **Records** page to open it.

- Reviewers can correct any field. Every change is kept: the record page shows a change history (who, when, old → new), and the audit log notes which fields changed.
- Admins can permanently delete a record, for example one entered twice. They must give a reason; the audit log keeps who, when and why, but not the deleted data.

### Measuring accuracy and cost

Before selling, measure how well the AI reads *your* forms:

```
python -m drugtest_import.evaluate init photos/ --form drug-test   # creates photos/labels.csv
# Fill in labels.csv with what each form actually says (empty = blank on the form, ? = don't check)
python -m drugtest_import.evaluate run photos/ --form drug-test    # reads each photo with the AI
```

The report shows:

- field accuracy, and the share of photos with every field right;
- **silent errors**: wrong values the AI did *not* flag as hard to read. These are the ones a reviewer could miss, so keep them near zero;
- accuracy for each field;
- time per photo (and cost, for the optional paid engine).

Use `--limit 5` for a quick trial and `--model <name>` to compare models (pull each one with `ollama pull` first). Per-field results are written to `photos/accuracy-results.csv`. That file and `labels.csv` contain form contents, so keep them as private as the photos.

#### No real forms yet? Make practice ones

`samples` creates practice forms with made-up people. The answer sheet (`labels.csv`) is filled in for you.

- **Instant check:** draws the forms itself in a handwriting-style font, slightly tilted and blurred like phone photos:
  ```
  python -m drugtest_import.samples make practice --form drug-test --count 20 --synthetic
  python -m drugtest_import.evaluate run practice --form drug-test
  ```
  Typed "handwriting" is much neater than real writing, so this **overstates** accuracy. Use it to check everything works and to compare models.

- **Realistic test:** real handwriting on printed forms:
  ```
  python -m drugtest_import.samples make handwritten --form drug-test --count 20
  ```
  1. Open `handwritten/print-forms.html` in a browser and print it.
  2. Fill in each form by hand from the answer key at the end, ideally with several people's handwriting.
  3. Photograph the forms in order and copy the photos into `handwritten/`.
  4. Run:
     ```
     python -m drugtest_import.samples rename handwritten      # names them form-01, form-02, ... in the order taken
     python -m drugtest_import.evaluate run handwritten --form drug-test
     ```
  These numbers are the ones to quote to customers.

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
| `DRUGTEST_DB` | `drugtest.db` | Database file |
| `DRUGTEST_UPLOADS` | `uploads/` next to the database | Encrypted photos awaiting review |
| `DRUGTEST_ENGINE` | `ollama` | `ollama` (free, local) or `anthropic` (optional paid API; needs `pip install anthropic` and `ANTHROPIC_API_KEY`) |
| `DRUGTEST_MODEL` | `qwen3-vl:8b` | Model used to read images |
| `DRUGTEST_OLLAMA_URL` | `http://localhost:11434` | Where Ollama is running |
| `DRUGTEST_TIMEOUT` | `900` | Seconds to wait for one photo |
| `DRUGTEST_PARALLEL` | `1` (Ollama) | Photos read at the same time |
| `DRUGTEST_HOST` / `DRUGTEST_PORT` | `127.0.0.1` / `8000` | Web server address |
| `DRUGTEST_INSECURE_COOKIES` | unset | Set to `1` only to test over plain HTTP on a non-localhost address |
| `DRUGTEST_CSV` | `DrugTestingOrganizer.csv` | Old CSV log the desktop app imports on first run |

## Security

- **Encryption at rest:** every field value and photo file name is encrypted in the database (Fernet: AES-128-CBC + HMAC-SHA256); only each record's sort date is stored in plain text. Uploaded photos are stored encrypted and deleted once saved or skipped. Exact-match lookups use keyed hashes, so duplicates can be found without decrypting.
- **Sign-in:** passwords are hashed with scrypt and must be at least 12 characters. An account locks for 15 minutes after 5 failed attempts. Sessions end after 30 minutes idle or 8 hours total.
- **Web hardening:** CSRF tokens on every form, `SameSite=Strict` secure cookies, a strict Content-Security-Policy with no JavaScript, and `no-store` caching.
- **Audit log:** sign-ins (including failures), uploads, saves, skips, record views, exports and user changes are logged with who and when. The log is append-only in the database.
- **Not encrypted:** Excel exports are plain workbooks for handing to people who need them; every export is logged.
- **No third party by default:** with the default Ollama engine, photos are read on your own computer and never sent anywhere. Only the optional `anthropic` engine sends photos to an outside API.

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
