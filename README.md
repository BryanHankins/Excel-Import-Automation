# Drug Test Import

Turn photos of handwritten drug-test notes into clean spreadsheet records.

1. Open one or more photos of notes. Up to three are read at a time in the background.
2. Claude's vision model reads the handwriting and fills in the form. Fields it found hard to read are flagged.
3. You check each field against the photo, fix anything wrong, then **Save** or **Skip** to move to the next photo. Nothing is written until you confirm.

Records are stored in a local SQLite database (`drugtest.db`). **Export to Excel…** writes a formatted `.xlsx` with real dates, filters, a frozen header row, and Positive / Pending / Inconclusive / Refused results highlighted.

Saving warns you if the same person already has a test of the same type on the same date.

Upgrading from the CSV version: on first run, records in `DrugTestingOrganizer.csv` are imported automatically (incomplete rows are skipped and counted).

## Setup

Requires Python 3.10+ with Tk (included with the python.org installers).

```
pip install -r requirements.txt
export ANTHROPIC_API_KEY=sk-ant-...      # Windows: set ANTHROPIC_API_KEY=sk-ant-...
python -m drugtest_import.app
```

Optional settings:

| Variable | Default | Purpose |
|---|---|---|
| `DRUGTEST_DB` | `drugtest.db` | Record database |
| `DRUGTEST_CSV` | `DrugTestingOrganizer.csv` | Old CSV log imported on first run |
| `DRUGTEST_MODEL` | `claude-opus-5-5` | Claude model used to read images |

## Fields and validation

| Field | Rule |
|---|---|
| EmployeeID | Optional; letters, digits, dashes (leading zeros kept) |
| Name | Required |
| Department | Optional |
| TestDate | Required; stored as `YYYY-MM-DD`; can't be in the future. `9/14/2025`, `09-14-25` etc. are converted |
| TestType | Required; Urine, Blood, Hair, Saliva or Breath |
| Result | Required; Negative, Positive, Pending, Inconclusive or Refused |
| Notes | Optional |

## Privacy

- Images are sent to the Anthropic API for reading; nothing else leaves the machine. No temporary image files are written.
- The database and exported workbooks are unencrypted. Store them somewhere access-controlled. Only the photo's file name is recorded, not its full path. `.gitignore` excludes databases, spreadsheets and images so real records aren't committed.

## Development

```
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

Sample data lives in `examples/`.
