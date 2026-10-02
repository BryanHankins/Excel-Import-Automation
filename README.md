# Drug Test Import

Turn photos of handwritten drug-test notes into clean spreadsheet records.

1. Open a photo of the note.
2. Claude's vision model reads the handwriting and fills in the form. Fields it found hard to read are flagged.
3. You check each field against the photo, fix anything wrong, and save. Nothing is written until you confirm.

Records are appended to `DrugTestingOrganizer.csv`, which opens in Excel.

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
| `DRUGTEST_CSV` | `DrugTestingOrganizer.csv` | Where records are saved |
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
- The CSV is unencrypted. Store it somewhere access-controlled. `.gitignore` excludes CSVs and images so real records aren't committed.

## Development

```
pip install -r requirements.txt -r requirements-dev.txt
pytest
```

Sample data lives in `examples/`.
