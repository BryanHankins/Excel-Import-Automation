"""Record definition and validation rules shared by extraction, review and storage."""
import re
from datetime import date, datetime
from typing import List, Literal, Optional

from pydantic import BaseModel, Field

FIELDS = ["EmployeeID", "Name", "Department", "TestDate", "TestType", "Result", "Notes"]
REQUIRED = ["Name", "TestDate", "TestType", "Result"]
TEST_TYPES = ["Urine", "Blood", "Hair", "Saliva", "Breath"]
RESULTS = ["Negative", "Positive", "Pending", "Inconclusive", "Refused"]

FieldName = Literal["EmployeeID", "Name", "Department", "TestDate", "TestType", "Result", "Notes"]


class Extraction(BaseModel):
    """What the vision model returns for one note. Values are transcribed as written."""

    EmployeeID: Optional[str] = Field(None, description="Employee ID exactly as written, e.g. 'EMP005' or '911911'")
    Name: Optional[str] = Field(None, description="Employee full name")
    Department: Optional[str] = None
    TestDate: Optional[str] = Field(None, description="Test date as written on the note")
    TestType: Optional[str] = Field(None, description="Specimen type, e.g. Urine, Blood, Hair, Saliva, Breath")
    Result: Optional[str] = Field(None, description="Negative, Positive, Pending, Inconclusive or Refused")
    Notes: Optional[str] = None
    uncertain_fields: List[FieldName] = Field(
        default_factory=list,
        description="Fields whose handwriting was hard to read or ambiguous",
    )


DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y", "%m.%d.%Y", "%b %d %Y", "%B %d %Y"]


def normalize_date(value: str) -> Optional[str]:
    """Return an ISO date (YYYY-MM-DD) or None if the value isn't a recognisable date."""
    cleaned = re.sub(r"\s+", " ", value.replace(",", " ")).strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(cleaned, fmt).date().isoformat()
        except ValueError:
            continue
    return None


def _match_choice(value: str, choices: List[str]) -> Optional[str]:
    lowered = value.strip().lower()
    for choice in choices:
        if lowered == choice.lower():
            return choice
    return None


def normalize(record: dict) -> dict:
    """Trim whitespace and canonicalise date, test type and result where possible.

    Values that can't be normalised are left as-is so the reviewer sees them and
    `validate` flags them.
    """
    out = {}
    for key in FIELDS:
        value = record.get(key)
        value = value.strip() if isinstance(value, str) else None
        out[key] = value or None
    if out["TestDate"]:
        out["TestDate"] = normalize_date(out["TestDate"]) or out["TestDate"]
    if out["TestType"]:
        out["TestType"] = _match_choice(out["TestType"], TEST_TYPES) or out["TestType"]
    if out["Result"]:
        out["Result"] = _match_choice(out["Result"], RESULTS) or out["Result"]
    if out["EmployeeID"]:
        out["EmployeeID"] = out["EmployeeID"].upper()
    return out


def validate(record: dict, today: Optional[date] = None) -> dict:
    """Return {field: problem} for every field that must be fixed before saving."""
    today = today or date.today()
    problems = {}
    for key in REQUIRED:
        if not record.get(key):
            problems[key] = "Required"
    test_date = record.get("TestDate")
    if test_date and "TestDate" not in problems:
        try:
            parsed = date.fromisoformat(test_date)
        except ValueError:
            problems["TestDate"] = "Use YYYY-MM-DD"
        else:
            if parsed > today:
                problems["TestDate"] = "Date is in the future"
    if record.get("TestType") and record["TestType"] not in TEST_TYPES:
        problems["TestType"] = f"Must be one of: {', '.join(TEST_TYPES)}"
    if record.get("Result") and record["Result"] not in RESULTS:
        problems["Result"] = f"Must be one of: {', '.join(RESULTS)}"
    employee_id = record.get("EmployeeID")
    if employee_id and not re.fullmatch(r"[A-Z0-9-]{1,20}", employee_id):
        problems["EmployeeID"] = "Letters, digits and dashes only"
    return problems
