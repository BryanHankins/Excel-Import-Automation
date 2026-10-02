"""Form types: which fields a kind of paper form has, and how to read, check and store them.

A form type (e.g. "Drug test", "Safety incident") is a list of fields. Everything
else - the extraction prompt and schema, the review screen, validation, duplicate
detection and the Excel export - is driven from it.
"""
import re
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from functools import lru_cache
from typing import List, Literal, Optional

from pydantic import BaseModel, Field, create_model, field_validator, model_validator

FIELD_TYPES = {
    "text": "Short text",
    "longtext": "Long text",
    "id": "ID / code",
    "date": "Date",
    "number": "Number",
    "choice": "Choice from a list",
}
KEY_PATTERN = r"[a-z][a-z0-9_]{0,39}"
RESERVED_KEYS = {"uncertain_fields", "id"}
MAX_LENGTH = {"text": 200, "longtext": 2000, "id": 30, "date": 40, "number": 40, "choice": 100}


class FieldDef(BaseModel):
    key: str = Field(pattern=KEY_PATTERN)
    label: str = Field(min_length=1, max_length=60)
    type: Literal["text", "longtext", "id", "date", "number", "choice"] = "text"
    required: bool = False
    choices: List[str] = Field(default_factory=list)
    hint: str = Field("", max_length=300, description="Shown to reviewers and to the AI reading the form")
    duplicate_check: bool = False

    @field_validator("key")
    @classmethod
    def not_reserved(cls, key):
        if key in RESERVED_KEYS:
            raise ValueError(f"'{key}' is reserved")
        return key

    @field_validator("choices")
    @classmethod
    def clean_choices(cls, choices):
        cleaned = []
        for choice in (c.strip() for c in choices):
            if choice and choice.lower() not in {c.lower() for c in cleaned}:
                cleaned.append(choice)
        return cleaned

    @model_validator(mode="after")
    def choices_match_type(self):
        if self.type == "choice" and not self.choices:
            raise ValueError(f"'{self.label}' needs at least one choice")
        if self.type != "choice":
            self.choices = []
        return self


class FormTemplate(BaseModel):
    key: str = Field(pattern=r"[a-z][a-z0-9-]{0,39}")
    name: str = Field(min_length=1, max_length=60)
    description: str = Field("", max_length=500, description="What the document is; helps the AI read it")
    fields: List[FieldDef] = Field(min_length=1, max_length=40)
    date_field: Optional[str] = Field(None, description="Date field used for sorting and date filters")
    flag_field: Optional[str] = Field(None, description="Choice field whose flagged values are highlighted")
    flag_values: List[str] = Field(default_factory=list)

    @model_validator(mode="after")
    def check_references(self):
        keys = [f.key for f in self.fields]
        if len(keys) != len(set(keys)):
            raise ValueError("Field keys must be unique")
        by_key = {f.key: f for f in self.fields}
        if self.date_field and (self.date_field not in by_key or by_key[self.date_field].type != "date"):
            raise ValueError("The sort date must be a date field")
        if self.flag_field:
            if self.flag_field not in by_key or by_key[self.flag_field].type != "choice":
                raise ValueError("The highlighted field must be a choice field")
            choices = by_key[self.flag_field].choices
            self.flag_values = [c for c in choices if c.lower() in {v.lower() for v in self.flag_values}]
        else:
            self.flag_values = []
        return self

    def field(self, key: str) -> FieldDef:
        return next(f for f in self.fields if f.key == key)

    @property
    def keys(self) -> list[str]:
        return [f.key for f in self.fields]

    @property
    def duplicate_keys(self) -> list[str]:
        return [f.key for f in self.fields if f.duplicate_check]


def make_key(label: str, taken: set[str] = frozenset()) -> str:
    """Turn a label like 'Employee ID' into a unique field key like 'employee_id'."""
    base = re.sub(r"[^a-z0-9]+", "_", label.lower()).strip("_")[:36] or "field"
    if not base[0].isalpha():
        base = f"f_{base}"[:36]
    if base in RESERVED_KEYS:
        base = f"{base}_value"
    key, n = base, 2
    while key in taken:
        key, n = f"{base}_{n}", n + 1
    return key


# --- normalising and validating values ----------------------------------

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


def _normalize_value(field: FieldDef, value: str) -> str:
    if field.type == "date":
        return normalize_date(value) or value
    if field.type == "choice":
        return next((c for c in field.choices if c.lower() == value.lower()), value)
    if field.type == "id":
        return value.upper()
    if field.type == "number":
        try:
            return str(Decimal(value.replace(",", "")).normalize())
        except InvalidOperation:
            return value
    return value


def normalize(form: FormTemplate, values: dict) -> dict:
    """Trim whitespace and canonicalise dates, choices, IDs and numbers where possible.

    Values that can't be normalised are left as-is so the reviewer sees them and
    `validate` flags them. Keys not in the form are dropped.
    """
    out = {}
    for field in form.fields:
        value = values.get(field.key)
        value = " ".join(value.split()) if isinstance(value, str) and field.type != "longtext" else value
        value = value.strip() if isinstance(value, str) else None
        out[field.key] = _normalize_value(field, value) if value else None
    return out


def validate(form: FormTemplate, record: dict, today: Optional[date] = None) -> dict:
    """Return {field key: problem} for every field that must be fixed before saving."""
    today = today or date.today()
    problems = {}
    for field in form.fields:
        value = record.get(field.key)
        if not value:
            if field.required:
                problems[field.key] = "Required"
            continue
        if len(value) > MAX_LENGTH[field.type]:
            problems[field.key] = f"Too long (max {MAX_LENGTH[field.type]} characters)"
        elif field.type == "date":
            try:
                if date.fromisoformat(value) > today:
                    problems[field.key] = "Date is in the future"
            except ValueError:
                problems[field.key] = "Use YYYY-MM-DD"
        elif field.type == "choice" and value not in field.choices:
            problems[field.key] = f"Must be one of: {', '.join(field.choices)}"
        elif field.type == "id" and not re.fullmatch(r"[A-Z0-9-]+", value):
            problems[field.key] = "Letters, digits and dashes only"
        elif field.type == "number":
            try:
                Decimal(value)
            except InvalidOperation:
                problems[field.key] = "Must be a number"
    return problems


# --- extraction -------------------------------------------------------------

@lru_cache(maxsize=64)
def _extraction_model(form_json: str) -> type[BaseModel]:
    form = FormTemplate.model_validate_json(form_json)
    # Python attribute names are positional so field keys can never clash with
    # pydantic internals; the JSON uses the field keys via aliases.
    definitions = {
        f"f{i}": (Optional[str], Field(None, alias=field.key, description=_describe(field)))
        for i, field in enumerate(form.fields)
    }
    definitions["uncertain_fields"] = (
        List[Literal[tuple(form.keys)]],
        Field(default_factory=list, description="Keys of fields that were hard to read or ambiguous"),
    )
    return create_model(f"Extraction_{form.key.replace('-', '_')}", **definitions)


def extraction_model(form: FormTemplate) -> type[BaseModel]:
    """Pydantic model the AI must fill in for this form type."""
    return _extraction_model(form.model_dump_json())


def _describe(field: FieldDef) -> str:
    parts = [field.label]
    if field.hint:
        parts.append(field.hint)
    if field.type == "choice":
        parts.append(f"One of: {', '.join(field.choices)}")
    if field.type == "date":
        parts.append("As written on the form")
    return ". ".join(parts)


def extraction_prompt(form: FormTemplate) -> str:
    lines = [f"- {f.key}: {_describe(f)}" for f in form.fields]
    about = f" ({form.description})" if form.description else ""
    return f"""This image is a handwritten or printed form: {form.name}{about}.
Transcribe these fields exactly as written:
{chr(10).join(lines)}

Labels may be abbreviated, misspelled or missing; infer which value belongs to which field from context.

Rules:
- Never guess. If a field is absent or unreadable, return null for it.
- List in uncertain_fields every field you are not confident you read correctly.
- For choice fields, return the listed option the writing clearly matches; otherwise return what is written.
- Do not correct or reformat values beyond fixing obvious letter-shape confusions."""


def extraction_values(form: FormTemplate, result: BaseModel) -> tuple[dict, list[str]]:
    """Split a parsed extraction into ({field key: value}, uncertain field keys)."""
    data = result.model_dump(by_alias=True)
    return {key: data.get(key) for key in form.keys}, list(data.get("uncertain_fields", []))


# --- built-in form types -----------------------------------------------------

def _f(key, label, type="text", required=False, choices=(), hint="", duplicate_check=False):
    return FieldDef(key=key, label=label, type=type, required=required, choices=list(choices), hint=hint,
                    duplicate_check=duplicate_check)


DRUG_TEST = FormTemplate(
    key="drug-test",
    name="Drug test",
    description="A note or form recording an employee drug test. A wrong result is a serious error.",
    fields=[
        _f("employee_id", "Employee ID", "id", hint="e.g. EMP005 or 911911"),
        _f("name", "Name", required=True, hint="Employee full name", duplicate_check=True),
        _f("department", "Department"),
        _f("test_date", "Test date", "date", required=True, duplicate_check=True),
        _f("test_type", "Test type", "choice", required=True, choices=["Urine", "Blood", "Hair", "Saliva", "Breath"],
           hint="Specimen type", duplicate_check=True),
        _f("result", "Result", "choice", required=True,
           choices=["Negative", "Positive", "Pending", "Inconclusive", "Refused"]),
        _f("notes", "Notes", "longtext"),
    ],
    date_field="test_date",
    flag_field="result",
    flag_values=["Positive", "Pending", "Inconclusive", "Refused"],
)

SAFETY_INCIDENT = FormTemplate(
    key="safety-incident",
    name="Safety incident",
    description="A workplace injury, near-miss or incident report.",
    fields=[
        _f("incident_date", "Incident date", "date", required=True, duplicate_check=True),
        _f("location", "Location", required=True, duplicate_check=True),
        _f("reported_by", "Reported by", required=True),
        _f("people_involved", "People involved"),
        _f("incident_type", "Type", "choice", required=True,
           choices=["Injury", "Near miss", "Property damage", "Spill", "Other"], duplicate_check=True),
        _f("severity", "Severity", "choice", required=True, choices=["Minor", "Moderate", "Serious", "Critical"]),
        _f("description", "What happened", "longtext", required=True),
        _f("action_taken", "Action taken", "longtext"),
    ],
    date_field="incident_date",
    flag_field="severity",
    flag_values=["Serious", "Critical"],
)

TRAINING_SIGNOFF = FormTemplate(
    key="training-signoff",
    name="Training sign-off",
    description="A sheet confirming an employee completed a training course.",
    fields=[
        _f("employee_id", "Employee ID", "id"),
        _f("employee_name", "Employee name", required=True, duplicate_check=True),
        _f("course", "Course", required=True, duplicate_check=True),
        _f("completed_on", "Completed on", "date", required=True, duplicate_check=True),
        _f("trainer", "Trainer"),
        _f("outcome", "Outcome", "choice", required=True, choices=["Passed", "Failed", "Incomplete"]),
        _f("hours", "Hours", "number"),
        _f("notes", "Notes", "longtext"),
    ],
    date_field="completed_on",
    flag_field="outcome",
    flag_values=["Failed", "Incomplete"],
)

EQUIPMENT_INSPECTION = FormTemplate(
    key="equipment-inspection",
    name="Equipment inspection",
    description="A checklist or log entry for inspecting a vehicle, machine or piece of safety equipment.",
    fields=[
        _f("equipment_id", "Equipment ID", "id", required=True, hint="Asset tag, unit or serial number",
           duplicate_check=True),
        _f("equipment", "Equipment", hint="e.g. Forklift 3, fire extinguisher"),
        _f("inspection_date", "Inspection date", "date", required=True, duplicate_check=True),
        _f("inspector", "Inspector", required=True),
        _f("status", "Status", "choice", required=True, choices=["Pass", "Needs repair", "Out of service"]),
        _f("issues", "Issues found", "longtext"),
    ],
    date_field="inspection_date",
    flag_field="status",
    flag_values=["Needs repair", "Out of service"],
)

BUILT_IN_FORMS = [DRUG_TEST, SAFETY_INCIDENT, TRAINING_SIGNOFF, EQUIPMENT_INSPECTION]

# Column names used by the CSV files earlier versions wrote
LEGACY_DRUG_TEST_COLUMNS = {
    "EmployeeID": "employee_id", "Name": "name", "Department": "department", "TestDate": "test_date",
    "TestType": "test_type", "Result": "result", "Notes": "notes",
}
