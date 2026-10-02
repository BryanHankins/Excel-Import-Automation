import io
import json
import sqlite3
from datetime import date
from types import SimpleNamespace

import anthropic
import pytest
from openpyxl import load_workbook
from PIL import Image
from pydantic import ValidationError

from drugtest_import import extract
from drugtest_import.crypto import Cipher, ConfigError, generate_key, hash_password, verify_password
from drugtest_import.forms import (DRUG_TEST, EQUIPMENT_INSPECTION, SAFETY_INCIDENT, TRAINING_SIGNOFF, FieldDef,
                                   FormTemplate, extraction_model, extraction_values, make_key, normalize,
                                   normalize_date, validate)
from drugtest_import.storage import Database, FormInUse

try:
    import httpx2 as httpx
except ImportError:  # older SDKs use httpx
    import httpx

KEY = generate_key()
GOOD = {"employee_id": "emp005", "name": " Bob  Wilson ", "department": "Sales", "test_date": "9/18/2025",
        "test_type": "urine", "result": "NEGATIVE", "notes": "Pre-employment screen"}


# --- form definitions --------------------------------------------------------

@pytest.mark.parametrize("raw,expected", [
    ("9/14/2025", "2025-09-14"), ("09-14-25", "2025-09-14"), ("2025-09-14", "2025-09-14"),
    ("Sept 14, 2025", None), ("Sep 14, 2025", "2025-09-14"), ("13/40/2025", None), ("garbage", None),
])
def test_normalize_date(raw, expected):
    assert normalize_date(raw) == expected


def test_normalize_canonicalises_fields():
    record = normalize(DRUG_TEST, {**GOOD, "unknown": "dropped"})
    assert record == {"employee_id": "EMP005", "name": "Bob Wilson", "department": "Sales", "test_date": "2025-09-18",
                      "test_type": "Urine", "result": "Negative", "notes": "Pre-employment screen"}
    assert validate(DRUG_TEST, record, today=date(2025, 10, 1)) == {}


def test_validate_flags_problems():
    record = normalize(DRUG_TEST, {**GOOD, "name": "", "test_date": "soon", "result": "Negatve",
                                   "test_type": "Sweat", "employee_id": "EMP 5"})
    problems = validate(DRUG_TEST, record, today=date(2025, 10, 1))
    assert set(problems) == {"name", "test_date", "result", "test_type", "employee_id"}
    assert "future" in validate(DRUG_TEST, normalize(DRUG_TEST, GOOD), today=date(2025, 9, 1))["test_date"]


def test_number_and_length_validation():
    record = normalize(TRAINING_SIGNOFF, {"employee_name": "A", "course": "Forklift", "completed_on": "2025-01-02",
                                          "outcome": "passed", "hours": "1,234.50"})
    assert record["hours"] == "1234.5" and record["outcome"] == "Passed"
    assert validate(TRAINING_SIGNOFF, {**record, "hours": "lots"})["hours"] == "Must be a number"
    assert "Too long" in validate(TRAINING_SIGNOFF, {**record, "course": "x" * 201})["course"]


@pytest.mark.parametrize("form", [DRUG_TEST, SAFETY_INCIDENT, TRAINING_SIGNOFF, EQUIPMENT_INSPECTION])
def test_built_in_forms_are_consistent(form):
    assert FormTemplate.model_validate_json(form.model_dump_json()) == form
    assert form.flag_values and form.duplicate_keys and form.date_field


def test_form_template_validation():
    with pytest.raises(ValidationError, match="unique"):
        FormTemplate(key="x", name="X", fields=[FieldDef(key="a", label="A"), FieldDef(key="a", label="B")])
    with pytest.raises(ValidationError, match="choice"):
        FieldDef(key="a", label="A", type="choice")
    with pytest.raises(ValidationError, match="date field"):
        FormTemplate(key="x", name="X", fields=[FieldDef(key="a", label="A")], date_field="a")
    with pytest.raises(ValidationError, match="reserved"):
        FieldDef(key="uncertain_fields", label="A")
    form = FormTemplate(key="x", name="X", fields=[FieldDef(key="s", label="S", type="choice", choices=["Ok", "Bad", " ok "])],
                        flag_field="s", flag_values=["bad", "missing"])
    assert form.fields[0].choices == ["Ok", "Bad"] and form.flag_values == ["Bad"]


def test_make_key():
    assert make_key("Employee ID") == "employee_id"
    assert make_key("Employee ID", {"employee_id"}) == "employee_id_2"
    assert make_key("2nd shift?") == "f_2nd_shift"
    assert make_key("!!!") == "field"
    assert make_key("Uncertain fields") == "uncertain_fields_value"


def test_extraction_model_uses_field_keys_even_when_they_clash_with_pydantic():
    form = FormTemplate(key="odd", name="Odd", fields=[FieldDef(key="model_config", label="Config"),
                                                       FieldDef(key="json", label="Json")])
    model = extraction_model(form)
    schema = model.model_json_schema()
    assert set(schema["properties"]) == {"model_config", "json", "uncertain_fields"}
    parsed = model.model_validate({"model_config": "a", "json": None, "uncertain_fields": ["json"]})
    assert extraction_values(form, parsed) == ({"model_config": "a", "json": None}, ["json"])


# --- extraction ----------------------------------------------------------------

def png_file(size=(3000, 2000)):
    buffer = io.BytesIO()
    Image.new("RGB", size, "yellow").save(buffer, format="PNG")
    buffer.seek(0)
    return buffer


def test_extract_fields_through_real_sdk():
    """Runs the real Anthropic SDK against a fake HTTP server to check the request and parsing."""
    sent = {}

    def handler(request):
        sent.update(json.loads(request.content))
        answer = {**GOOD, "department": None, "uncertain_fields": ["result"]}
        return httpx.Response(200, json={
            "id": "msg_1", "type": "message", "role": "assistant", "model": sent["model"],
            "content": [{"type": "text", "text": json.dumps(answer)}], "stop_reason": "end_turn",
            "stop_sequence": None, "usage": {"input_tokens": 1, "output_tokens": 1},
        })

    client = anthropic.Anthropic(api_key="test", http_client=anthropic.DefaultHttpxClient(transport=httpx.MockTransport(handler)))
    values, uncertain = extract.extract_fields(png_file(), DRUG_TEST, client=client)
    assert values["name"] == " Bob  Wilson " and values["department"] is None and uncertain == ["result"]
    content = sent["messages"][0]["content"]
    assert content[0]["source"]["media_type"] == "image/jpeg"
    assert "test_type: Test type. Specimen type. One of: Urine" in content[1]["text"]
    schema = sent["output_config"]["format"]["schema"]
    assert set(schema["properties"]) == set(DRUG_TEST.keys) | {"uncertain_fields"}
    assert sent["fallbacks"] == "default"


class FakeClient:
    def __init__(self, response):
        self.beta = SimpleNamespace(messages=SimpleNamespace(parse=lambda **kwargs: response))


@pytest.mark.parametrize("stop_reason", ["refusal", "max_tokens"])
def test_extract_fields_raises_on_bad_stop(stop_reason):
    client = FakeClient(SimpleNamespace(stop_reason=stop_reason, parsed_output=None))
    with pytest.raises(extract.ExtractionError):
        extract.extract_fields(png_file((10, 10)), DRUG_TEST, client=client)


# --- storage ---------------------------------------------------------------

@pytest.fixture
def store(tmp_path):
    store = Database(str(tmp_path / "records.db"), Cipher(KEY))
    yield store
    store.close()


def add(store, form=DRUG_TEST, **overrides):
    return store.add_entry(form, normalize(form, {**GOOD, **overrides}), "tester")


def test_built_in_forms_seeded(store):
    assert {f.key for f in store.forms()} == {"drug-test", "safety-incident", "training-signoff", "equipment-inspection"}


def test_entry_round_trip(store):
    store.add_entry(DRUG_TEST, normalize(DRUG_TEST, {**GOOD, "employee_id": "007123"}), "tester",
                    source_file="/private/photos/note1.jpg")
    [entry] = store.entries("drug-test")
    assert entry["values"]["employee_id"] == "007123"
    assert entry["source_file"] == "note1.jpg"  # full path not stored
    assert store.count() == 1 and store.count("drug-test") == 1 and store.count("safety-incident") == 0


def test_find_duplicate(store):
    record = normalize(DRUG_TEST, GOOD)
    assert store.find_duplicate(DRUG_TEST, record) is None
    add(store)
    assert store.find_duplicate(DRUG_TEST, {**record, "employee_id": "OTHER", "result": "Positive"})
    assert store.find_duplicate(DRUG_TEST, {**record, "test_type": "Hair"}) is None
    form = FormTemplate.model_validate({**DRUG_TEST.model_dump(), "fields": [
        {**f.model_dump(), "duplicate_check": f.key in ("employee_id", "test_date")} for f in DRUG_TEST.fields]})
    store.save_form(form, "tester")  # changing duplicate fields re-indexes existing entries
    assert store.find_duplicate(form, {**record, "name": "Someone Else"})
    assert store.find_duplicate(form, {**record, "employee_id": None}) is None


def test_saving_forms_protects_existing_data(store):
    add(store)
    without_notes = DRUG_TEST.model_copy(update={"fields": DRUG_TEST.fields[:-1]})
    with pytest.raises(FormInUse, match="Notes"):
        store.save_form(without_notes, "tester")
    retyped = DRUG_TEST.model_copy(update={"fields": DRUG_TEST.fields[:-1] + [FieldDef(key="notes", label="Notes", type="number")]})
    with pytest.raises(FormInUse, match="type"):
        store.save_form(retyped, "tester")
    relabelled = DRUG_TEST.model_copy(update={"fields": DRUG_TEST.fields[:-1] + [FieldDef(key="notes", label="Comments", type="longtext")]})
    store.save_form(relabelled, "tester")
    assert store.get_form("drug-test").field("notes").label == "Comments"
    with pytest.raises(FormInUse):
        store.delete_form("drug-test", "tester")
    store.delete_form("safety-incident", "tester")
    assert store.get_form("safety-incident") is None


def test_import_csv(store, tmp_path):
    path = tmp_path / "old.csv"
    path.write_text("EmployeeID,Name,Department,TestDate,TestType,Result,Notes\n"
                    ",John Doe,HR,9/15/2025,urine,negative,ok\n"
                    "EMP9,,IT,2025-09-15,Blood,Pending,missing name\n")
    assert store.import_csv(str(path), "tester") == (1, 1)
    [entry] = store.entries("drug-test")
    assert entry["values"]["test_date"] == "2025-09-15" and entry["values"]["test_type"] == "Urine"

    path.write_text("Equipment ID,Inspection date,Inspector,Status\nfl-3,2025-03-01,Ana,needs repair\n")
    assert store.import_csv(str(path), "tester", EQUIPMENT_INSPECTION) == (1, 0)
    assert store.entries("equipment-inspection")[0]["values"]["status"] == "Needs repair"


def test_export_xlsx(store, tmp_path):
    add(store, employee_id="007123")
    add(store, name="Ann Lee", result="Positive")
    path = tmp_path / "out.xlsx"
    assert store.export_xlsx(str(path), DRUG_TEST) == 2
    ws = load_workbook(path).active
    assert [c.value for c in ws[1]][:7] == [f.label for f in DRUG_TEST.fields]
    assert ws["A3"].value == "007123"  # newest first
    assert ws["D2"].value.date() == date(2025, 9, 18)
    assert ws["F2"].value == "Positive"
    assert ws.freeze_panes == "A2" and ws.title == "Drug test"


def test_sensitive_fields_encrypted_at_rest(store, tmp_path):
    add(store, employee_id="ZQ7781")
    store.create_upload("bob_wilson_note.png", b"image-bytes", "drug-test", "tester")
    raw = open(store.path, "rb").read()
    for secret in (b"Bob Wilson", b"ZQ7781", b"Pre-employment", b"bob_wilson"):
        assert secret not in raw
    # Choice values like "Negative" also appear in form definitions, so check the data rows themselves
    rows = store.conn.execute("SELECT * FROM entries UNION ALL SELECT id, form_key, filename, status, "
                              "extraction, error, uploaded_by, created_at FROM uploads").fetchall()
    stored = b"".join(v if isinstance(v, bytes) else str(v).encode() for row in rows for v in row)
    assert b"Negative" not in stored and b"Urine" not in stored
    [upload_file] = (tmp_path / "uploads").iterdir()
    assert b"image-bytes" not in upload_file.read_bytes()


def test_migrates_drug_test_tables_from_earlier_version(tmp_path):
    cipher = Cipher(KEY)
    path = str(tmp_path / "v1.db")
    conn = sqlite3.connect(path)
    conn.executescript("""
        CREATE TABLE meta (key TEXT PRIMARY KEY, value BLOB);
        CREATE TABLE records (id INTEGER PRIMARY KEY, employee_id BLOB, employee_index TEXT, name BLOB NOT NULL,
            name_index TEXT NOT NULL, department TEXT, test_date TEXT NOT NULL, test_type TEXT NOT NULL,
            result BLOB NOT NULL, notes BLOB, source_file BLOB, created_at TEXT NOT NULL, created_by TEXT NOT NULL);
        CREATE TABLE uploads (id TEXT PRIMARY KEY, filename BLOB NOT NULL, status TEXT NOT NULL, extraction BLOB,
            error TEXT, uploaded_by TEXT NOT NULL, created_at TEXT NOT NULL);
    """)
    conn.execute("INSERT INTO meta VALUES ('key_check', ?)", (cipher.encrypt("ok"),))
    conn.execute(
        "INSERT INTO records (employee_id, name, name_index, department, test_date, test_type, result, notes, "
        "source_file, created_at, created_by) VALUES (?, ?, 'x', 'HR', '2025-09-15', 'Urine', ?, NULL, ?, "
        "'2025-09-16T00:00:00+00:00', 'maria')",
        (cipher.encrypt("EMP1"), cipher.encrypt("John Doe"), cipher.encrypt("Negative"), cipher.encrypt("a.jpg")))
    conn.execute("INSERT INTO uploads VALUES ('6f0b2c58-4a8e-4e8e-9d6f-0f6a4a0d2b11', ?, 'ready', NULL, NULL, 'maria', 'now')",
                 (cipher.encrypt("note.png"),))
    conn.commit()
    conn.close()

    db = Database(path, cipher)
    [entry] = db.entries("drug-test")
    assert entry["values"] == {"employee_id": "EMP1", "name": "John Doe", "department": "HR", "test_date": "2025-09-15",
                               "test_type": "Urine", "result": "Negative", "notes": None}
    assert (entry["source_file"], entry["entered_by"]) == ("a.jpg", "maria")
    assert db.find_duplicate(DRUG_TEST, entry["values"])
    [upload] = db.pending_uploads()
    assert (upload["form_key"], upload["status"], upload["filename"]) == ("drug-test", "reading", "note.png")
    db.close()
    Database(path, cipher).close()  # re-opening a migrated database is a no-op


def test_wrong_key_rejected(store):
    store.close()
    with pytest.raises(ConfigError):
        Database(store.path, Cipher(generate_key()))


def test_audit_log_is_append_only(store):
    store.log("tester", "something")
    with pytest.raises(sqlite3.DatabaseError):
        store.conn.execute("DELETE FROM audit_log")
    with pytest.raises(sqlite3.DatabaseError):
        store.conn.execute("UPDATE audit_log SET actor = 'x'")


def test_upload_lifecycle(store):
    upload_id = store.create_upload("/x/note.png", b"data", "drug-test", "tester")
    assert store.get_upload(upload_id)["status"] == "reading"
    store.set_extraction(upload_id, {"values": {"name": "Bob"}, "uncertain": []})
    upload = store.get_upload(upload_id)
    assert (upload["status"], upload["extraction"]["values"], upload["filename"]) == ("ready", {"name": "Bob"}, "note.png")
    assert store.upload_image(upload_id) == b"data"
    store.delete_upload(upload_id)
    assert store.get_upload(upload_id) is None and store.pending_uploads() == []


def test_users(store):
    with pytest.raises(ValueError):
        store.create_user("amy", "short", "viewer", "tester")
    user_id = store.create_user("amy", "long-enough-pw", "viewer", "tester")
    with pytest.raises(ValueError):
        store.create_user("AMY", "long-enough-pw", "viewer", "tester")
    store.update_user(user_id, "tester", role="admin", active=False)
    user = store.get_user_by_name("Amy")
    assert (user["role"], user["active"]) == ("admin", 0)


def test_password_hashing():
    stored = hash_password("correct horse battery")
    assert verify_password("correct horse battery", stored)
    assert not verify_password("wrong", stored)
    assert not verify_password("x", "garbage")
