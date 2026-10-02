import csv
from datetime import date
from types import SimpleNamespace

import pytest

from drugtest_import import extract
from drugtest_import.schema import Extraction, normalize, normalize_date, validate
from drugtest_import.storage import append_record

GOOD = {"EmployeeID": "emp005", "Name": " Bob Wilson ", "Department": "Sales", "TestDate": "9/18/2025",
        "TestType": "urine", "Result": "NEGATIVE", "Notes": "Pre-employment screen"}


@pytest.mark.parametrize("raw,expected", [
    ("9/14/2025", "2025-09-14"), ("09-14-25", "2025-09-14"), ("2025-09-14", "2025-09-14"),
    ("Sept 14, 2025", None), ("Sep 14, 2025", "2025-09-14"), ("13/40/2025", None), ("garbage", None),
])
def test_normalize_date(raw, expected):
    assert normalize_date(raw) == expected


def test_normalize_canonicalises_fields():
    record = normalize(GOOD)
    assert record == {"EmployeeID": "EMP005", "Name": "Bob Wilson", "Department": "Sales", "TestDate": "2025-09-18",
                      "TestType": "Urine", "Result": "Negative", "Notes": "Pre-employment screen"}
    assert validate(record, today=date(2025, 10, 1)) == {}


def test_validate_flags_problems():
    record = normalize({**GOOD, "Name": "", "TestDate": "soon", "Result": "Negatve", "TestType": "Sweat"})
    problems = validate(record, today=date(2025, 10, 1))
    assert set(problems) == {"Name", "TestDate", "Result", "TestType"}


def test_validate_rejects_future_date():
    assert "TestDate" in validate(normalize(GOOD), today=date(2025, 9, 1))


def test_append_record_keeps_leading_zeros(tmp_path):
    path = tmp_path / "log.csv"
    append_record(normalize({**GOOD, "EmployeeID": "007123"}), str(path))
    append_record(normalize(GOOD), str(path))
    with open(path, newline="") as f:
        rows = list(csv.DictReader(f))
    assert [r["EmployeeID"] for r in rows] == ["007123", "EMP005"]
    assert list(rows[0]) == ["EmployeeID", "Name", "Department", "TestDate", "TestType", "Result", "Notes"]


class FakeClient:
    def __init__(self, response):
        self.calls = []
        self.beta = SimpleNamespace(messages=SimpleNamespace(parse=self._parse))
        self.response = response

    def _parse(self, **kwargs):
        self.calls.append(kwargs)
        return self.response


@pytest.fixture
def image(tmp_path):
    from PIL import Image
    path = tmp_path / "note.png"
    Image.new("RGB", (3000, 2000), "yellow").save(path)
    return str(path)


def test_extract_fields_returns_parsed_output(image):
    parsed = Extraction(Name="Bob Wilson", Result="Negative", uncertain_fields=["Result"])
    client = FakeClient(SimpleNamespace(stop_reason="end_turn", parsed_output=parsed))
    assert extract.extract_fields(image, client=client) is parsed
    content = client.calls[0]["messages"][0]["content"]
    assert content[0]["source"]["media_type"] == "image/jpeg"


@pytest.mark.parametrize("stop_reason", ["refusal", "max_tokens"])
def test_extract_fields_raises_on_bad_stop(image, stop_reason):
    client = FakeClient(SimpleNamespace(stop_reason=stop_reason, parsed_output=None))
    with pytest.raises(extract.ExtractionError):
        extract.extract_fields(image, client=client)
