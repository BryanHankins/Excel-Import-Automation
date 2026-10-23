import csv

import pytest
from PIL import Image

from drugtest_import import evaluate
from drugtest_import.extract import ExtractionError, Reading
from drugtest_import.forms import DRUG_TEST


@pytest.fixture
def folder(tmp_path):
    for name in ("a.jpg", "b.png", "c.png"):
        Image.new("RGB", (20, 20)).save(tmp_path / name)
    (tmp_path / "notes.txt").write_text("not an image")
    return tmp_path


def fake_reader(path, form):
    name = path.rsplit("/", 1)[-1]
    if name == "c.png":
        raise ExtractionError("blurry")
    values = {"employee_id": "emp005", "name": "bob  wilson", "department": None, "test_date": "9/18/2025",
              "test_type": "urine", "result": "Negative", "notes": None}
    uncertain = []
    if name == "b.png":
        values |= {"result": "Positive", "test_date": "9/19/2025"}  # two mistakes, one flagged
        uncertain = ["result"]
    return Reading(values, uncertain, model="claude-opus-5-5", input_tokens=2000, output_tokens=500)


def test_init_writes_template_once(folder):
    path = evaluate.write_template(str(folder), DRUG_TEST)
    rows = list(csv.reader(open(path)))
    assert rows[0] == ["file"] + [f.label for f in DRUG_TEST.fields]
    assert [r[0] for r in rows[1:]] == ["a.jpg", "b.png", "c.png"]
    with pytest.raises(ValueError, match="already exists"):
        evaluate.write_template(str(folder), DRUG_TEST)


def label(folder, rows, header=None):
    header = header or ["file", "Employee ID", "Name", "Department", "Test date", "Test type", "Result", "Notes"]
    with open(folder / "labels.csv", "w", newline="") as f:
        csv.writer(f).writerows([header] + rows)


def test_evaluate_scores_fields_and_flags(folder, monkeypatch):
    monkeypatch.setattr(evaluate.extract, "ENGINE", "anthropic")
    good = ["EMP005", "Bob Wilson", "", "2025-09-18", "Urine", "Negative", "?"]
    label(folder, [["a.jpg"] + good, ["b.png"] + good, ["c.png"] + good])
    report = evaluate.evaluate(str(folder), DRUG_TEST, workers=2, reader=fake_reader)

    assert report.documents == 3 and report.failures == {"c.png": "blurry"}
    assert len(report.results) == 12  # 6 checked fields x 2 readable photos (notes skipped)
    wrong = {(r.file, r.key, r.flagged) for r in report.results if not r.correct}
    assert wrong == {("b.png", "result", True), ("b.png", "test_date", False)}
    assert report.cost() == pytest.approx(2 * (2000 * 4 + 500 * 20) / 1_000_000)

    text = evaluate.summary(report)
    assert "Field accuracy:            83.3%" in text
    assert "Photos with every field right: 50.0%" in text
    assert "Silent errors (wrong and NOT flagged): 1" in text
    assert "b.png: Test date - expected '2025-09-18', read '9/19/2025'" in text
    assert "c.png: blurry" in text

    out = folder / "results.csv"
    evaluate.write_results(report, str(out))
    rows = list(csv.DictReader(open(out)))
    assert {r["correct"] for r in rows} == {"yes", "NO"}


def test_blank_expected_catches_invented_values(folder):
    label(folder, [["a.jpg", "", "Bob Wilson", "", "2025-09-18", "Urine", "Negative", ""]])
    report = evaluate.evaluate(str(folder), DRUG_TEST, reader=fake_reader)
    assert [r.key for r in report.results if not r.correct] == ["employee_id"]


def test_labels_validation(folder):
    label(folder, [["a.jpg", "x"]], header=["file", "Shoe size"])
    with pytest.raises(ValueError, match="Shoe size"):
        evaluate.evaluate(str(folder), DRUG_TEST, reader=fake_reader)
    label(folder, [["missing.jpg", "x"]], header=["file", "name"])
    with pytest.raises(ValueError, match="missing.jpg"):
        evaluate.evaluate(str(folder), DRUG_TEST, reader=fake_reader)


def test_cli_init_and_unknown_form(folder, capsys):
    assert evaluate.main(["init", str(folder), "--form", "safety-incident"]) == 0
    assert "Incident date" in (folder / "labels.csv").read_text()
    assert evaluate.main(["init", str(folder), "--form", "nope"]) == 1
    assert "No form type 'nope'" in capsys.readouterr().err


def test_local_engine_reports_free(folder, monkeypatch):
    monkeypatch.setattr(evaluate.extract, "ENGINE", "ollama")
    label(folder, [["a.jpg", "EMP005", "Bob Wilson", "", "2025-09-18", "Urine", "Negative", ""]])
    report = evaluate.evaluate(str(folder), DRUG_TEST, reader=fake_reader)
    assert report.cost() == 0
    assert "Cost: $0 - runs on this computer" in evaluate.summary(report)
