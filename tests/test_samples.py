import csv
import os
from datetime import date

import pytest
from PIL import Image

from drugtest_import import evaluate, samples
from drugtest_import.forms import BUILT_IN_FORMS, DRUG_TEST, FieldDef, FormTemplate, normalize, validate


@pytest.mark.parametrize("form", BUILT_IN_FORMS, ids=lambda f: f.key)
def test_fake_records_are_valid_and_vary(form):
    records = samples.fake_records(form, 40, seed=3, today=date(2026, 1, 15))
    for record in records:
        expected = {k: v[1] or None for k, v in record.items()}
        assert validate(form, expected, today=date(2026, 1, 15)) == {}
        # What's written normalizes to the expected answer (dates in mixed styles, etc.)
        written = normalize(form, {k: v[0] for k, v in record.items()})
        assert written == normalize(form, expected)
    assert len({r[form.keys[0]] for r in records}) > 3
    assert samples.fake_records(form, 5, seed=3) == samples.fake_records(form, 5, seed=3)


def test_make_writes_labels_printable_and_synthetic_photos(tmp_path):
    folder = tmp_path / "practice"
    samples.make(str(folder), DRUG_TEST, 3, synthetic=True)
    assert sorted(os.listdir(folder)) == ["form-01.jpg", "form-02.jpg", "form-03.jpg", "labels.csv", "print-forms.html"]
    rows = list(csv.DictReader(open(folder / "labels.csv")))
    assert [r["file"] for r in rows] == ["form-01.jpg", "form-02.jpg", "form-03.jpg"]
    page = (folder / "print-forms.html").read_text()
    assert page.count('class="form"') == 3 and "What to write on each form" in page
    with Image.open(folder / "form-01.jpg") as img:
        assert img.format == "JPEG" and img.size[1] > 1500
    # The output plugs straight into the accuracy tool
    assert len(evaluate.load_labels(str(folder), DRUG_TEST)) == 3
    with pytest.raises(ValueError, match="already exists"):
        samples.make(str(folder), DRUG_TEST, 3, synthetic=False)


def test_rename_orders_photos_and_keeps_extension(tmp_path):
    samples.make(str(tmp_path), DRUG_TEST, 3, synthetic=False)
    for i, name in enumerate(["IMG_9.JPG", "IMG_10.png", "IMG_2.jpg"]):
        Image.new("RGB", (10, 10)).save(tmp_path / name)
        os.utime(tmp_path / name, (1000 + i, 1000 + i))  # taken in this order
    pairs = samples.rename(str(tmp_path))
    assert pairs == [("IMG_9.JPG", "form-01.jpg"), ("IMG_10.png", "form-02.png"), ("IMG_2.jpg", "form-03.jpg")]
    rows = list(csv.DictReader(open(tmp_path / "labels.csv")))
    assert [r["file"] for r in rows] == ["form-01.jpg", "form-02.png", "form-03.jpg"]
    assert all((tmp_path / r["file"]).exists() for r in rows)


def test_rename_needs_one_photo_per_form(tmp_path):
    samples.make(str(tmp_path), DRUG_TEST, 3, synthetic=False)
    with pytest.raises(ValueError, match="No new photos"):
        samples.rename(str(tmp_path))
    Image.new("RGB", (10, 10)).save(tmp_path / "IMG_1.jpg")
    with pytest.raises(ValueError, match="Found 1 photos but labels.csv lists 3"):
        samples.rename(str(tmp_path))


def test_custom_field_the_generator_cant_fill(tmp_path):
    form = FormTemplate(key="odd", name="Odd", fields=[FieldDef(key="code", label="Code", type="id", required=True),
                                                       FieldDef(key="when", label="When", type="date")])
    assert samples.fake_records(form, 3)  # ids and dates are fine
    samples.make(str(tmp_path / "x"), form, 2, synthetic=True)
