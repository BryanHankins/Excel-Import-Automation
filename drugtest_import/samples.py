"""Make practice forms with made-up people, so you can measure accuracy without real records.

  python -m drugtest_import.samples make FOLDER --form drug-test --count 20 [--synthetic]
    Writes FOLDER/labels.csv (the correct answers) and FOLDER/print-forms.html:
    blank forms to print, plus a sheet saying what to write on each one.
    Fill them in by hand (several people's handwriting is best), photograph each
    one, and save the photos as form-01.jpg, form-02.jpg, ...
    With --synthetic, it also draws the forms itself in a handwriting-style font,
    so you can test right away. Fonts are neater than real handwriting, so
    synthetic results overstate accuracy.

  python -m drugtest_import.samples rename FOLDER
    Renames your photos to form-01.jpg, form-02.jpg, ... in the order they were taken.

Then: python -m drugtest_import.evaluate run FOLDER --form drug-test
"""
import argparse
import csv
import html
import os
import random
import sys
import textwrap
from datetime import date, timedelta

from PIL import Image, ImageDraw, ImageFilter, ImageFont

from .evaluate import IMAGE_EXTENSIONS, load_form
from .forms import FieldDef, FormTemplate, validate

FIRST = ["James", "Maria", "Robert", "Linda", "Michael", "Aisha", "David", "Sofia", "Daniel", "Grace", "Kevin",
         "Priya", "Thomas", "Rosa", "Andre", "Mei", "Patrick", "Fatima", "Luis", "Hannah", "Omar", "Chloe"]
LAST = ["Johnson", "Garcia", "Nguyen", "Okafor", "Smith", "Patel", "Kowalski", "Brown", "Martinez", "O'Neil",
        "Hughes", "Fischer", "Rossi", "Kim", "Baptiste", "Walker", "Silva", "Novak", "Reyes", "Campbell"]
WORDS = {
    "department": ["Warehouse", "HR", "Sales", "Maintenance", "Shipping", "Finance", "IT", "Operations", "Kitchen"],
    "location": ["Loading dock B", "Warehouse aisle 4", "Main office", "Parking lot", "Break room", "Shop floor",
                 "Yard gate 2", "Stairwell east"],
    "equipment": ["Forklift 3", "Fire extinguisher A2", "Pallet jack 7", "Ladder 12", "Truck 21", "Eye wash station"],
    "course": ["Forklift safety", "First aid", "Fire safety", "Lockout tagout", "Ladder safety", "Hazard communication"],
}
NOTES = {
    "drug-test": ["Random selection", "Pre-employment screen", "Post-accident test", "Annual test",
                  "Retest requested", "Employee was cooperative", "Sample sent to lab"],
    "safety-incident": ["Slipped on wet floor, no injury", "Box fell from top shelf", "Cut finger on pallet strap",
                        "Area cordoned off and cleaned", "Supervisor notified, first aid given"],
    "training-signoff": ["Completed on time", "Needs refresher in six months", "Practical test passed"],
    "equipment-inspection": ["Battery cable worn", "Guard rail loose, reported to supervisor", "Tyre pressure low",
                             "Pressure gauge in green", "No issues found"],
}
DEFAULT_NOTES = ["Reviewed with supervisor", "Needs follow-up next week", "See attached sheet"]
DATE_STYLES = ["{d.month}/{d.day}/{d.year}", "{d.month:02}/{d.day:02}/{d.year}", "{d.month}-{d.day}-{d:%y}",
               "{d:%b} {d.day} {d.year}", "{d.year}-{d.month:02}-{d.day:02}"]

PAGE = (1240, 1754)  # A4 at 150 dpi
HANDWRITING_FONTS = ["Inkfree.ttf", "segoepr.ttf", "segoesc.ttf", "Bradley Hand Bold.ttf", "Noteworthy.ttc",
                     "Chalkboard.ttc", "comic.ttf", "DejaVuSans-Oblique.ttf", "DejaVuSans.ttf"]
PRINT_FONTS = ["arial.ttf", "Arial.ttf", "Helvetica.ttc", "DejaVuSans.ttf", "LiberationSans-Regular.ttf"]
FONT_DIRS = [r"C:\Windows\Fonts", os.path.expanduser(r"~\AppData\Local\Microsoft\Windows\Fonts"), "/Library/Fonts",
             "/System/Library/Fonts", "/System/Library/Fonts/Supplemental", "/usr/share/fonts"]


# --- made-up values ---------------------------------------------------------------

def fake_value(field: FieldDef, rng: random.Random, today: date, form_key: str = "") -> tuple[str, str]:
    """Return (what to write on the form, the correct answer for labels.csv)."""
    key = field.key
    if field.type == "date":
        d = today - timedelta(days=rng.randint(1, 300))
        return rng.choice(DATE_STYLES).format(d=d), d.isoformat()
    if field.type == "choice":
        value = rng.choice(field.choices)
        return value, value
    if field.type == "id":
        value = rng.choice([f"EMP{rng.randint(1, 999):03}", f"{rng.randint(100000, 999999)}",
                            f"{rng.choice('ABCFHKT')}{rng.choice('LMRSTX')}-{rng.randint(1, 99)}"])
        return value, value
    if field.type == "number":
        value = str(rng.choice([1, 2, 2.5, 3, 4, 6, 8]))
        return value, value
    if field.type == "longtext":
        value = rng.choice(NOTES.get(form_key, DEFAULT_NOTES))
        return value, value
    for word, choices in WORDS.items():
        if word in key:
            value = rng.choice(choices)
            return value, value
    if "name" in key or key in ("reported_by", "trainer", "inspector", "people_involved"):
        value = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
        return value, value
    value = rng.choice(WORDS["department"])
    return value, value


def fake_records(form: FormTemplate, count: int, seed: int = 1, today: date | None = None) -> list[dict]:
    """[{key: (written, expected)}]; optional fields are left blank about a quarter of the time."""
    rng = random.Random(seed)
    today = today or date.today()
    records = []
    for _ in range(count):
        record = {}
        for field in form.fields:
            if not field.required and rng.random() < 0.25:
                record[field.key] = ("", "")
            else:
                record[field.key] = fake_value(field, rng, today, form.key)
        problems = validate(form, {k: v[1] or None for k, v in record.items()}, today)
        if problems:  # e.g. an admin-defined field this generator can't fill sensibly
            raise ValueError(f"Can't make practice values for: {', '.join(form.field(k).label for k in problems)}")
        records.append(record)
    return records


# --- outputs ------------------------------------------------------------------

def file_name(n: int) -> str:
    return f"form-{n:02}.jpg"


def write_labels(folder: str, form: FormTemplate, records: list[dict]) -> None:
    with open(os.path.join(folder, "labels.csv"), "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file"] + [field.label for field in form.fields])
        for n, record in enumerate(records, start=1):
            writer.writerow([file_name(n)] + [record[field.key][1] for field in form.fields])


def write_printable(folder: str, form: FormTemplate, records: list[dict]) -> str:
    """One blank form per page, then the sheets saying what to write on each."""
    esc = html.escape
    pages = []
    for n, _ in enumerate(records, start=1):
        rows = "".join(
            f'<div class="field {"tall" if field.type == "longtext" else ""}"><span>{esc(field.label)}'
            f'{" (" + esc(", ".join(field.choices)) + ")" if field.type == "choice" else ""}</span></div>'
            for field in form.fields)
        pages.append(f'<section class="form"><p class="id">Form {n:02}</p><h1>{esc(form.name)}</h1>{rows}</section>')
    key_rows = "".join(
        f"<tr><th>Form {n:02}</th>" + "".join(f"<td>{esc(record[f.key][0]) or '<em>leave blank</em>'}</td>"
                                               for f in form.fields) + "</tr>"
        for n, record in enumerate(records, start=1))
    header = "".join(f"<th>{esc(f.label)}</th>" for f in form.fields)
    page = f"""<!doctype html><html><head><meta charset="utf-8"><title>{esc(form.name)} practice forms</title>
<style>
body {{ font-family: Arial, sans-serif; margin: 0; }}
.form {{ page-break-after: always; padding: 20mm 18mm; }}
.form h1 {{ font-size: 22pt; margin: 0 0 10mm; }}
.id {{ float: right; margin: 0; font-size: 11pt; color: #555; }}
.field {{ border-bottom: 1px solid #000; height: 18mm; margin-bottom: 4mm; position: relative; }}
.field.tall {{ height: 36mm; }}
.field span {{ font-size: 11pt; position: absolute; top: 0; }}
.key {{ padding: 12mm; font-size: 10pt; }}
.key table {{ border-collapse: collapse; width: 100%; }}
.key td, .key th {{ border: 1px solid #999; padding: 3px 5px; text-align: left; vertical-align: top; }}
.howto {{ padding: 12mm; font-size: 12pt; page-break-after: always; }}
@media screen {{ .form, .key, .howto {{ max-width: 210mm; margin: 10mm auto; box-shadow: 0 0 4px #aaa; }} }}
</style></head><body>
<section class="howto"><h1>Practice forms: how to use</h1><ol>
<li>Print this whole document. The first page is these instructions; then one blank form per page; the last pages are the answer key.</li>
<li>Fill in each form <strong>by hand</strong>, copying what the key says for that form number. Use your normal handwriting. Better still, ask a few people so there's a mix of handwriting.</li>
<li>Photograph each filled-in form with a phone, in order (Form 01 first). Normal lighting is fine; a few slightly crooked or shadowy photos make the test more realistic.</li>
<li>Copy the photos into the <code>{esc(os.path.basename(os.path.abspath(folder)))}</code> folder, then run<br>
<code>python -m drugtest_import.samples rename {esc(folder)}</code><br>to name them form-01.jpg, form-02.jpg, ... in the order taken.</li>
<li>Run <code>python -m drugtest_import.evaluate run {esc(folder)} --form {esc(form.key)}</code></li>
</ol><p>All names and details are made up.</p></section>
{''.join(pages)}
<section class="key"><h1>What to write on each form</h1><table><tr><th></th>{header}</tr>{key_rows}</table></section>
</body></html>"""
    path = os.path.join(folder, "print-forms.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(page)
    return path


# --- synthetic photos -----------------------------------------------------------

def find_font(names: list[str]) -> str | None:
    for directory in FONT_DIRS:
        if not os.path.isdir(directory):
            continue
        for root, _, files in os.walk(directory):
            lower = {f.lower(): f for f in files}
            for name in names:
                if name.lower() in lower:
                    return os.path.join(root, lower[name.lower()])
    return None


def load_font(path: str | None, size: int) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    return ImageFont.truetype(path, size) if path else ImageFont.load_default(size)


def draw_written(page: Image.Image, text: str, xy: tuple[int, int], font, rng: random.Random, width: int) -> None:
    """Draw text as if written: slightly rotated, dark blue/black ink, wrapped for long text."""
    lines = textwrap.wrap(text, width=max(10, width // max(1, int(font.size * 0.55)))) or [""]
    line_height = int(font.size * 1.25)
    layer = Image.new("RGBA", (width + 80, line_height * len(lines) + 40), (0, 0, 0, 0))
    ink = rng.choice([(20, 30, 90, 255), (25, 25, 25, 255), (10, 20, 120, 255)])
    draw = ImageDraw.Draw(layer)
    for i, line in enumerate(lines):
        draw.text((rng.randint(0, 25), 10 + i * line_height), line, font=font, fill=ink)
    layer = layer.rotate(rng.uniform(-2.5, 2.5), resample=Image.Resampling.BICUBIC, expand=True)
    page.paste(layer, xy, layer)


def render_synthetic(form: FormTemplate, number: int, record: dict, rng: random.Random,
                     print_font_path: str | None, hand_font_path: str | None) -> Image.Image:
    page = Image.new("RGB", PAGE, rng.choice([(250, 249, 244), (255, 255, 255), (246, 242, 230)]))
    draw = ImageDraw.Draw(page)
    title, label_font = load_font(print_font_path, 48), load_font(print_font_path, 24)
    draw.text((90, 80), form.name, font=title, fill=(0, 0, 0))
    draw.text((PAGE[0] - 260, 90), f"Form {number:02}", font=label_font, fill=(90, 90, 90))
    y = 200
    available = PAGE[1] - y - 100
    unit = available // sum(2 if f.type == "longtext" else 1 for f in form.fields)
    unit = min(unit, 170)
    for field in form.fields:
        height = unit * (2 if field.type == "longtext" else 1)
        label = field.label + (f" ({', '.join(field.choices)})" if field.type == "choice" else "")
        draw.text((90, y), label, font=label_font, fill=(0, 0, 0))
        draw.line((90, y + height - 18, PAGE[0] - 90, y + height - 18), fill=(0, 0, 0), width=2)
        written = record[field.key][0]
        if written:
            hand = load_font(hand_font_path, rng.randint(34, 44))
            draw_written(page, written, (100 + rng.randint(0, 40), y + 30), hand, rng, PAGE[0] - 260)
        y += height
    # Make it look like a phone photo: tilt, softness, uneven light, sensor noise
    page = page.rotate(rng.uniform(-3, 3), resample=Image.Resampling.BICUBIC, expand=True, fillcolor=(90, 85, 80))
    page = page.filter(ImageFilter.GaussianBlur(rng.uniform(0.4, 1.1)))
    darkest = rng.uniform(0.1, 0.25)  # one side of the page up to this much darker
    shade = Image.linear_gradient("L").rotate(rng.choice([0, 90, 180, 270])).resize(page.size)
    page = Image.composite(page, Image.new("RGB", page.size, (0, 0, 0)),
                           shade.point(lambda v: 255 - int(v * darkest)))
    noise = Image.effect_noise(page.size, rng.uniform(8, 18)).convert("RGB")
    return Image.blend(page, noise, 0.06)


# --- commands -------------------------------------------------------------------

def make(folder: str, form: FormTemplate, count: int, synthetic: bool, seed: int = 1) -> list[str]:
    os.makedirs(folder, exist_ok=True)
    if os.path.exists(os.path.join(folder, "labels.csv")):
        raise ValueError(f"{os.path.join(folder, 'labels.csv')} already exists. Use an empty folder.")
    if not 1 <= count <= 99:
        raise ValueError("--count must be between 1 and 99")
    records = fake_records(form, count, seed)
    write_labels(folder, form, records)
    made = [write_printable(folder, form, records), os.path.join(folder, "labels.csv")]
    if synthetic:
        print_font, hand_font = find_font(PRINT_FONTS), find_font(HANDWRITING_FONTS)
        rng = random.Random(seed)
        for n, record in enumerate(records, start=1):
            render_synthetic(form, n, record, rng, print_font, hand_font).save(
                os.path.join(folder, file_name(n)), quality=rng.randint(70, 88))
        made.append(f"{count} synthetic photos (handwriting font: "
                    f"{os.path.basename(hand_font) if hand_font else 'built-in'})")
    return made


def taken_at(path: str) -> tuple:
    """When the photo was taken (EXIF), falling back to the file's modified time."""
    try:
        with Image.open(path) as img:
            stamp = img.getexif().get_ifd(0x8769).get(36867) or img.getexif().get(306)
        if stamp:
            return (0, str(stamp), os.path.basename(path))
    except OSError:
        pass
    return (1, os.path.getmtime(path), os.path.basename(path))


def rename(folder: str) -> list[tuple[str, str]]:
    """Name new photos after the forms in labels.csv, in the order taken. Returns [(old, new)]."""
    path = os.path.join(folder, "labels.csv")
    if not os.path.exists(path):
        raise ValueError(f"No labels.csv in {folder}. Run 'make' first.")
    with open(path, newline="", encoding="utf-8-sig") as f:
        rows = list(csv.reader(f))
    expected = [row[0] for row in rows[1:]]
    photos = [f for f in os.listdir(folder) if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS
              and f not in expected]
    if not photos:
        raise ValueError("No new photos to rename. Copy your photos into the folder first.")
    if len(photos) != len(expected):
        raise ValueError(f"Found {len(photos)} photos but labels.csv lists {len(expected)} forms. "
                         "Add or remove photos so there's exactly one per form.")
    photos.sort(key=lambda name: taken_at(os.path.join(folder, name)))
    # Keep each photo's own extension (.jpg, .png, ...) and update labels.csv to match
    pairs = [(old, os.path.splitext(label)[0] + os.path.splitext(old)[1].lower())
             for old, label in zip(photos, expected)]
    clashes = [new for _, new in pairs if os.path.exists(os.path.join(folder, new))]
    if clashes:
        raise ValueError(f"{clashes[0]} already exists in the folder; move it away first.")
    for old, new in pairs:
        os.rename(os.path.join(folder, old), os.path.join(folder, new))
    for row, (_, new) in zip(rows[1:], pairs):
        row[0] = new
    with open(path, "w", newline="", encoding="utf-8") as f:
        csv.writer(f).writerows(rows)
    return pairs


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m drugtest_import.samples",
                                     description="Make practice forms with made-up people for accuracy testing.")
    commands = parser.add_subparsers(dest="command", required=True)
    mk = commands.add_parser("make", help="write labels.csv and printable practice forms")
    mk.add_argument("folder")
    mk.add_argument("--form", default="drug-test", help="form type key")
    mk.add_argument("--count", type=int, default=20)
    mk.add_argument("--synthetic", action="store_true", help="also draw the forms as fake photos")
    mk.add_argument("--seed", type=int, default=1, help="change to get different made-up people")
    rn = commands.add_parser("rename", help="name your photos form-01, form-02, ... in the order taken")
    rn.add_argument("folder")
    args = parser.parse_args(argv)
    try:
        if args.command == "make":
            form = load_form(args.form)
            made = make(args.folder, form, args.count, args.synthetic, args.seed)
            print("Made:\n  " + "\n  ".join(made))
            if args.synthetic:
                print(f"\nNow run:  python -m drugtest_import.evaluate run {args.folder} --form {form.key}")
                print("Remember: typed 'handwriting' is easier than real handwriting, so real accuracy will be lower.")
            else:
                print(f"\nOpen print-forms.html in a browser and print it. Instructions are on the first page.")
        else:
            for old, new in rename(args.folder):
                print(f"  {old} -> {new}")
            print(f"\nNow run:  python -m drugtest_import.evaluate run {args.folder}")
    except (ValueError, OSError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
