"""Measure how accurately the AI reads your forms, and what each scan costs.

  1. Put photos of real forms in a folder.
  2. python -m drugtest_import.evaluate init FOLDER --form drug-test
     -> writes FOLDER/labels.csv with one row per photo. Type in what each form
        actually says (leave a cell empty if the form left it blank, or put ? to
        skip that field for that photo).
  3. python -m drugtest_import.evaluate run FOLDER --form drug-test
     -> reads every photo with the AI (this calls the paid API), compares against
        labels.csv and prints a report. Per-field results go to
        FOLDER/accuracy-results.csv.

The results and labels contain the forms' contents, so keep them as private as the photos.
"""
import argparse
import csv
import os
import sys
import time
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field

from . import extract
from .forms import BUILT_IN_FORMS, FormTemplate, normalize

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
SKIP = "?"
# USD per million tokens (input, output). Update when prices change.
PRICES = {
    "claude-opus-5-5": (4.00, 20.00),
    "claude-sonnet-5-5": (2.00, 10.00),
    "claude-haiku-4-5": (1.00, 5.00),
    "claude-fable-5-1": (10.00, 50.00),
}


def find_images(folder: str) -> list[str]:
    return sorted(f for f in os.listdir(folder) if os.path.splitext(f)[1].lower() in IMAGE_EXTENSIONS)


def load_form(key: str) -> FormTemplate:
    """Built-in form types by key; custom ones from the database (needs DRUGTEST_KEY)."""
    for form in BUILT_IN_FORMS:
        if form.key == key:
            built_in = form
            break
    else:
        built_in = None
    if os.environ.get("DRUGTEST_KEY"):
        from .crypto import Cipher
        from .storage import Database
        db_path = os.environ.get("DRUGTEST_DB", "drugtest.db")
        if os.path.exists(db_path):
            db = Database(db_path, Cipher.from_env(), os.environ.get("DRUGTEST_UPLOADS"))
            try:
                if form := db.get_form(key):
                    return form  # the customer's edited version wins
            finally:
                db.close()
    if built_in:
        return built_in
    raise ValueError(f"No form type {key!r}. Built in: {', '.join(f.key for f in BUILT_IN_FORMS)}")


def write_template(folder: str, form: FormTemplate) -> str:
    path = os.path.join(folder, "labels.csv")
    if os.path.exists(path):
        raise ValueError(f"{path} already exists; not overwriting it.")
    images = find_images(folder)
    if not images:
        raise ValueError(f"No JPEG/PNG/WebP images in {folder}")
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file"] + [field.label for field in form.fields])
        for image in images:
            writer.writerow([image] + [""] * len(form.fields))
    return path


def load_labels(folder: str, form: FormTemplate) -> dict[str, dict]:
    """{file name: {field key: expected value, or SKIP}} from labels.csv (headers = labels or keys)."""
    by_header = {f.key: f.key for f in form.fields} | {f.label.lower(): f.key for f in form.fields}
    labels = {}
    with open(os.path.join(folder, "labels.csv"), newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        unknown = [h for h in reader.fieldnames or [] if h != "file" and (h or "").lower() not in by_header]
        if unknown:
            raise ValueError(f"labels.csv has columns that aren't fields of {form.name}: {', '.join(unknown)}")
        for row in reader:
            name = (row.pop("file", "") or "").strip()
            if not name:
                continue
            labels[name] = {by_header[h.lower()]: (v or "").strip() for h, v in row.items() if h}
    return labels


def comparable(form: FormTemplate, key: str, value: str | None) -> str:
    """Normalise a value the same way the app does, then ignore case and spacing for text."""
    value = normalize(form, {key: value})[key] or ""
    if form.field(key).type in ("text", "longtext"):
        value = " ".join(value.casefold().split()).rstrip(".")
    return value


@dataclass
class FieldResult:
    file: str
    key: str
    expected: str
    got: str
    correct: bool
    flagged: bool


@dataclass
class Report:
    form: FormTemplate
    results: list[FieldResult] = field(default_factory=list)
    failures: dict[str, str] = field(default_factory=dict)  # file -> error
    input_tokens: int = 0
    output_tokens: int = 0
    models: set = field(default_factory=set)
    seconds: float = 0.0
    documents: int = 0

    def cost(self) -> float | None:
        if len(self.models) != 1 or (model := next(iter(self.models))) not in PRICES:
            return None
        price_in, price_out = PRICES[model]
        return (self.input_tokens * price_in + self.output_tokens * price_out) / 1_000_000


def evaluate(folder: str, form: FormTemplate, workers: int = 3, limit: int | None = None,
             reader=extract.read_form, progress=None) -> Report:
    labels = load_labels(folder, form)
    missing = [name for name in labels if not os.path.exists(os.path.join(folder, name))]
    if missing:
        raise ValueError(f"labels.csv lists files that aren't in {folder}: {', '.join(missing[:5])}")
    names = list(labels)[:limit] if limit else list(labels)
    report = Report(form=form, documents=len(names))
    started = time.monotonic()

    def run(name):
        try:
            return name, reader(os.path.join(folder, name), form), None
        except (extract.ExtractionError, OSError) as e:
            return name, None, str(e)

    with ThreadPoolExecutor(max_workers=workers) as pool:
        for done, (name, reading, error) in enumerate(pool.map(run, names), start=1):
            if progress:
                progress(done, len(names), name, error)
            if error:
                report.failures[name] = error
                continue
            report.input_tokens += reading.input_tokens
            report.output_tokens += reading.output_tokens
            report.models.add(reading.model)
            for key, expected in labels[name].items():
                if expected == SKIP:
                    continue
                want = comparable(form, key, expected)
                got = comparable(form, key, reading.values.get(key))
                report.results.append(FieldResult(name, key, expected, reading.values.get(key) or "",
                                                  want == got, key in reading.uncertain))
    report.seconds = time.monotonic() - started
    return report


def pct(part: int, whole: int) -> str:
    return f"{100 * part / whole:.1f}%" if whole else "n/a"


def summary(report: Report) -> str:
    results = report.results
    wrong = [r for r in results if not r.correct]
    silent = [r for r in wrong if not r.flagged]
    flagged = [r for r in results if r.flagged]
    read_ok = report.documents - len(report.failures)
    perfect = {r.file for r in results} - {r.file for r in wrong}
    lines = [
        f"Form type: {report.form.name}",
        f"Photos: {report.documents} ({len(report.failures)} couldn't be read)",
        f"Fields checked: {len(results)}",
        "",
        f"Field accuracy:            {pct(len(results) - len(wrong), len(results))}  ({len(wrong)} wrong)",
        f"Photos with every field right: {pct(len(perfect), read_ok)}",
        f"Wrong fields that were flagged 'hard to read': {pct(len(wrong) - len(silent), len(wrong))}",
        f"Silent errors (wrong and NOT flagged): {len(silent)}  ({pct(len(silent), len(results))} of fields)",
        f"Fields flagged for double-checking: {pct(len(flagged), len(results))}",
        "",
        "Per field:",
    ]
    width = max(len(f.label) for f in report.form.fields)
    for f in report.form.fields:
        rows = [r for r in results if r.key == f.key]
        if not rows:
            continue
        right = sum(r.correct for r in rows)
        silent_here = sum(1 for r in rows if not r.correct and not r.flagged)
        lines.append(f"  {f.label:<{width}}  {pct(right, len(rows)):>6}  ({len(rows) - right} wrong, {silent_here} silent)")
    if silent:
        lines += ["", "Silent errors (check these first):"]
        lines += [f"  {r.file}: {report.form.field(r.key).label} - expected {r.expected!r}, read {r.got!r}"
                  for r in silent[:20]]
        if len(silent) > 20:
            lines.append(f"  ... and {len(silent) - 20} more in accuracy-results.csv")
    if report.failures:
        lines += ["", "Couldn't read:"] + [f"  {name}: {error}" for name, error in report.failures.items()]
    lines += ["", f"Time: {report.seconds:.0f}s"]
    if read_ok:
        lines.append(f"Tokens per photo: {report.input_tokens / read_ok:,.0f} in, {report.output_tokens / read_ok:,.0f} out"
                     f" (model: {', '.join(sorted(report.models))})")
        cost = report.cost()
        lines.append(f"Cost: ${cost:.4f} total, ${cost / read_ok:.4f} per photo" if cost is not None
                     else "Cost: unknown model price - see PRICES in evaluate.py")
    return "\n".join(lines)


def write_results(report: Report, path: str) -> None:
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["file", "field", "expected", "read", "correct", "flagged"])
        for r in report.results:
            writer.writerow([r.file, report.form.field(r.key).label, r.expected, r.got,
                             "yes" if r.correct else "NO", "yes" if r.flagged else ""])


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(prog="python -m drugtest_import.evaluate",
                                     description="Measure AI reading accuracy against hand-checked labels.")
    commands = parser.add_subparsers(dest="command", required=True)
    init = commands.add_parser("init", help="create FOLDER/labels.csv to fill in")
    run = commands.add_parser("run", help="read every photo and compare with labels.csv (uses the paid API)")
    for sub in (init, run):
        sub.add_argument("folder")
        sub.add_argument("--form", default="drug-test", help="form type key")
    run.add_argument("--workers", type=int, default=3, help="photos read in parallel")
    run.add_argument("--limit", type=int, help="only the first N photos")
    run.add_argument("--model", help="override the model, e.g. claude-sonnet-5-5")
    args = parser.parse_args(argv)

    try:
        form = load_form(args.form)
        if args.command == "init":
            path = write_template(args.folder, form)
            print(f"Wrote {path}. Fill in what each form actually says, then run:\n"
                  f"  python -m drugtest_import.evaluate run {args.folder} --form {form.key}")
            return 0
        if args.model:
            extract.MODEL = args.model

        def progress(done, total, name, error):
            print(f"\r{done}/{total} read{' - ' + name + ' failed' if error else ''}", end="", file=sys.stderr)

        report = evaluate(args.folder, form, workers=args.workers, limit=args.limit, progress=progress)
        print(file=sys.stderr)
        results_path = os.path.join(args.folder, "accuracy-results.csv")
        write_results(report, results_path)
        print(summary(report))
        print(f"\nPer-field results: {results_path}")
    except (ValueError, OSError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
