"""Append confirmed records to the CSV log."""
import csv
import os
import tempfile

from .schema import FIELDS


def append_record(record: dict, path: str) -> None:
    """Append one record, writing the header if the file is new.

    The whole file is rewritten via a temp file + rename so a crash mid-write
    can't corrupt existing records. Values are kept as text, so IDs like
    '007123' keep their leading zeros.
    """
    rows = []
    if os.path.exists(path):
        with open(path, newline="", encoding="utf-8-sig") as f:
            rows = list(csv.DictReader(f))
    rows.append({key: record.get(key) or "" for key in FIELDS})

    directory = os.path.dirname(os.path.abspath(path))
    fd, tmp_path = tempfile.mkstemp(dir=directory, suffix=".tmp")
    try:
        with os.fdopen(fd, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
            writer.writeheader()
            writer.writerows(rows)
        os.replace(tmp_path, path)
    except BaseException:
        os.unlink(tmp_path)
        raise
