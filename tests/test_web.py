import io
import re
import time

import pytest
from fastapi.testclient import TestClient
from openpyxl import load_workbook
from PIL import Image

from drugtest_import import web
from drugtest_import.crypto import Cipher, generate_key
from drugtest_import.extract import ExtractionError
from drugtest_import.schema import Extraction
from drugtest_import.storage import Database

PASSWORD = "a-long-test-password"


def png(color="yellow") -> bytes:
    buffer = io.BytesIO()
    Image.new("RGB", (60, 40), color).save(buffer, format="PNG")
    return buffer.getvalue()


def fake_extractor(image):
    color = Image.open(image).getpixel((0, 0))
    if color == (255, 0, 0):
        raise ExtractionError("blurry")
    return Extraction(EmployeeID="emp005", Name="Bob Wilson", TestDate="9/18/2025", TestType="urine",
                      Result="Negative", uncertain_fields=["Result"])


@pytest.fixture
def db(tmp_path):
    db = Database(str(tmp_path / "web.db"), Cipher(generate_key()))
    for role in ("admin", "reviewer", "viewer"):
        db.create_user(role, PASSWORD, role, "setup")
    yield db
    db.close()


@pytest.fixture
def app(db, monkeypatch):
    monkeypatch.setenv("DRUGTEST_INSECURE_COOKIES", "1")
    return web.create_app(db, extractor=fake_extractor, workers=1)


def csrf(client, path):
    match = re.search(r'name="csrf" value="([^"]+)"', client.get(path).text)
    return match.group(1)


def login(app, username):
    client = TestClient(app)
    response = client.post("/login", data={"username": username, "password": PASSWORD,
                                           "csrf": csrf(client, "/login")}, follow_redirects=False)
    assert response.status_code == 303
    return client


def wait_until_read(db, timeout=5):
    deadline = time.time() + timeout
    while any(u["status"] == "reading" for u in db.pending_uploads()):
        assert time.time() < deadline, "extraction didn't finish"
        time.sleep(0.02)


def test_pages_require_login(app):
    client = TestClient(app)
    for path in ("/", "/review", "/records", "/admin/users", "/records/export.xlsx"):
        response = client.get(path, follow_redirects=False)
        assert response.status_code == 303 and response.headers["location"] == "/login"


def test_security_headers(app):
    response = TestClient(app).get("/login")
    assert "frame-ancestors 'none'" in response.headers["content-security-policy"]
    assert response.headers["cache-control"] == "no-store"


def test_wrong_password_then_lockout(app, db):
    client = TestClient(app)
    token = csrf(client, "/login")
    for _ in range(web.LOGIN_ATTEMPTS):
        response = client.post("/login", data={"username": "viewer", "password": "nope", "csrf": token})
        assert response.status_code == 401
    response = client.post("/login", data={"username": "viewer", "password": PASSWORD, "csrf": token})
    assert response.status_code == 429
    actions = [e["action"] for e in db.audit_entries()]
    assert actions.count("login.failed") == web.LOGIN_ATTEMPTS and "login.locked" in actions


def test_post_without_csrf_rejected(app):
    client = TestClient(app)
    client.get("/login")
    response = client.post("/login", data={"username": "admin", "password": PASSWORD})
    assert response.status_code == 400


def test_disabled_user_logged_out(app, db):
    client = login(app, "viewer")
    db.update_user(db.get_user_by_name("viewer")["id"], "test", active=False)
    assert client.get("/records", follow_redirects=False).headers["location"] == "/login"


def test_idle_session_expires(app, monkeypatch):
    client = login(app, "viewer")
    assert client.get("/records").status_code == 200
    real_time = time.time
    monkeypatch.setattr(web.time, "time", lambda: real_time() + web.IDLE_TIMEOUT + 1)
    assert client.get("/records", follow_redirects=False).headers["location"] == "/login"


def test_roles(app):
    viewer = login(app, "viewer")
    assert viewer.get("/records").status_code == 200
    assert viewer.get("/review").status_code == 403
    assert viewer.get("/records/export.xlsx").status_code == 403
    assert viewer.get("/admin/users").status_code == 403
    reviewer = login(app, "reviewer")
    assert reviewer.get("/review").status_code == 200
    assert reviewer.get("/admin/audit").status_code == 403


def test_upload_review_save_export(app, db):
    client = login(app, "reviewer")
    token = csrf(client, "/review")
    files = [("files", ("note1.png", png(), "image/png")), ("files", ("bad.png", png("red"), "image/png")),
             ("files", ("notes.txt", b"not an image", "text/plain"))]
    response = client.post("/uploads", data={"csrf": token}, files=files)
    assert response.status_code == 200 and "notes.txt" in response.text  # partial success page
    wait_until_read(db)
    first, second = db.pending_uploads()
    assert second["status"] == "failed" and second["error"] == "blurry"

    page = client.get(f"/review/{first['id']}").text
    assert 'value="2025-09-18"' in page and "Hard to read" in page
    image = client.get(f"/uploads/{first['id']}/image")
    assert image.headers["content-type"] == "image/png" and image.content == png()

    form = {"csrf": token, "action": "save", "EmployeeID": "EMP005", "Name": "Bob Wilson",
            "TestDate": "2025-09-18", "TestType": "Urine", "Result": "Negative", "Notes": ""}
    assert client.post(f"/review/{first['id']}", data={**form, "Result": "Negatve"}).status_code == 422
    response = client.post(f"/review/{first['id']}", data=form)
    assert response.status_code == 422 and "Tick the box" in response.text
    response = client.post(f"/review/{first['id']}", data={**form, "confirm_uncertain": "1"}, follow_redirects=False)
    assert response.status_code == 303
    assert db.count() == 1 and db.get_upload(first["id"]) is None

    # The failed one: entered by hand, then caught as a duplicate
    response = client.post(f"/review/{second['id']}", data=form)
    assert response.status_code == 409 and "already recorded" in response.text
    client.post(f"/review/{second['id']}", data={**form, "action": "skip"})
    assert db.pending_uploads() == [] and db.count() == 1

    records = client.get("/records?q=bob").text
    assert "Bob Wilson" in records and "reviewer" in records
    assert "Bob Wilson" not in client.get("/records?result=Positive").text

    export = client.get("/records/export.xlsx")
    ws = load_workbook(io.BytesIO(export.content)).active
    assert ws["B2"].value == "Bob Wilson" and ws["J2"].value == "reviewer"

    actions = [e["action"] for e in db.audit_entries()]
    for action in ("upload.create", "record.create", "upload.skip", "records.view", "records.export"):
        assert action in actions


def test_admin_user_management(app, db):
    client = login(app, "admin")
    token = csrf(client, "/admin/users")
    response = client.post("/admin/users", data={"csrf": token, "username": "nurse", "password": "short", "role": "viewer"})
    assert response.status_code == 422
    client.post("/admin/users", data={"csrf": token, "username": "nurse", "password": PASSWORD, "role": "reviewer"})
    assert db.get_user_by_name("nurse")["role"] == "reviewer"

    admin_id = db.get_user_by_name("admin")["id"]
    response = client.post(f"/admin/users/{admin_id}", data={"csrf": token, "role": "viewer"})
    assert response.status_code == 422 and "at least one active admin" in response.text
    nurse_id = db.get_user_by_name("nurse")["id"]
    client.post(f"/admin/users/{nurse_id}", data={"csrf": token, "role": "reviewer", "action": "disable"})
    assert db.get_user(nurse_id)["active"] == 0
    assert "user.disable" in [e["action"] for e in db.audit_entries()]
    assert "user.disable" in client.get("/admin/audit").text
