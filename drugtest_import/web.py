"""Web app: sign in, upload photos, review what Claude read, browse and export records.

Run with:  python -m drugtest_import.web   (or: uvicorn --factory drugtest_import.web:create_app)
"""
import io
import logging
import os
import secrets
import tempfile
import time
from collections import defaultdict
from concurrent.futures import ThreadPoolExecutor
from contextlib import asynccontextmanager
from datetime import date

from fastapi import Depends, FastAPI, File, Form, Request, UploadFile
from fastapi.responses import HTMLResponse, RedirectResponse, Response
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from PIL import Image, UnidentifiedImageError
from starlette.middleware.sessions import SessionMiddleware

from .crypto import Cipher, hash_password, verify_password
from .extract import ExtractionError, extract_fields
from .schema import FIELDS, RESULTS, TEST_TYPES, normalize, validate
from .storage import ROLES, Database

log = logging.getLogger(__name__)
HERE = os.path.dirname(__file__)

ROLE_RANK = {"viewer": 0, "reviewer": 1, "admin": 2}
SESSION_MAX_AGE = 8 * 3600
IDLE_TIMEOUT = 30 * 60
MAX_UPLOAD_BYTES = 15 * 1024 * 1024
MAX_FILES_PER_UPLOAD = 50
LOGIN_ATTEMPTS = 5
LOGIN_WINDOW = 15 * 60
IMAGE_TYPES = {"JPEG": "image/jpeg", "PNG": "image/png", "WEBP": "image/webp"}
LABELS = {"EmployeeID": "Employee ID", "TestDate": "Test date", "TestType": "Test type"}
# Compared against when a username doesn't exist, so response time doesn't reveal valid usernames
DUMMY_HASH = hash_password("timing-equaliser")

SECURITY_HEADERS = {
    "Content-Security-Policy": "default-src 'self'; img-src 'self'; style-src 'self'; script-src 'none'; "
                               "frame-ancestors 'none'; form-action 'self'; base-uri 'none'",
    "X-Content-Type-Options": "nosniff",
    "X-Frame-Options": "DENY",
    "Referrer-Policy": "no-referrer",
    "Cache-Control": "no-store",
}


class Redirect(Exception):
    def __init__(self, url: str):
        self.url = url


class Forbidden(Exception):
    pass


class BadRequest(Exception):
    pass


def create_app(db: Database | None = None, extractor=extract_fields, workers: int = 3) -> FastAPI:
    if db is None:
        db_path = os.environ.get("DRUGTEST_DB", "drugtest.db")
        db = Database(db_path, Cipher.from_env(), os.environ.get("DRUGTEST_UPLOADS"))
    secure_cookies = os.environ.get("DRUGTEST_INSECURE_COOKIES") != "1"

    pool = ThreadPoolExecutor(max_workers=workers)

    @asynccontextmanager
    async def lifespan(app):
        yield
        pool.shutdown(wait=False, cancel_futures=True)

    app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None, lifespan=lifespan)
    app.add_middleware(
        SessionMiddleware, secret_key=db.cipher.session_secret, session_cookie="drugtest_session",
        max_age=SESSION_MAX_AGE, same_site="strict", https_only=secure_cookies,
    )
    app.mount("/static", StaticFiles(directory=os.path.join(HERE, "static")), name="static")
    templates = Jinja2Templates(directory=os.path.join(HERE, "templates"))
    failed_logins: dict[str, list[float]] = defaultdict(list)
    app.state.db, app.state.pool = db, pool

    @app.middleware("http")
    async def security_headers(request: Request, call_next):
        response = await call_next(request)
        for name, value in SECURITY_HEADERS.items():
            response.headers.setdefault(name, value)
        return response

    @app.exception_handler(Redirect)
    async def on_redirect(request: Request, exc: Redirect):
        return RedirectResponse(exc.url, status_code=303)

    @app.exception_handler(Forbidden)
    async def on_forbidden(request: Request, exc: Forbidden):
        return render(request, "message.html", {"title": "Not allowed",
                      "message": "Your account doesn't have access to this page."}, status_code=403)

    @app.exception_handler(BadRequest)
    async def on_bad_request(request: Request, exc: BadRequest):
        return render(request, "message.html", {"title": "Request rejected", "message": str(exc)}, status_code=400)

    # --- helpers ----------------------------------------------------------

    def csrf_token(request: Request) -> str:
        if "csrf" not in request.session:
            request.session["csrf"] = secrets.token_urlsafe(32)
        return request.session["csrf"]

    async def check_csrf(request: Request) -> None:
        form = await request.form()
        expected = request.session.get("csrf")
        if not expected or not secrets.compare_digest(str(form.get("csrf", "")), expected):
            raise BadRequest("Your session expired or the form was stale. Go back, reload the page and try again.")

    def render(request: Request, template: str, context: dict | None = None, status_code: int = 200):
        context = dict(context or {})
        context.setdefault("user", getattr(request.state, "user", None))
        context["csrf"] = csrf_token(request)
        context["can"] = lambda role: bool(context["user"]) and ROLE_RANK[context["user"]["role"]] >= ROLE_RANK[role]
        return templates.TemplateResponse(request, template, context, status_code=status_code)

    def current_user(request: Request) -> dict:
        user_id = request.session.get("user_id")
        last_seen = request.session.get("last_seen", 0)
        user = db.get_user(user_id) if user_id else None
        if not user or not user["active"] or time.time() - last_seen > IDLE_TIMEOUT:
            request.session.clear()
            raise Redirect("/login")
        request.session["last_seen"] = int(time.time())
        request.state.user = user
        return user

    def require(role: str):
        def dependency(user: dict = Depends(current_user)) -> dict:
            if ROLE_RANK[user["role"]] < ROLE_RANK[role]:
                raise Forbidden()
            return user
        return dependency

    viewer, reviewer, admin = require("viewer"), require("reviewer"), require("admin")

    def run_extraction(upload_id: str) -> None:
        try:
            image = db.upload_image(upload_id)
            result = extractor(io.BytesIO(image))
            db.set_extraction(upload_id, result.model_dump())
        except (ExtractionError, OSError) as e:
            db.set_extraction(upload_id, None, str(e))
        except Exception:
            log.exception("Extraction failed for upload %s", upload_id)
            db.set_extraction(upload_id, None, "Unexpected error while reading the image.")

    # Re-queue anything that was still being read when the server last stopped
    for upload in db.pending_uploads():
        if upload["status"] == "reading":
            pool.submit(run_extraction, upload["id"])

    # --- sign in ----------------------------------------------------------

    @app.get("/login", response_class=HTMLResponse)
    def login_page(request: Request):
        return render(request, "login.html", {"no_users": db.user_count() == 0})

    @app.post("/login")
    async def login(request: Request, username: str = Form(""), password: str = Form("")):
        await check_csrf(request)
        key = username.strip().lower()
        recent = [t for t in failed_logins[key] if time.time() - t < LOGIN_WINDOW]
        failed_logins[key] = recent
        if len(recent) >= LOGIN_ATTEMPTS:
            db.log(username or "-", "login.locked")
            return render(request, "login.html", {"error": "Too many failed attempts. Try again in 15 minutes."},
                          status_code=429)
        user = db.get_user_by_name(username) if username.strip() else None
        valid = verify_password(password, user["password_hash"] if user else DUMMY_HASH)
        if not (user and valid and user["active"]):
            failed_logins[key].append(time.time())
            db.log(username or "-", "login.failed")
            return render(request, "login.html", {"error": "Wrong username or password."}, status_code=401)
        failed_logins.pop(key, None)
        request.session.clear()  # new session id contents on login
        request.session.update({"user_id": user["id"], "last_seen": int(time.time()),
                                "csrf": secrets.token_urlsafe(32)})
        db.record_login(user["id"])
        db.log(user["username"], "login")
        return RedirectResponse("/", status_code=303)

    @app.post("/logout")
    async def logout(request: Request, user: dict = Depends(current_user)):
        await check_csrf(request)
        db.log(user["username"], "logout")
        request.session.clear()
        return RedirectResponse("/login", status_code=303)

    @app.get("/")
    def home(user: dict = Depends(current_user)):
        return RedirectResponse("/review" if ROLE_RANK[user["role"]] >= ROLE_RANK["reviewer"] else "/records",
                                status_code=303)

    # --- upload and review -----------------------------------------------

    @app.get("/review", response_class=HTMLResponse)
    def review_index(request: Request, user: dict = Depends(reviewer)):
        uploads = db.pending_uploads()
        return render(request, "review_index.html", {
            "uploads": uploads, "reading": any(u["status"] == "reading" for u in uploads),
            "max_files": MAX_FILES_PER_UPLOAD, "max_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
        })

    @app.post("/uploads")
    async def upload(request: Request, files: list[UploadFile] = File(...), user: dict = Depends(reviewer)):
        await check_csrf(request)
        files = [f for f in files if f.filename]
        if not files:
            raise BadRequest("Choose at least one photo.")
        if len(files) > MAX_FILES_PER_UPLOAD:
            raise BadRequest(f"Upload at most {MAX_FILES_PER_UPLOAD} photos at a time.")
        accepted, rejected = [], []
        for f in files:
            data = await f.read(MAX_UPLOAD_BYTES + 1)
            if len(data) > MAX_UPLOAD_BYTES or not image_format(data):
                rejected.append(f.filename)
                continue
            accepted.append(db.create_upload(f.filename, data, user["username"]))
        for upload_id in accepted:
            pool.submit(run_extraction, upload_id)
        if rejected:
            return render(request, "message.html", {
                "title": "Some files were skipped",
                "message": f"{len(accepted)} photo(s) uploaded. These weren't JPEG/PNG/WebP images under "
                           f"{MAX_UPLOAD_BYTES // (1024 * 1024)} MB: {', '.join(rejected)}",
                "link": "/review/next", "link_text": "Start reviewing",
            }, status_code=400 if not accepted else 200)
        return RedirectResponse("/review/next", status_code=303)

    @app.get("/review/next")
    def review_next(user: dict = Depends(reviewer)):
        uploads = db.pending_uploads()
        if not uploads:
            return RedirectResponse("/review", status_code=303)
        return RedirectResponse(f"/review/{uploads[0]['id']}", status_code=303)

    def get_upload_or_404(upload_id: str) -> dict:
        try:
            upload = db.get_upload(upload_id)
        except ValueError:
            upload = None
        if not upload:
            raise Redirect("/review/next")
        return upload

    def review_context(upload: dict, values: dict | None = None, **extra) -> dict:
        extraction = upload["extraction"] or {}
        if values is None:
            values = normalize(extraction) if extraction else {key: None for key in FIELDS}
        uploads = db.pending_uploads()
        position = next((i for i, u in enumerate(uploads) if u["id"] == upload["id"]), 0)
        return {"upload": upload, "values": values, "fields": FIELDS, "labels": LABELS,
                "choices": {"TestType": TEST_TYPES, "Result": RESULTS},
                "uncertain": set(extraction.get("uncertain_fields", [])),
                "position": position + 1, "total": len(uploads), "problems": {}, **extra}

    @app.get("/review/{upload_id}", response_class=HTMLResponse)
    def review(request: Request, upload_id: str, user: dict = Depends(reviewer)):
        upload = get_upload_or_404(upload_id)
        return render(request, "review.html", review_context(upload))

    @app.get("/uploads/{upload_id}/image")
    def upload_image(upload_id: str, user: dict = Depends(reviewer)):
        get_upload_or_404(upload_id)
        try:
            data = db.upload_image(upload_id)
        except FileNotFoundError:
            return Response(status_code=404)
        return Response(data, media_type=IMAGE_TYPES.get(image_format(data), "application/octet-stream"))

    @app.post("/review/{upload_id}")
    async def review_submit(request: Request, upload_id: str, user: dict = Depends(reviewer)):
        await check_csrf(request)
        form = await request.form()
        upload = get_upload_or_404(upload_id)
        if form.get("action") == "skip":
            db.delete_upload(upload_id)
            db.log(user["username"], "upload.skip", upload_id)
            return RedirectResponse("/review/next", status_code=303)

        record = normalize({key: form.get(key) for key in FIELDS})
        context = review_context(upload, values=record)
        problems = validate(record)
        if problems:
            return render(request, "review.html", {**context, "problems": problems,
                          "error": "Fix the highlighted fields before saving."}, status_code=422)
        if context["uncertain"] and not form.get("confirm_uncertain"):
            return render(request, "review.html", {**context, "error":
                          "Tick the box to confirm you checked the hard-to-read fields against the photo."},
                          status_code=422)
        duplicate = db.find_duplicate(record)
        if duplicate and not form.get("confirm_duplicate"):
            return render(request, "review.html", {**context, "duplicate": duplicate}, status_code=409)
        db.add_record(record, user["username"], source_file=upload["filename"])
        db.delete_upload(upload_id)
        return RedirectResponse("/review/next", status_code=303)

    # --- records ----------------------------------------------------------

    def filtered_records(q: str, result: str, start: str, end: str) -> list[dict]:
        records = db.records()
        q = q.strip().lower()
        if q:
            records = [r for r in records if q in (r["Name"] or "").lower() or q in (r["EmployeeID"] or "").lower()]
        if result:
            records = [r for r in records if r["Result"] == result]
        if start:
            records = [r for r in records if r["TestDate"] >= start]
        if end:
            records = [r for r in records if r["TestDate"] <= end]
        return records

    def filter_summary(q, result, start, end) -> str | None:
        parts = [f"{k}={v}" for k, v in (("q", q), ("result", result), ("from", start), ("to", end)) if v]
        return ", ".join(parts) or None

    @app.get("/records", response_class=HTMLResponse)
    def records(request: Request, q: str = "", result: str = "", start: str = "", end: str = "",
                user: dict = Depends(viewer)):
        rows = filtered_records(q, result, start, end)
        db.log(user["username"], "records.view", detail=filter_summary(q, result, start, end))
        return render(request, "records.html", {
            "records": rows, "fields": FIELDS, "labels": LABELS, "results": RESULTS,
            "filters": {"q": q, "result": result, "start": start, "end": end},
        })

    @app.get("/records/export.xlsx")
    def export(q: str = "", result: str = "", start: str = "", end: str = "", user: dict = Depends(reviewer)):
        rows = filtered_records(q, result, start, end)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "export.xlsx")
            db.export_xlsx(path, rows)
            with open(path, "rb") as f:
                data = f.read()
        db.log(user["username"], "records.export", detail=f"{len(rows)} rows; {filter_summary(q, result, start, end) or 'all'}")
        return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": f'attachment; filename="DrugTests-{date.today()}.xlsx"'})

    # --- admin ------------------------------------------------------------

    @app.get("/admin/users", response_class=HTMLResponse)
    def users_page(request: Request, user: dict = Depends(admin)):
        return render(request, "users.html", {"users": db.users(), "roles": ROLES})

    @app.post("/admin/users")
    async def create_user(request: Request, username: str = Form(""), password: str = Form(""),
                          role: str = Form("viewer"), user: dict = Depends(admin)):
        await check_csrf(request)
        try:
            db.create_user(username, password, role, user["username"])
        except ValueError as e:
            return render(request, "users.html", {"users": db.users(), "roles": ROLES, "error": str(e)},
                          status_code=422)
        return RedirectResponse("/admin/users", status_code=303)

    @app.post("/admin/users/{user_id}")
    async def update_user(request: Request, user_id: int, user: dict = Depends(admin)):
        await check_csrf(request)
        form = await request.form()
        target = db.get_user(user_id)
        if not target:
            raise BadRequest("No such user.")
        role = form.get("role") if form.get("role") != target["role"] else None
        active = {"enable": True, "disable": False}.get(form.get("action", ""))
        password = form.get("password") or None
        removes_admin = target["role"] == "admin" and target["active"] and (
            (role and role != "admin") or active is False)
        error = None
        if removes_admin and db.active_admin_count() <= 1:
            error = "There must always be at least one active admin."
        elif target["id"] == user["id"] and active is False:
            error = "You can't disable your own account."
        else:
            try:
                db.update_user(user_id, user["username"], role=role, active=active, password=password)
            except ValueError as e:
                error = str(e)
        if error:
            return render(request, "users.html", {"users": db.users(), "roles": ROLES, "error": error},
                          status_code=422)
        return RedirectResponse("/admin/users", status_code=303)

    @app.get("/admin/audit", response_class=HTMLResponse)
    def audit(request: Request, user: dict = Depends(admin)):
        return render(request, "audit.html", {"entries": db.audit_entries()})

    return app


def image_format(data: bytes) -> str | None:
    """Return the image format if `data` is an allowed, readable image."""
    try:
        with Image.open(io.BytesIO(data)) as img:
            img.verify()
            return img.format if img.format in IMAGE_TYPES else None
    except (UnidentifiedImageError, OSError, SyntaxError, ValueError):
        return None


def main():
    import uvicorn

    host = os.environ.get("DRUGTEST_HOST", "127.0.0.1")
    port = int(os.environ.get("DRUGTEST_PORT", "8000"))
    uvicorn.run(create_app(), host=host, port=port, proxy_headers=True)


if __name__ == "__main__":
    main()
