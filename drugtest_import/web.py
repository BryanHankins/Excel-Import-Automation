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
from pydantic import ValidationError
from starlette.middleware.sessions import SessionMiddleware

from .crypto import Cipher, hash_password, verify_password
from .extract import ExtractionError, extract_fields
from .forms import FIELD_TYPES, FieldDef, FormTemplate, make_key, normalize, validate
from .storage import ROLES, Database, FormInUse

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
            upload = db.get_upload(upload_id)
            form = db.get_form(upload["form_key"]) if upload else None
            if not form:
                return  # skipped or form type deleted while queued
            image = db.upload_image(upload_id)
            values, uncertain = extractor(io.BytesIO(image), form)
            db.set_extraction(upload_id, {"values": values, "uncertain": uncertain})
        except (ExtractionError, OSError) as e:
            db.set_extraction(upload_id, None, str(e))
        except Exception:
            log.exception("Extraction failed for upload %s", upload_id)
            db.set_extraction(upload_id, None, "Unexpected error while reading the image.")

    # Re-queue anything that was still being read when the server last stopped
    for upload in db.pending_uploads():
        if upload["status"] == "reading":
            pool.submit(run_extraction, upload["id"])

    def get_form_or_400(key: str) -> FormTemplate:
        form = db.get_form(key)
        if not form:
            raise BadRequest("Unknown form type.")
        return form

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
        names = {f.key: f.name for f in db.forms()}
        return render(request, "review_index.html", {
            "uploads": uploads, "reading": any(u["status"] == "reading" for u in uploads),
            "forms": db.forms(), "form_names": names, "selected": request.session.get("last_form"),
            "max_files": MAX_FILES_PER_UPLOAD, "max_mb": MAX_UPLOAD_BYTES // (1024 * 1024),
        })

    @app.post("/uploads")
    async def upload(request: Request, files: list[UploadFile] = File(...), form_key: str = Form(""),
                     user: dict = Depends(reviewer)):
        await check_csrf(request)
        form = get_form_or_400(form_key)
        files = [f for f in files if f.filename]
        if not files:
            raise BadRequest("Choose at least one photo.")
        if len(files) > MAX_FILES_PER_UPLOAD:
            raise BadRequest(f"Upload at most {MAX_FILES_PER_UPLOAD} photos at a time.")
        request.session["last_form"] = form.key
        accepted, rejected = [], []
        for f in files:
            data = await f.read(MAX_UPLOAD_BYTES + 1)
            if len(data) > MAX_UPLOAD_BYTES or not image_format(data):
                rejected.append(f.filename)
                continue
            accepted.append(db.create_upload(f.filename, data, form.key, user["username"]))
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

    def review_context(upload: dict, form: FormTemplate, values: dict | None = None, **extra) -> dict:
        extraction = upload["extraction"] or {}
        if values is None:
            values = normalize(form, extraction.get("values", {}))
        uploads = db.pending_uploads()
        position = next((i for i, u in enumerate(uploads) if u["id"] == upload["id"]), 0)
        return {"upload": upload, "form": form, "values": values, "uncertain": set(extraction.get("uncertain", [])),
                "position": position + 1, "total": len(uploads), "problems": {}, **extra}

    @app.get("/review/{upload_id}", response_class=HTMLResponse)
    def review(request: Request, upload_id: str, user: dict = Depends(reviewer)):
        upload = get_upload_or_404(upload_id)
        return render(request, "review.html", review_context(upload, get_form_or_400(upload["form_key"])))

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
        submitted = await request.form()
        upload = get_upload_or_404(upload_id)
        if submitted.get("action") == "skip":
            db.delete_upload(upload_id)
            db.log(user["username"], "upload.skip", upload_id)
            return RedirectResponse("/review/next", status_code=303)

        form = get_form_or_400(upload["form_key"])
        values = normalize(form, {key: submitted.get(f"field_{key}") for key in form.keys})
        context = review_context(upload, form, values=values)
        problems = validate(form, values)
        if problems:
            return render(request, "review.html", {**context, "problems": problems,
                          "error": "Fix the highlighted fields before saving."}, status_code=422)
        if context["uncertain"] and not submitted.get("confirm_uncertain"):
            return render(request, "review.html", {**context, "error":
                          "Tick the box to confirm you checked the hard-to-read fields against the photo."},
                          status_code=422)
        duplicate = db.find_duplicate(form, values)
        if duplicate and not submitted.get("confirm_duplicate"):
            return render(request, "review.html", {**context, "duplicate": duplicate}, status_code=409)
        db.add_entry(form, values, user["username"], source_file=upload["filename"])
        db.delete_upload(upload_id)
        return RedirectResponse("/review/next", status_code=303)

    # --- records ----------------------------------------------------------

    def filtered_entries(form: FormTemplate, q: str, flag: str, start: str, end: str) -> list[dict]:
        entries = db.entries(form.key)
        q = q.strip().lower()
        if q:
            entries = [e for e in entries if any(q in str(v).lower() for v in e["values"].values() if v)]
        if flag and form.flag_field:
            if flag == "*flagged":
                entries = [e for e in entries if e["values"].get(form.flag_field) in form.flag_values]
            else:
                entries = [e for e in entries if e["values"].get(form.flag_field) == flag]
        if form.date_field:
            if start:
                entries = [e for e in entries if (e["values"].get(form.date_field) or "") >= start]
            if end:
                entries = [e for e in entries if (e["values"].get(form.date_field) or "9999") <= end]
        return entries

    def filter_summary(form, q, flag, start, end) -> str:
        parts = [f"{k}={v}" for k, v in (("q", q), ("flag", flag), ("from", start), ("to", end)) if v]
        return "; ".join([form.key] + parts)

    def records_form(form_key: str) -> FormTemplate:
        forms = db.forms()
        if not form_key:
            return forms[0]
        return get_form_or_400(form_key)

    @app.get("/records", response_class=HTMLResponse)
    def records(request: Request, form: str = "", q: str = "", flag: str = "", start: str = "", end: str = "",
                user: dict = Depends(viewer)):
        selected = records_form(form)
        rows = filtered_entries(selected, q, flag, start, end)
        db.log(user["username"], "records.view", detail=filter_summary(selected, q, flag, start, end))
        return render(request, "records.html", {
            "form": selected, "forms": db.forms(), "entries": rows,
            "filters": {"form": selected.key, "q": q, "flag": flag, "start": start, "end": end},
        })

    @app.get("/records/export.xlsx")
    def export(form: str = "", q: str = "", flag: str = "", start: str = "", end: str = "",
               user: dict = Depends(reviewer)):
        selected = records_form(form)
        rows = filtered_entries(selected, q, flag, start, end)
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "export.xlsx")
            db.export_xlsx(path, selected, rows)
            with open(path, "rb") as f:
                data = f.read()
        db.log(user["username"], "records.export", detail=f"{len(rows)} rows; {filter_summary(selected, q, flag, start, end)}")
        filename = f"{selected.name.replace(' ', '')}-{date.today()}.xlsx"
        return Response(data, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        headers={"Content-Disposition": f'attachment; filename="{filename}"'})

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

    # --- admin: form types --------------------------------------------------

    def forms_page(request: Request, error: str | None = None, status_code: int = 200):
        counts = {f.key: db.count(f.key) for f in db.forms()}
        return render(request, "forms.html", {"forms": db.forms(), "counts": counts, "error": error},
                      status_code=status_code)

    def form_editor(request: Request, form: FormTemplate, error: str | None = None, status_code: int = 200):
        return render(request, "form_edit.html", {
            "form": form, "types": FIELD_TYPES, "locked": set(form.keys) if db.count(form.key) else set(),
            "entry_count": db.count(form.key), "error": error,
        }, status_code=status_code)

    @app.get("/admin/forms", response_class=HTMLResponse)
    def list_forms(request: Request, user: dict = Depends(admin)):
        return forms_page(request)

    @app.post("/admin/forms")
    async def create_form(request: Request, name: str = Form(""), copy_from: str = Form(""),
                          user: dict = Depends(admin)):
        await check_csrf(request)
        name = name.strip()
        if not name:
            return forms_page(request, "Give the form type a name.", 422)
        taken = {f.key for f in db.forms()}
        key = make_key(name, taken).replace("_", "-")
        source = db.get_form(copy_from) if copy_from else None
        try:
            form = FormTemplate(
                key=key, name=name, description=source.description if source else "",
                fields=source.fields if source else [FieldDef(key="name", label="Name", required=True)],
                date_field=source.date_field if source else None,
                flag_field=source.flag_field if source else None,
                flag_values=source.flag_values if source else [],
            )
        except ValidationError as e:
            return forms_page(request, friendly(e), 422)
        db.save_form(form, user["username"])
        return RedirectResponse(f"/admin/forms/{key}", status_code=303)

    @app.get("/admin/forms/{key}", response_class=HTMLResponse)
    def edit_form(request: Request, key: str, user: dict = Depends(admin)):
        return form_editor(request, get_form_or_400(key))

    @app.post("/admin/forms/{key}")
    async def save_form(request: Request, key: str, user: dict = Depends(admin)):
        await check_csrf(request)
        existing = get_form_or_400(key)
        submitted = await request.form()
        rows = []
        for i in range(int(submitted.get("field_count", 0)) + 1):  # +1: the blank "add field" row
            label = (submitted.get(f"f{i}_label") or "").strip()
            if not label or submitted.get(f"f{i}_remove"):
                continue
            rows.append((submitted.get(f"f{i}_order") or str(i + 1), i, {
                "key": submitted.get(f"f{i}_key") or None,
                "label": label,
                "type": submitted.get(f"f{i}_type") or "text",
                "required": bool(submitted.get(f"f{i}_required")),
                "duplicate_check": bool(submitted.get(f"f{i}_duplicate")),
                "choices": (submitted.get(f"f{i}_choices") or "").split(","),
                "hint": (submitted.get(f"f{i}_hint") or "").strip(),
            }))

        def order(row):
            try:
                return (float(row[0]), row[1])
            except ValueError:
                return (float(row[1] + 1), row[1])

        fields, taken = [], set()
        for _, _, field in sorted(rows, key=order):
            if field["key"] not in existing.keys:  # new field: derive a key from its label
                field["key"] = make_key(field["label"], taken | set(existing.keys))
            taken.add(field["key"])
            fields.append(field)
        try:
            form = FormTemplate(
                key=existing.key,
                name=(submitted.get("name") or "").strip(),
                description=(submitted.get("description") or "").strip(),
                fields=fields,
                date_field=submitted.get("date_field") or None,
                flag_field=submitted.get("flag_field") or None,
                flag_values=(submitted.get("flag_values") or "").split(","),
            )
            db.save_form(form, user["username"])
        except ValidationError as e:
            return form_editor(request, existing, friendly(e), 422)
        except FormInUse as e:
            return form_editor(request, existing, str(e), 422)
        return RedirectResponse(f"/admin/forms/{key}?saved=1", status_code=303)

    @app.post("/admin/forms/{key}/delete")
    async def delete_form(request: Request, key: str, user: dict = Depends(admin)):
        await check_csrf(request)
        get_form_or_400(key)
        try:
            db.delete_form(key, user["username"])
        except FormInUse as e:
            return forms_page(request, str(e), 422)
        return RedirectResponse("/admin/forms", status_code=303)

    @app.get("/admin/audit", response_class=HTMLResponse)
    def audit(request: Request, user: dict = Depends(admin)):
        return render(request, "audit.html", {"entries": db.audit_entries()})

    return app


def friendly(error: ValidationError) -> str:
    """Turn a pydantic error into a short message for the form editor."""
    messages = []
    for item in error.errors():
        message = item["msg"].removeprefix("Value error, ")
        if item["type"] == "string_pattern_mismatch":
            message = "Labels must contain at least one letter"
        elif item["type"] == "too_short" and item["loc"] and item["loc"][0] == "fields":
            message = "A form type needs at least one field"
        elif item["type"] == "string_too_short" and item["loc"] and item["loc"][0] == "name":
            message = "Give the form type a name"
        messages.append(message)
    return "; ".join(dict.fromkeys(messages))


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
