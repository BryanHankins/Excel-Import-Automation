"""Read a photo of a form with a vision AI model and return the fields its form type defines.

Two engines:
  - "ollama" (default): a free, open-source model running on this computer via Ollama
    (https://ollama.com). Photos never leave the machine and there are no per-scan fees.
  - "anthropic": Claude through the paid Anthropic API. Optional; needs `pip install anthropic`
    and ANTHROPIC_API_KEY.

Choose with DRUGTEST_ENGINE; pick the model with DRUGTEST_MODEL.
"""
import base64
import io
import json
import os
import socket
import urllib.error
import urllib.request
from dataclasses import dataclass
from typing import BinaryIO

from PIL import Image, ImageOps

from .forms import FormTemplate, extraction_model, extraction_prompt, extraction_values

ENGINE = os.environ.get("DRUGTEST_ENGINE", "ollama").strip().lower()
DEFAULT_MODELS = {"ollama": "qwen3-vl:8b", "anthropic": "claude-opus-5-5"}
MODEL = os.environ.get("DRUGTEST_MODEL") or DEFAULT_MODELS.get(ENGINE, "")
OLLAMA_URL = os.environ.get("DRUGTEST_OLLAMA_URL", "http://localhost:11434").rstrip("/")
# Local models on a CPU can take minutes per photo
TIMEOUT = float(os.environ.get("DRUGTEST_TIMEOUT", "900"))
# Ollama reads one photo at a time by default, so parallel requests would only queue
PARALLEL = int(os.environ.get("DRUGTEST_PARALLEL") or (1 if ENGINE == "ollama" else 3))
MAX_SIDE = int(os.environ.get("DRUGTEST_MAX_SIDE", "1568"))


class ExtractionError(Exception):
    pass


@dataclass
class Reading:
    """What the model read from one image."""
    values: dict  # field key -> value or None
    uncertain: list[str]  # keys of fields that were hard to read
    model: str = ""
    input_tokens: int = 0
    output_tokens: int = 0  # includes thinking


def encode_image(image: str | BinaryIO) -> tuple[str, str]:
    """Return (media_type, base64 data) for a JPEG version of the image, auto-rotated and resized.

    `image` is a file path or a binary file object. Works in memory so no temp
    files with sensitive content are left on disk.
    """
    with Image.open(image) as img:
        img = ImageOps.exif_transpose(img).convert("RGB")
        img.thumbnail((MAX_SIDE, MAX_SIDE), Image.Resampling.LANCZOS)
        buffer = io.BytesIO()
        img.save(buffer, format="JPEG", quality=90)
    return "image/jpeg", base64.standard_b64encode(buffer.getvalue()).decode("ascii")


def extract_fields(image: str | BinaryIO, form: FormTemplate, client=None) -> tuple[dict, list[str]]:
    """Read the form in the image. Returns ({field key: value or None}, keys of hard-to-read fields)."""
    reading = read_form(image, form, client)
    return reading.values, reading.uncertain


def read_form(image: str | BinaryIO, form: FormTemplate, client=None) -> Reading:
    """Like extract_fields, but also reports which model answered and the tokens used.

    `client` is an Anthropic client, only used (and only needed in tests) for the anthropic engine.
    """
    if ENGINE == "ollama":
        return _read_with_ollama(image, form)
    if ENGINE == "anthropic":
        return _read_with_anthropic(image, form, client)
    raise ExtractionError(f"Unknown DRUGTEST_ENGINE {ENGINE!r}. Use 'ollama' or 'anthropic'.")


# --- Ollama (local, free) -------------------------------------------------------

def ollama_schema(form: FormTemplate) -> dict:
    """JSON schema for the answer. Plain strings (empty = blank) keep it simple for small local models."""
    described = extraction_model(form).model_json_schema()["properties"]
    properties = {key: {"type": "string", "description": described[key].get("description", "")} for key in form.keys}
    properties["uncertain_fields"] = {"type": "array", "items": {"type": "string", "enum": form.keys}}
    return {"type": "object", "properties": properties, "required": list(properties)}


def _clean(value) -> str | None:
    if value is None:
        return None
    value = str(value).strip()
    return None if value.lower() in ("", "null", "none") else value


def _ollama_request(path: str, body: dict | None = None, timeout: float = TIMEOUT) -> dict:
    request = urllib.request.Request(
        OLLAMA_URL + path, data=json.dumps(body).encode() if body is not None else None,
        headers={"Content-Type": "application/json"}, method="POST" if body is not None else "GET")
    try:
        with urllib.request.urlopen(request, timeout=timeout) as response:
            return json.load(response)
    except urllib.error.HTTPError as e:
        try:
            message = json.load(e).get("error", "")
        except (ValueError, AttributeError):
            message = ""
        if e.code == 404 or "not found" in message:
            raise ExtractionError(f"The model {MODEL!r} isn't downloaded. Run: ollama pull {MODEL}") from e
        raise ExtractionError(f"Ollama error ({e.code}): {message or e.reason}") from e
    except (TimeoutError, socket.timeout) as e:
        raise ExtractionError(f"The model took longer than {TIMEOUT:.0f} seconds. Try a smaller model or "
                              "set DRUGTEST_TIMEOUT higher.") from e
    except urllib.error.URLError as e:
        if isinstance(e.reason, (TimeoutError, socket.timeout)):
            raise ExtractionError(f"The model took longer than {TIMEOUT:.0f} seconds.") from e
        raise ExtractionError(f"Can't reach Ollama at {OLLAMA_URL}. Is it installed and running? "
                              "See https://ollama.com") from e


def _read_with_ollama(image: str | BinaryIO, form: FormTemplate) -> Reading:
    _, data = encode_image(image)
    prompt = extraction_prompt(form) + (
        "\n\nAnswer only with JSON. Use an empty string for any field that is blank or unreadable.")
    payload = _ollama_request("/api/chat", {
        "model": MODEL,
        "stream": False,
        "format": ollama_schema(form),
        "options": {"temperature": 0},
        "messages": [{"role": "user", "content": prompt, "images": [data]}],
    })
    content = (payload.get("message") or {}).get("content", "")
    try:
        answer = json.loads(content)
        if not isinstance(answer, dict):
            raise ValueError
    except ValueError as e:
        raise ExtractionError("The model's answer wasn't valid JSON. Try again or enter the fields manually.") from e
    uncertain = answer.get("uncertain_fields") or []
    return Reading(
        values={key: _clean(answer.get(key)) for key in form.keys},
        uncertain=[key for key in uncertain if key in form.keys] if isinstance(uncertain, list) else [],
        model=payload.get("model") or MODEL,
        input_tokens=payload.get("prompt_eval_count") or 0,
        output_tokens=payload.get("eval_count") or 0,
    )


def check_ollama() -> str:
    """Return a status message, or raise ExtractionError explaining what to fix."""
    models = _ollama_request("/api/tags", timeout=10).get("models", [])
    names = {m.get("name", "") for m in models}
    wanted = MODEL if ":" in MODEL else f"{MODEL}:latest"
    if wanted not in names and MODEL not in names:
        raise ExtractionError(f"Ollama is running, but the model {MODEL!r} isn't downloaded. Run: ollama pull {MODEL}")
    return f"Ollama is running at {OLLAMA_URL} and {MODEL} is ready."


# --- Anthropic API (optional, paid) ---------------------------------------------

def _read_with_anthropic(image: str | BinaryIO, form: FormTemplate, client=None) -> Reading:
    try:
        import anthropic
    except ImportError as e:
        raise ExtractionError("The anthropic engine needs: pip install anthropic") from e
    client = client or anthropic.Anthropic()
    media_type, data = encode_image(image)
    try:
        response = client.beta.messages.parse(
            model=MODEL,
            max_tokens=4096,
            output_config={"effort": "medium"},
            betas=["server-side-fallback-2026-07-01"],
            fallbacks="default",
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64", "media_type": media_type, "data": data}},
                    {"type": "text", "text": extraction_prompt(form)},
                ],
            }],
            output_format=extraction_model(form),
        )
    except anthropic.AuthenticationError as e:
        raise ExtractionError("Invalid API key. Set ANTHROPIC_API_KEY.") from e
    except anthropic.RateLimitError as e:
        raise ExtractionError("Rate limited by the API. Wait a minute and try again.") from e
    except anthropic.APIStatusError as e:
        raise ExtractionError(f"API error ({e.status_code}): {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise ExtractionError("Could not reach the API. Check your internet connection.") from e

    if response.stop_reason == "refusal":
        raise ExtractionError("The model declined to read this image. Enter the fields manually.")
    if response.stop_reason == "max_tokens" or response.parsed_output is None:
        raise ExtractionError("The model returned an incomplete answer. Try again or enter fields manually.")
    values, uncertain = extraction_values(form, response.parsed_output)
    usage = getattr(response, "usage", None)
    return Reading(
        values=values, uncertain=uncertain, model=getattr(response, "model", "") or MODEL,
        input_tokens=(getattr(usage, "input_tokens", 0) or 0) + (getattr(usage, "cache_read_input_tokens", 0) or 0)
                     + (getattr(usage, "cache_creation_input_tokens", 0) or 0),
        output_tokens=getattr(usage, "output_tokens", 0) or 0,
    )
