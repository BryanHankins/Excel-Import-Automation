"""Read a photo of a form with Claude's vision model and return the fields its form type defines."""
import base64
import io
import os
from dataclasses import dataclass
from typing import BinaryIO

import anthropic
from PIL import Image, ImageOps

from .forms import FormTemplate, extraction_model, extraction_prompt, extraction_values

MODEL = os.environ.get("DRUGTEST_MODEL", "claude-opus-5-5")
MAX_SIDE = 1568  # larger images are downscaled by the API anyway


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


def extract_fields(image: str | BinaryIO, form: FormTemplate,
                   client: anthropic.Anthropic | None = None) -> tuple[dict, list[str]]:
    """Read the form in the image. Returns ({field key: value or None}, keys of hard-to-read fields)."""
    reading = read_form(image, form, client)
    return reading.values, reading.uncertain


def read_form(image: str | BinaryIO, form: FormTemplate, client: anthropic.Anthropic | None = None) -> Reading:
    """Like extract_fields, but also reports which model answered and the tokens used."""
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
