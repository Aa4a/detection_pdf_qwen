"""Qwen vision adapter for packaging panel detection.

The service expects an OpenAI-compatible endpoint, which is the most common
interface exposed by vLLM, SGLang and other self-hosted Qwen deployments.
"""

from __future__ import annotations

import base64
import io
import json
import os
import re

import pymupdf
from openai import OpenAI
from PIL import Image

Image.MAX_IMAGE_PIXELS = None

QWEN_BASE_URL = os.getenv("QWEN_BASE_URL", "http://127.0.0.1:8000/v1")
QWEN_API_KEY = os.getenv("QWEN_API_KEY", "")
QWEN_MODEL = os.getenv("QWEN_MODEL", "qwen-vl")
QWEN_TIMEOUT = float(os.getenv("QWEN_TIMEOUT", "180"))


def pdf_to_image(pdf_path: str, zoom_factor: int = 2) -> Image.Image:
    """Render the first PDF page to an RGB PIL image."""
    with pymupdf.open(pdf_path) as document:
        pixmap = document[0].get_pixmap(
            matrix=pymupdf.Matrix(zoom_factor, zoom_factor), alpha=False
        )
        return Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)


def _data_url(image: Image.Image) -> str:
    buffer = io.BytesIO()
    image.convert("RGB").save(buffer, format="JPEG", quality=92)
    encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
    return f"data:image/jpeg;base64,{encoded}"


def _extract_json(text: str):
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*|\s*```$", "", text, flags=re.I | re.S)
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}|\[.*\]", text, re.S)
        if not match:
            raise
        return json.loads(match.group(0))


def get_packaging_detections(image_input, requirements: str = "") -> list[dict]:
    """Detect structural panels with a self-hosted Qwen vision model."""
    image = Image.open(image_input) if isinstance(image_input, (str, os.PathLike)) else image_input
    if not isinstance(image, Image.Image):
        raise TypeError("image_input must be a path or PIL.Image.Image")

    custom = f"\nAdditional user requirements (highest priority): {requirements}" if requirements else ""
    prompt = f"""You are a packaging structural engineer. Detect every structural panel
and flap bounded by die-lines, including blank panels. Work from left to right.
Classify wide main panels as 正唛内容 and narrow main panels as 侧唛内容.
Label flaps using their order and type, for example 第一正唛上摇盖、第一侧唛下摇盖.
Boxes must align tightly to physical cut/fold boundaries.{custom}

Return JSON only in this schema:
{{"detections":[{{"label":"标签","box_2d":[ymin,xmin,ymax,xmax]}}]}}
Coordinates are integers normalized to 0..1000 relative to the full image.
"""

    client = OpenAI(base_url=QWEN_BASE_URL, api_key=QWEN_API_KEY, timeout=QWEN_TIMEOUT)
    response = client.chat.completions.create(
        model=QWEN_MODEL,
        temperature=0,
        messages=[{
            "role": "user",
            "content": [
                {"type": "text", "text": prompt},
                {"type": "image_url", "image_url": {"url": _data_url(image)}},
            ],
        }],
    )
    payload = _extract_json(response.choices[0].message.content or "")
    items = payload.get("detections", []) if isinstance(payload, dict) else payload

    results = []
    for item in items:
        label, box = item.get("label"), item.get("box_2d")
        if not label or not isinstance(box, list) or len(box) != 4:
            continue
        ymin, xmin, ymax, xmax = (float(value) for value in box)
        results.append({
            "box_norm": [xmin / 1000, ymin / 1000, xmax / 1000, ymax / 1000],
            "class_name": label,
        })
    return results
