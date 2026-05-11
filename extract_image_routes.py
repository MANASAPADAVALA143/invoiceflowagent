"""POST /api/agent/extract-image — mounted from agent.py with prefix /api/agent."""

from __future__ import annotations

import base64
import json
import os
from typing import Any

from anthropic import Anthropic
from fastapi import APIRouter, File, HTTPException, UploadFile

router = APIRouter()
_client: Anthropic | None = None


def _anthropic() -> Anthropic:
    global _client
    if _client is None:
        key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
        if not key:
            raise HTTPException(status_code=500, detail="ANTHROPIC_API_KEY is not set")
        _client = Anthropic(api_key=key)
    return _client


def _media_block(content_type: str, data: bytes) -> dict[str, Any]:
    ct = (content_type or "application/octet-stream").lower()
    b64 = base64.standard_b64encode(data).decode("ascii")
    if ct == "application/pdf" or data[:4] == b"%PDF":
        return {
            "type": "document",
            "source": {"type": "base64", "media_type": "application/pdf", "data": b64},
        }
    if not ct.startswith("image/"):
        ct = "image/jpeg"
    return {
        "type": "image",
        "source": {"type": "base64", "media_type": ct, "data": b64},
    }


SYSTEM = """You extract structured data from invoice images or PDF pages.
Return a single JSON object only (no markdown), shape:
{
  "invoice": {
    "invoice_number": string,
    "invoice_date": "YYYY-MM-DD",
    "due_date": "YYYY-MM-DD",
    "vendor_name": string,
    "customer_name": string or "",
    "customer_gstin": string or "",
    "vendor_gstin": string or "",
    "total_amount": number,
    "currency": string (ISO 4217, 3 letters),
    "tax_amount": number or null,
    "invoice_kind": "purchase" or "sales"
  },
  "confidence": number from 0 to 100 (your confidence in extraction accuracy)
}
invoice_kind: purchase = vendor bill (AP), sales = customer bill (AR). Default purchase if unclear.
Use empty string for unknown text fields. Guess due_date as invoice_date + 30 days if missing."""


@router.post("/extract-image")
async def extract_image(file: UploadFile = File(...)) -> dict[str, Any]:
    raw = await file.read()
    if not raw:
        raise HTTPException(status_code=400, detail="Empty file")

    media = _media_block(file.content_type or "", raw)
    client = _anthropic()
    model = os.environ.get("ANTHROPIC_VISION_MODEL", "claude-sonnet-4-20250514")

    try:
        msg = client.messages.create(
            model=model,
            max_tokens=4096,
            system=SYSTEM,
            messages=[
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "Extract the invoice. Respond with JSON only."},
                        media,
                    ],
                }
            ],
        )
    except Exception as e:
        raise HTTPException(status_code=502, detail=str(e)) from e

    text_parts: list[str] = []
    for block in msg.content:
        if hasattr(block, "text"):
            text_parts.append(block.text)
    raw_text = "".join(text_parts).strip()
    if raw_text.startswith("```"):
        raw_text = raw_text.strip("`")
        if raw_text.lower().startswith("json"):
            raw_text = raw_text[4:].lstrip()

    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError:
        raise HTTPException(
            status_code=502,
            detail="Model did not return valid JSON: " + raw_text[:500],
        )

    if "invoice" not in data:
        data = {"invoice": data, "confidence": data.get("confidence", 70)}
    if "confidence" not in data:
        data["confidence"] = 70
    return data
