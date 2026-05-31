"""
Unified LLM caller — Gemini only (OpenAI paths are commented out).

Public API:
    call_model(model_id, prompt) -> raw response text
    parse_json(text) -> dict   (with fallback for trailing junk)

Tunables:
    DEFAULT_MODELS    which models to run by default for multi-model studies
    MAX_RETRIES       per-call retry budget
    RETRY_BACKOFF     base seconds; doubles per attempt
"""

from __future__ import annotations

import json
import os
import re
import time
from typing import Optional

# default model list (OpenAI models disabled — Gemini only)
DEFAULT_MODELS = [
    "gemini-3.5-flash",
    # "gpt-5.4",
    # "gpt-5.4-mini",
]

MAX_RETRIES = 5
RETRY_BACKOFF = 2.0


def _is_gemini(model_id: str) -> bool:
    return model_id.startswith("gemini")


# def _is_openai(model_id: str) -> bool:
#     return model_id.startswith("gpt-")


def call_gemini(model_id: str, prompt: str) -> str:
    """Call Google Gemini with retries. Returns raw JSON-formatted text."""
    from google import genai
    from google.genai import types

    api_key = os.environ.get("GEMINI_API_KEY", "")
    if not api_key:
        raise RuntimeError("GEMINI_API_KEY not set")

    client = genai.Client(api_key=api_key)
    last_err: Optional[Exception] = None
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            resp = client.models.generate_content(
                model=model_id,
                contents=prompt,
                config=types.GenerateContentConfig(
                    response_mime_type="application/json",
                ),
            )
            return resp.text or ""
        except Exception as e:
            last_err = e
            print(f"    {model_id} attempt {attempt}/{MAX_RETRIES}: {e}")
            if attempt < MAX_RETRIES:
                time.sleep(RETRY_BACKOFF * attempt)
    raise RuntimeError(f"{model_id} failed after {MAX_RETRIES} retries: {last_err}")


# OpenAI backend disabled — Gemini-only run.
# def call_openai(model_id: str, prompt: str) -> str:
#     """Call an OpenAI chat completion with retries. Returns raw JSON-formatted text."""
#     from openai import OpenAI
#
#     api_key = os.environ.get("OPENAI_API_KEY", "")
#     if not api_key:
#         raise RuntimeError("OPENAI_API_KEY not set")
#
#     client = OpenAI(api_key=api_key)
#     last_err: Optional[Exception] = None
#     for attempt in range(1, MAX_RETRIES + 1):
#         try:
#             kwargs: dict = {
#                 "model": model_id,
#                 "messages": [{"role": "user", "content": prompt}],
#                 "response_format": {"type": "json_object"},
#             }
#             # gpt-5.x is a reasoning model. leave max_completion_tokens unset
#             # so the model decides when to stop (matches ChatGPT website behavior).
#             # for older non-reasoning models keep a sane cap.
#             if not model_id.startswith("gpt-5"):
#                 kwargs["max_tokens"] = 1500
#             resp = client.chat.completions.create(**kwargs)
#             return (resp.choices[0].message.content or "").strip()
#         except Exception as e:
#             last_err = e
#             print(f"    {model_id} attempt {attempt}/{MAX_RETRIES}: {e}")
#             if attempt < MAX_RETRIES:
#                 time.sleep(RETRY_BACKOFF * attempt)
#     raise RuntimeError(f"{model_id} failed after {MAX_RETRIES} retries: {last_err}")


def call_model(model_id: str, prompt: str) -> str:
    """Dispatch to the right backend based on the model id prefix."""
    if _is_gemini(model_id):
        return call_gemini(model_id, prompt)
    if _is_openai(model_id):
        return call_openai(model_id, prompt)
    raise ValueError(f"Unrecognized model id: {model_id!r}")


def parse_json(text: str) -> dict:
    """Robust JSON parse with fallbacks for the common Gemini/GPT quirks.

    Handles:
      - markdown code fences (```json ... ```)
      - trailing junk after a valid JSON object (uses raw_decode)
      - array-wrapped single object: [{...}] -> {...}
    Returns {} on total failure.
    """
    text = (text or "").strip()
    if not text:
        return {}
    # strip code fences
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text)
        text = re.sub(r"\s*```$", "", text)
    # primary parse
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        # try parsing just the first valid JSON value at the start
        try:
            decoder = json.JSONDecoder()
            data, _ = decoder.raw_decode(text)
        except Exception:
            return {}
    # unwrap single-element list (some Gemini runs do this)
    if isinstance(data, list) and len(data) == 1 and isinstance(data[0], dict):
        data = data[0]
    return data if isinstance(data, dict) else {}


def format_evidence(evidence) -> str:
    """Flatten a list of {quote, source} dicts to a single string for an Excel cell."""
    if not evidence:
        return ""
    if isinstance(evidence, list):
        parts = []
        for e in evidence:
            if isinstance(e, dict):
                q = e.get("quote", "")
                s = e.get("source", "")
                parts.append(f'"{q}"  [{s}]' if s else f'"{q}"')
            else:
                parts.append(str(e))
        return "\n---\n".join(parts)
    return str(evidence)
