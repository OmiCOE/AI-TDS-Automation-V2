"""
llm_judge.py  -  optional AI step that READS the shortlisted datasheets
======================================================================
The free ranker (tds_matcher) shortlists the ~8 most likely TDS files for a BOQ
line. This module shows the BOQ line and those datasheets' opening text to a
language model and asks it to do what an engineer would:
  * pick the datasheet that is really the product described - or say none fits,
  * read the model number and manufacturer from the datasheet.

Providers:
  "anthropic"  - Anthropic Messages API (api.anthropic.com)
  "openai"     - any OpenAI-compatible /chat/completions endpoint (many free-tier
                 services, local servers, etc.) - give base_url + model.
Only short excerpts (about 800 characters per candidate) are sent, never whole PDFs.
"""

import json
import re
import time

import requests

SYSTEM = (
    "You are a technical-submittal engineer. You match a Bill-of-Quantities (BOQ) line item to the "
    "correct manufacturer datasheet (TDS). Be strict: choose a datasheet only if it is the product the "
    "BOQ line describes (same product type and compatible key specifications). A datasheet from a "
    "similar category or a different product is NOT a match. If none of the candidates fits, choose 0."
)

PROMPT = """BOQ LINE ITEM
Description: {description}
Make (may be blank/NA): {make}

CANDIDATE DATASHEETS
{candidates}

Reply with ONLY a JSON object, no other text:
{{"choice": <number of the matching candidate, or 0 if none fits>,
  "confidence": "high" | "medium" | "low",
  "reason": "<one short sentence>",
  "model": "<model number / product name as printed on that datasheet, else empty>",
  "make": "<manufacturer shown on that datasheet, else empty>"}}"""


def parse_json(text):
    m = re.search(r"\{.*\}", text or "", re.S)
    if not m:
        raise ValueError("no JSON in model reply")
    d = json.loads(m.group(0))
    try:
        d["choice"] = int(d.get("choice") or 0)
    except (TypeError, ValueError):
        d["choice"] = 0
    return d


class Judge:
    def __init__(self, provider, api_key, model, base_url=None, session=None, timeout=60):
        self.provider = provider
        self.api_key = api_key
        self.model = model
        self.base_url = (base_url or "").rstrip("/")
        self.session = session or requests.Session()
        self.timeout = timeout

    # ---- transport -----------------------------------------------------------------
    def _post(self, url, headers, body):
        last = None
        for attempt in range(4):
            r = self.session.post(url, headers=headers, json=body, timeout=self.timeout)
            if r.status_code in (429, 500, 502, 503, 529):
                last = f"HTTP {r.status_code}"
                time.sleep(2 * (attempt + 1))
                continue
            if r.status_code >= 400:
                raise RuntimeError(f"HTTP {r.status_code}: {r.text[:200]}")
            return r.json()
        raise RuntimeError(f"model service busy ({last})")

    def _complete(self, prompt):
        if self.provider == "anthropic":
            data = self._post(
                "https://api.anthropic.com/v1/messages",
                {"x-api-key": self.api_key, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                {"model": self.model, "max_tokens": 400, "system": SYSTEM,
                 "messages": [{"role": "user", "content": prompt}]})
            return "".join(b.get("text", "") for b in data.get("content", []))
        data = self._post(
            f"{self.base_url}/chat/completions",
            {"Authorization": f"Bearer {self.api_key}", "content-type": "application/json"},
            {"model": self.model, "temperature": 0, "max_tokens": 400,
             "messages": [{"role": "system", "content": SYSTEM}, {"role": "user", "content": prompt}]})
        return data["choices"][0]["message"]["content"]

    # ---- the decision -----------------------------------------------------------------
    def decide(self, row, docs):
        cands = "\n".join(
            f"[{i}] file: {d.name} | folder: {d.folder or '-'}\n    text: {d.excerpt or '(scanned PDF - no text; judge by filename only)'}"
            for i, d in enumerate(docs, 1))
        prompt = PROMPT.format(description=(row.get("Description") or "").strip()[:1500],
                               make=(row.get("Make") or "").strip() or "-", candidates=cands)
        out = parse_json(self._complete(prompt))
        if out["choice"] < 0 or out["choice"] > len(docs):
            out["choice"] = 0
        return out

    def ping(self):
        """Cheap connectivity/key check for the UI."""
        self._complete('Reply with the single word OK.')
        return True
