import json
import os
from typing import Any, Dict, List

from openai import OpenAI


"""
Usage:
  python3 -m venv .venv
  source .venv/bin/activate
  pip install --upgrade openai
  # Uses .env automatically if present (OPENAI_API_KEY / OPENAI_MODEL).
  python semantic_cluster_test.py

This script reads ./fields.json and asks the LLM to cluster semantically equivalent fields
into canonical meanings (canonical_id). This is the "semantic interconnectedness" test.
"""


MODEL = os.getenv("OPENAI_MODEL", "gpt-4o-mini")

SYSTEM = """You cluster fields from a legal document into canonical meanings so semantically equivalent blanks share the same value.

Return ONLY valid JSON: an object with:
{
  "clusters": [
    {
      "canonical_id": "string",
      "reason": "string",
      "member_field_ids": ["string", ...]
    }
  ]
}

Rules:
- Do NOT merge fields that mean different money concepts (e.g., investment/purchase amount vs valuation cap).
- Only merge if expected_type matches AND the evidence/context clearly points to the same meaning.
- Use stable canonical_id values, like:
  - money.investment_amount
  - money.valuation_cap
  - party.company.name
  - party.investor.name
  - legal.governing_law
- If a field is unique, still include it as a cluster with one member.
"""


def main() -> None:
    # Load .env if present (minimal parser; no extra dependency).
    if os.path.exists(".env"):
        try:
            with open(".env", "r", encoding="utf-8") as f:
                for line in f:
                    s = line.strip()
                    if not s or s.startswith("#"):
                        continue
                    if "=" not in s:
                        continue
                    k, v = s.split("=", 1)
                    k = k.strip()
                    v = v.strip().strip('"').strip("'")
                    # Don't override already-set env vars.
                    if k and k not in os.environ:
                        os.environ[k] = v
        except Exception:
            pass

    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        raise SystemExit("Missing OPENAI_API_KEY")

    with open("fields.json", "r", encoding="utf-8") as f:
        fields: List[Dict[str, Any]] = json.load(f)

    payload: List[Dict[str, Any]] = []
    for fo in fields:
        payload.append(
            {
                "field_id": fo.get("field_id"),
                "label_for_user": fo.get("label_for_user"),
                "expected_type": fo.get("expected_type"),
                "who_should_fill": fo.get("who_should_fill"),
                "evidence_quote": fo.get("evidence_quote"),
                "location_hint": fo.get("location_hint"),
                "raw_placeholder": fo.get("raw_placeholder"),
            }
        )

    client = OpenAI(api_key=api_key)
    resp = client.chat.completions.create(
        model=MODEL,
        temperature=0.1,
        messages=[
            {"role": "system", "content": SYSTEM},
            {"role": "user", "content": json.dumps({"fields": payload}, indent=2)},
        ],
    )

    raw = (resp.choices[0].message.content or "").strip()
    try:
        out = json.loads(raw)
    except Exception:
        print("MODEL OUTPUT (not JSON):")
        print(raw)
        raise

    print(json.dumps(out, indent=2))

    clusters = out.get("clusters", [])
    multi = [c for c in clusters if len(c.get("member_field_ids", [])) > 1]
    print(f"\nClusters with >1 member: {len(multi)}")
    for c in multi:
        print(f"- {c.get('canonical_id')}: {c.get('member_field_ids')}")


if __name__ == "__main__":
    main()


