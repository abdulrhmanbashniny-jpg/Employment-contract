# -*- coding: utf-8 -*-
import json
import re
import time
from dataclasses import dataclass
from typing import Dict, List, Tuple, Any, Optional

import requests

PERPLEXITY_CHAT_URL = "https://api.perplexity.ai/chat/completions"  # official :contentReference[oaicite:3]{index=3}

@dataclass
class AIResult:
    values: Dict[str, str]
    evidence: Dict[str, str]
    confidence: Dict[str, float]
    raw_text: str
    error: str = ""

def _extract_json_block(text: str) -> Optional[dict]:
    """
    Robust JSON extraction: accepts pure JSON or JSON wrapped in text.
    """
    if not text:
        return None

    text = text.strip()

    # direct JSON
    if text.startswith("{") and text.endswith("}"):
        try:
            return json.loads(text)
        except Exception:
            pass

    # try to find the first {...} block
    m = re.search(r"\{.*\}", text, flags=re.DOTALL)
    if not m:
        return None
    candidate = m.group(0)
    try:
        return json.loads(candidate)
    except Exception:
        return None

def _safe_float(x, default=0.0) -> float:
    try:
        return float(x)
    except Exception:
        return default

def _post_perplexity(api_key: str, payload: dict, timeout_s: int = 60) -> dict:
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    r = requests.post(PERPLEXITY_CHAT_URL, headers=headers, json=payload, timeout=timeout_s)
    r.raise_for_status()
    return r.json()

def ai_fill_missing_fields(
    api_key: str,
    model: str,
    normalized_text: str,
    missing_fields: List[str],
    headers_all: List[str],
    max_chars: int = 22000,
    temperature: float = 0.0,
    retry: int = 2,
) -> AIResult:
    """
    Ask AI ONLY for missing fields. Returns:
      values[field] = extracted value
      evidence[field] = short snippet proving it
      confidence[field] = 0..1
    """
    if not api_key:
        return AIResult(values={}, evidence={}, confidence={}, raw_text="", error="Missing PERPLEXITY_API_KEY")

    text = (normalized_text or "").strip()
    if len(text) > max_chars:
        text = text[:max_chars]  # keep it safe for token limits

    # Keep the AI strictly on schema + Arabic values + required formatting
    # (We want numbers only where required, and DD/MM/YYYY for dates)
    schema_fields = missing_fields[:]  # only missing
    schema_fields = [f for f in schema_fields if f in headers_all]

    prompt = f"""
أنت وكيل استخراج بيانات عقود عمل سعودية (قوى). 
لديك نص عقد (بعد التطبيع)، وأحتاج منك فقط تعبئة الحقول الناقصة التالية بدقة:

{schema_fields}

قواعد الإخراج:
- أعد JSON فقط (بدون شرح).
- المفاتيح يجب أن تكون EXACT نفس أسماء الحقول أعلاه.
- أي حقل غير موجود بالنص: اجعله "".
- التواريخ بصيغة DD/MM/YYYY فقط.
- الحقول المالية/العددية: أرقام فقط بدون فواصل/رموز/عملة.
- رقم الجوال: بدون مسافات، وإذا بدأ بـ 9660 احذف الصفر بعد 966 (9665....).
- أجر الساعة الإضافية: رقم فقط (مثال 50).
- مدة العقد: رقم فقط (1 للسنة، 6 لستة أشهر...).
- أضف كائنات إضافية داخل JSON:
  - "_evidence": قاموس (field -> مقتطف قصير من النص يثبت القيمة)
  - "_confidence": قاموس (field -> رقم من 0 إلى 1)

نص العقد:
\"\"\"{text}\"\"\"
""".strip()

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "أنت مساعد دقيق للغاية لاستخراج الحقول بشكل منظم بصيغة JSON فقط."},
            {"role": "user", "content": prompt},
        ],
        "temperature": temperature,
    }

    last_err = ""
    for attempt in range(retry + 1):
        try:
            data = _post_perplexity(api_key, payload)
            content = data["choices"][0]["message"]["content"]
            obj = _extract_json_block(content)
            if not obj or not isinstance(obj, dict):
                return AIResult(values={}, evidence={}, confidence={}, raw_text=content, error="AI returned non-JSON")

            evidence = obj.pop("_evidence", {}) or {}
            conf = obj.pop("_confidence", {}) or {}

            values: Dict[str, str] = {}
            for f in schema_fields:
                v = obj.get(f, "")
                values[f] = "" if v is None else str(v).strip()

            evidence_map = {k: str(v).strip() for k, v in (evidence.items() if isinstance(evidence, dict) else [])}
            conf_map = {k: _safe_float(v, 0.0) for k, v in (conf.items() if isinstance(conf, dict) else [])}

            return AIResult(values=values, evidence=evidence_map, confidence=conf_map, raw_text=content)

        except Exception as e:
            last_err = f"{type(e).__name__}: {e}"
            time.sleep(0.6)

    return AIResult(values={}, evidence={}, confidence={}, raw_text="", error=last_err)

def merge_row_with_ai(row: Dict[str, str], ai: AIResult, only_fill_empty: bool = True) -> Tuple[Dict[str, str], int]:
    """
    Merge AI values into row. Returns (row, ai_filled_count)
    """
    filled = 0
    for k, v in ai.values.items():
        if not v:
            continue
        if only_fill_empty and str(row.get(k, "")).strip():
            continue
        row[k] = v
        filled += 1
    return row, filled
