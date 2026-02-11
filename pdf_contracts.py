# -*- coding: utf-8 -*-
import re
import io
import unicodedata
import pdfplumber

# =======================
#  Columns / Headers
# =======================
HEADERS = [
    "رقم العقد","تاريخ العقد","شركة/مؤسسة","الرقم الوطني الموحد","رقم المنشأة","السجل التجاري","عنوان الشركة","مكان العمل",
    "بريد الشركة","المسؤول الموقع","الصفة","اسم الموظف","رقم الهوية","نوع الهوية","تاريخ الميلاد","تاريخ انتهاء الهوية",
    "الجنسية","الجنس","الديانة","الحالة الاجتماعية","المؤهل العلمي","التخصص","المهنة","الرقم الوظيفي","رقم الآيبان",
    "اسم البنك","بريد الموظف","رقم الجوال","بدء العقد","انتهاء العقد","تاريخ المباشرة الفعلية","مدة العقد","فترة التجربة",
    "أيام العمل الأسبوعية","ساعات العمل اليومية","الراتب الأساسي","بدل السكن","الإجازة السنوية","أجر الساعة الإضافية",
    "التعويض عند الفسخ بدون سبب"
]

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")
AR_RE = re.compile(r"[\u0600-\u06FF]")
LAT_RE = re.compile(r"[A-Za-z@]")

# =======================
#  Basic Utils
# =======================
def normalize_text(t: str) -> str:
    t = unicodedata.normalize("NFKC", t or "")
    t = t.replace("\u200f", "").replace("\u200e", "")
    return t

def digits_only(s: str) -> str:
    if not s:
        return ""
    return "".join(re.findall(r"\d+", str(s)))

def clean_amount_to_int(s: str) -> str:
    """
    2,000.00 -> 2000
    """
    if not s:
        return ""
    m = re.search(r"(\d[\d,]*)(?:\.\d+)?", str(s))
    if not m:
        return ""
    return m.group(1).replace(",", "")

def format_date_any(s: str) -> str:
    """
    Accepts: YYYY-MM-DD / DD-MM-YYYY / YYYY/MM/DD / DD/MM/YYYY
    Returns: DD/MM/YYYY
    """
    if not s:
        return ""
    s = str(s).strip()
    m = re.search(r"(\d{1,4})[-/](\d{1,2})[-/](\d{1,4})", s)
    if not m:
        return ""

    a, b, c = m.group(1), m.group(2), m.group(3)

    # Year first
    if len(a) == 4:
        year = int(a); month = int(b); day = int(c)
    else:
        day = int(a); month = int(b); year = int(c)
        if year < 100:
            year += 2000

    if not (1 <= month <= 12 and 1 <= day <= 31):
        return ""  # invalid

    return f"{day:02d}/{month:02d}/{year:04d}"

def fix_rtl_value(v: str) -> str:
    """
    Fix values that come reversed (Arabic text) but keep emails/latin/numbers.
    """
    if not v:
        return ""
    v = str(v).strip()
    if "@" in v or re.search(r"[A-Za-z]", v):
        return v
    # if mostly Arabic, reverse it
    ar = len(AR_RE.findall(v))
    lat = len(LAT_RE.findall(v))
    if ar >= 3 and ar > lat:
        return v[::-1]
    return v

# =======================
#  RTL Text Normalization (IMPORTANT)
# =======================
def _ar_count(s: str) -> int:
    return len(AR_RE.findall(s))

def _lat_count(s: str) -> int:
    return len(LAT_RE.findall(s))

def _dig_count(s: str) -> int:
    return len(re.findall(r"\d", s))

def smart_normalize_line(line: str) -> str:
    """
    Converts Qiwa-style extracted lines like:
        22477445 :دقعلا مقر
    to:
        رقم العقد: 22477445

    And:
        a@b.com :ينورتكللإا ديربلا
    to:
        البريد اإللكتروني: a@b.com

    Also:
        NAME :مسلاا
    to:
        االسم: NAME
    """
    line = (line or "").strip()
    if not line:
        return ""

    if ":" not in line:
        return line

    left, right = line.split(":", 1)
    left = left.strip()
    right = right.strip()

    # Heuristic: if right looks like a reversed Arabic label -> reverse it and swap
    if _ar_count(right) > 0 and _lat_count(right) == 0 and _dig_count(right) < 3:
        label = right[::-1].strip()
        value = left.strip()
        return f"{label}: {value}"

    # Or if left looks like reversed Arabic label
    if _ar_count(left) > 0 and _lat_count(left) == 0 and _dig_count(left) < 3:
        label = left[::-1].strip()
        value = right.strip()
        return f"{label}: {value}"

    return line

def normalize_contract_text(raw_text: str) -> str:
    """
    Normalize extracted PDF text:
    - Fix "value : reversed-label" lines into "label: value"
    - Keep sentences as-is
    """
    lines = []
    for ln in (raw_text or "").splitlines():
        ln = normalize_text(ln)
        ln2 = smart_normalize_line(ln)
        if ln2:
            lines.append(ln2)
    return "\n".join(lines)

# =======================
#  PDF Extraction
# =======================
def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    """
    Streamlit Cloud fix: pdfplumber needs file-like with seek()
    + normalize RTL quirks
    """
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(normalize_text(page.extract_text() or ""))
    raw = "\n".join(parts)
    return normalize_contract_text(raw)

# =======================
#  Label Getters
# =======================
def get_value_after_label(text: str, label: str) -> str:
    m = re.search(rf"{re.escape(label)}\s*:\s*([^\n]+)", text)
    return m.group(1).strip() if m else ""

def get_bi(text: str, labels) -> str:
    for lab in labels:
        v = get_value_after_label(text, lab)
        if v:
            return v
    return ""

def find_line_containing(text: str, keyword: str) -> str:
    for ln in text.splitlines():
        if keyword in ln:
            return ln.strip()
    return ""

def normalize_mobile_from_line(line: str) -> str:
    """
    Input examples:
    - "رقم الجوال: 966 0590728938"
    - "رقم الجوال: 9660550266101"
    Output:
    - "966590728938"  (remove spaces, remove 0 after 966)
    """
    if not line:
        return ""
    groups = re.findall(r"\d+", line)
    if not groups:
        return ""

    has_966 = any(g == "966" or g.startswith("966") for g in groups)

    # local number
    local = ""
    for g in groups:
        if len(g) == 10 and g.startswith("05"):
            local = g
            break
        if len(g) == 9 and g.startswith("5"):
            local = "0" + g
            break

    if has_966 and local:
        return "966" + local[1:]  # remove 0 after 966

    joined = "".join(groups)
    if joined.startswith("9660"):
        joined = "966" + joined[4:]
    return joined

# =======================
#  Strong Parser (Qiwa-style)
# =======================
def parse_contract(text: str) -> dict:
    out = {h: "" for h in HEADERS}
    if not text:
        return out

    # ---- Contract No / Date ----
    m = re.search(r"رقم العقد\s*:\s*(\d+)", text)
    if m:
        out["رقم العقد"] = m.group(1)

    # "في يوم (2024-09-21)" sometimes embedded in sentence
    m = re.search(r"في يوم.*?\(?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})\s*\)?", text)
    if m:
        out["تاريخ العقد"] = format_date_any(m.group(1))

    # ---- Company (Party 1) ----
    out["شركة/مؤسسة"] = fix_rtl_value(get_bi(text, ["شركة/مؤسسة", "Corporation/Company"]))
    out["الرقم الوطني الموحد"] = digits_only(get_bi(text, ["الرقم الوطني الموحد", "National Unified Number"]))
    out["رقم المنشأة"] = get_bi(text, ["رقم المنشأة", "Establishment Number"]).strip()
    out["السجل التجاري"] = digits_only(get_bi(text, ["السجل التجاري", "Commercial Registration"]))
    out["عنوان الشركة"] = fix_rtl_value(get_bi(text, ["العنوان", "Address"]))
    out["مكان العمل"] = fix_rtl_value(get_bi(text, ["مكان العمل", "Work Location"]))

    # Emails: prefer first email as company email, second as employee email (if present)
    emails = EMAIL_RE.findall(text)
    if emails:
        out["بريد الشركة"] = emails[0]
        if len(emails) >= 2:
            out["بريد الموظف"] = emails[1]

    # Signatory: "ويمثلها بالتوقيع: فلان بصفته كذا"
    sign = get_bi(text, ["ويمثلها بالتوقيع"])
    if sign:
        if "بصفته" in sign:
            left, right = sign.split("بصفته", 1)
            out["المسؤول الموقع"] = fix_rtl_value(left.strip())
            out["الصفة"] = fix_rtl_value(right.strip())
        else:
            out["المسؤول الموقع"] = fix_rtl_value(sign.strip())

    # ---- Employee (Party 2) ----
    out["اسم الموظف"] = get_bi(text, ["االسم", "الاسم", "Name"]).strip()
    out["المهنة"] = fix_rtl_value(get_bi(text, ["المهنة", "Profession"]))
    out["الرقم الوظيفي"] = digits_only(get_bi(text, ["الرقم الوظيفي", "Employee Number"]))
    out["الجنسية"] = fix_rtl_value(get_bi(text, ["الجنسية", "Nationality"]))
    out["تاريخ الميلاد"] = format_date_any(get_bi(text, ["تاريخ الميالد", "Date of Birth"]))
    out["رقم الهوية"] = digits_only(get_bi(text, ["رقم الهوية", "Identity Number"]))
    out["نوع الهوية"] = fix_rtl_value(get_bi(text, ["نوع الهوية", "ID Type"]))
    out["تاريخ انتهاء الهوية"] = format_date_any(get_bi(text, ["تاريخ اإلنتهاء", "تاريخ الانتهاء", "ID Expiry Date"]))
    out["الجنس"] = fix_rtl_value(get_bi(text, ["الجنس", "Gender"]))
    out["الديانة"] = fix_rtl_value(get_bi(text, ["الديانة", "Religion"]))
    out["الحالة الاجتماعية"] = fix_rtl_value(get_bi(text, ["الحالة االجتماعية", "الحالة الاجتماعية", "Marital Status"]))
    out["المؤهل العلمي"] = fix_rtl_value(get_bi(text, ["المؤهل العلمي", "Education"]))
    out["التخصص"] = fix_rtl_value(get_bi(text, ["التخصص", "Speciality"]))

    iban = get_bi(text, ["رقم اآليبان", "رقم الآيبان", "Iban"])
    out["رقم الآيبان"] = re.sub(r"\s+", "", iban).strip()
    out["اسم البنك"] = fix_rtl_value(get_bi(text, ["اسم البنك", "Bank Name"]))

    mobile_line = find_line_containing(text, "رقم الجوال")
    if mobile_line:
        out["رقم الجوال"] = normalize_mobile_from_line(mobile_line)

    # ---- Terms ----
    # "مدة هذا العقد 1 سنة ... يبدأ من تاريخ 2025-09-25 وينتهي في 2026-09-24"
    m = re.search(r"مدة هذا العقد\s+(\d+)\s*(سنة|سنوات|شهر|أشهر)", text)
    if m:
        out["مدة العقد"] = m.group(1)

    m = re.search(r"يبدأ\s+من\s+تاريخ\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2}).*?وينتهي\s+في\s*[,،]?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})", text)
    if m:
        out["بدء العقد"] = format_date_any(m.group(1))
        out["انتهاء العقد"] = format_date_any(m.group(2))

    m = re.search(r"تاريخ\s+مباشرة.*?([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})", text)
    if m:
        out["تاريخ المباشرة الفعلية"] = format_date_any(m.group(1))

    m = re.search(r"فترة\s+تجربة.*?مدتها\s*(\d+)\s*يوم", text)
    if m:
        out["فترة التجربة"] = m.group(1)

    m = re.search(r"تحدد\s+أيام\s+العمل.*?ب\s*(\d+)\s*أيام", text)
    if m:
        out["أيام العمل الأسبوعية"] = m.group(1)

    m = re.search(r"تحدد\s+ساعات\s+العمل.*?ب\s*(\d+)\s*يومي", text)
    if m:
        out["ساعات العمل اليومية"] = m.group(1)

    # overtime %50
    m = re.search(r"٪\s*(\d+)\s*من\s*أجره\s*األساسي", text)
    if m:
        out["أجر الساعة الإضافية"] = m.group(1)

    # ---- Money ----
    m = re.search(r"أجرًا\s*أساسي\s*قدره\s*([0-9][0-9,]*(?:\.[0-9]+)?)", text)
    if m:
        out["الراتب الأساسي"] = clean_amount_to_int(m.group(1))

    m = re.search(r"أجر\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*ريال\s*سعودي\s*[,،]?\s*بدل\s*سكن", text)
    if m:
        out["بدل السكن"] = clean_amount_to_int(m.group(1))

    m = re.search(r"إجازة\s*سنوية\s*مدتها\s*(\d+)\s*يوم", text)
    if m:
        out["الإجازة السنوية"] = m.group(1)

    m = re.search(r"تعويضًا.*?قدره\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*ريال\s*سعودي", text)
    if m:
        out["التعويض عند الفسخ بدون سبب"] = clean_amount_to_int(m.group(1))

    # Final cleanup (None -> "")
    for k, v in list(out.items()):
        out[k] = "" if v is None else str(v).strip()

    return out
