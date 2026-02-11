# -*- coding: utf-8 -*-
import re
import io
import unicodedata
import pdfplumber

# ========= Columns =========
HEADERS = [
    "رقم العقد","تاريخ العقد","شركة/مؤسسة","الرقم الوطني الموحد","رقم المنشأة","السجل التجاري","عنوان الشركة","مكان العمل",
    "بريد الشركة","المسؤول الموقع","الصفة","اسم الموظف","رقم الهوية","نوع الهوية","تاريخ الميلاد","تاريخ انتهاء الهوية",
    "الجنسية","الجنس","الديانة","الحالة الاجتماعية","المؤهل العلمي","التخصص","المهنة","الرقم الوظيفي","رقم الآيبان",
    "اسم البنك","بريد الموظف","رقم الجوال","بدء العقد","انتهاء العقد","تاريخ المباشرة الفعلية","مدة العقد","فترة التجربة",
    "أيام العمل الأسبوعية","ساعات العمل اليومية","الراتب الأساسي","بدل السكن","الإجازة السنوية","أجر الساعة الإضافية",
    "التعويض عند الفسخ بدون سبب"
]

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

# ========= Helpers =========
def normalize_text(t: str) -> str:
    t = unicodedata.normalize("NFKC", t or "")
    t = t.replace("\u200f", "").replace("\u200e", "")
    return t

def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            parts.append(normalize_text(page.extract_text() or ""))
    return "\n".join(parts)

def digits_only(s: str) -> str:
    if not s:
        return ""
    return "".join(re.findall(r"\d+", str(s)))

def clean_amount_to_int(s: str) -> str:
    """
    يرجّع رقم صحيح كنص: 2,000.00 -> 2000
    """
    if not s:
        return ""
    m = re.search(r"(\d[\d,]*)(?:\.\d+)?", str(s))
    if not m:
        return ""
    return m.group(1).replace(",", "")

def format_date_any(s: str) -> str:
    """
    يدعم: YYYY-MM-DD / DD-MM-YYYY / YYYY/MM/DD / DD/MM/YYYY
    ويرجع: DD/MM/YYYY
    """
    if not s:
        return ""
    s = str(s).strip()

    m = re.search(r"(\d{1,4})[-/](\d{1,2})[-/](\d{1,4})", s)
    if not m:
        return ""

    a, b, c = m.group(1), m.group(2), m.group(3)

    # إذا بدأت بـ 4 أرقام => Year first
    if len(a) == 4:
        year = int(a); month = int(b); day = int(c)
    else:
        day = int(a); month = int(b); year = int(c)
        if year < 100:
            year += 2000

    return f"{day:02d}/{month:02d}/{year:04d}"

def get_value_after_label(text: str, label: str) -> str:
    # label: value
    m = re.search(rf"{re.escape(label)}\s*:\s*([^\n]+)", text)
    return (m.group(1).strip() if m else "")

def get_value_before_label(text: str, label: str) -> str:
    # value : label  (مثل: email :البريد الإلكتروني)
    m = re.search(rf"([^\n:]+)\s*:\s*{re.escape(label)}", text)
    return (m.group(1).strip() if m else "")

def get_bi(text: str, labels):
    """
    labels: قائمة احتمالات (بالعربي/الانجليزي) لنفس الحقل
    يحاول label:value ثم value:label
    """
    for lab in labels:
        v = get_value_after_label(text, lab)
        if v:
            return v
    for lab in labels:
        v = get_value_before_label(text, lab)
        if v:
            return v
    return ""

def slice_between(text: str, start_kw: str, end_kw: str) -> str:
    s = text.find(start_kw)
    if s == -1:
        return ""
    e = text.find(end_kw, s + len(start_kw))
    if e == -1:
        return text[s:]
    return text[s:e]

def normalize_mobile_from_line(line: str) -> str:
    """
    يدعم:
    - 966 0550xxxxxx
    - 96605xxxxxx => 9665xxxxxx
    - 0590xxxxxx 966 (زي ملفات قوى)
    الناتج: رقم متصل بدون مسافات
    """
    if not line:
        return ""
    groups = re.findall(r"\d+", line)
    if not groups:
        return ""

    # هل يوجد كود الدولة؟
    has_966 = any(g == "966" or g.startswith("966") for g in groups)

    # استخرج الرقم المحلي (05xxxxxxxx أو 5xxxxxxxx)
    local = ""
    for g in groups:
        if len(g) == 10 and g.startswith("05"):
            local = g
            break
        if len(g) == 9 and g.startswith("5"):
            local = "0" + g
            break

    if has_966 and local:
        return "966" + local[1:]  # احذف 0 بعد 966

    # أحيانًا يكون كله متصل مثل 9660xxxxxxxx
    joined = "".join(groups)
    if joined.startswith("9660"):
        joined = "966" + joined[4:]
    return joined

def find_line_containing(text: str, keyword: str) -> str:
    for ln in text.splitlines():
        if keyword in ln:
            return ln.strip()
    return ""

# ========= Strong Parser for Qiwa-style contracts =========
def parse_contract(text: str) -> dict:
    out = {h: "" for h in HEADERS}
    if not text:
        return out

    # ---- Sections (stronger email split) ----
    sec_first  = slice_between(text, "الطرف الأول", "الطرف الثاني")
    sec_second = slice_between(text, "الطرف الثاني", "اتفق الطرفان")

    # ---- Contract number/date ----
    m = re.search(r"رقم العقد\s*:\s*(\d+)", text)
    if m:
        out["رقم العقد"] = m.group(1)

    # يوم )2024-09-21( أو (2024-09-21)
    m = re.search(r"في يوم\s*\)?\(?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})\s*\)?", text)
    if m:
        out["تاريخ العقد"] = format_date_any(m.group(1))

    # ---- First party fields ----
    out["شركة/مؤسسة"] = get_bi(text, ["شركة/مؤسسة", "Corporation/Company"])
    out["الرقم الوطني الموحد"] = get_bi(text, ["الرقم الوطني الموحد", "National Unified Number"])
    out["رقم المنشأة"] = get_bi(text, ["رقم المنشأة", "Establishment Number"])
    out["السجل التجاري"] = get_bi(text, ["السجل التجاري", "Commercial Registration"])
    out["عنوان الشركة"] = get_bi(text, ["العنوان", "Address"])
    out["مكان العمل"] = get_bi(text, ["مكان العمل", "Work Location"])

    # بريد الشركة (من قسم الطرف الأول)
    if sec_first:
        ems = EMAIL_RE.findall(sec_first)
        if ems:
            out["بريد الشركة"] = ems[0]

    # المسؤول الموقع + الصفة
    m = re.search(r"ويمثلها بالتوقيع\s*:\s*([^\n]+)", text)
    if m:
        line = m.group(1).strip()
        # مثال: عبدالرحمن ... بصفته مدير الموارد البشرية
        if "بصفته" in line:
            left, right = line.split("بصفته", 1)
            out["المسؤول الموقع"] = left.strip()
            out["الصفة"] = right.strip()
        else:
            out["المسؤول الموقع"] = line

    # ---- Second party fields ----
    # الاسم في ملفات قوى يأتي: NAME :االسم
    name = get_value_before_label(text, "االسم") or get_value_after_label(text, "االسم") or get_value_after_label(text, "Name")
    out["اسم الموظف"] = name

    out["المهنة"] = get_bi(text, ["المهنة", "Profession"])
    out["الرقم الوظيفي"] = digits_only(get_bi(text, ["الرقم الوظيفي", "Employee Number"]))
    out["الجنسية"] = get_bi(text, ["الجنسية", "Nationality"])
    out["تاريخ الميلاد"] = format_date_any(get_bi(text, ["تاريخ الميالد", "Date of Birth"]))
    out["رقم الهوية"] = digits_only(get_bi(text, ["رقم الهوية", "Identity Number"]))
    out["نوع الهوية"] = get_bi(text, ["نوع الهوية", "ID Type"])
    out["تاريخ انتهاء الهوية"] = format_date_any(get_bi(text, ["تاريخ اإلنتهاء", "تاريخ الانتهاء", "ID Expiry Date"]))
    out["الجنس"] = get_bi(text, ["الجنس", "Gender"])
    out["الديانة"] = get_bi(text, ["الديانة", "Religion"])
    out["الحالة الاجتماعية"] = get_bi(text, ["الحالة االجتماعية", "الحالة الاجتماعية", "Marital Status"])
    out["المؤهل العلمي"] = get_bi(text, ["المؤهل العلمي", "Education"])
    out["التخصص"] = get_bi(text, ["التخصص", "Speciality"])

    # IBAN / Bank
    iban = get_bi(text, ["رقم اآليبان", "رقم الآيبان", "Iban"])
    out["رقم الآيبان"] = re.sub(r"\s+", "", iban).strip()
    out["اسم البنك"] = get_bi(text, ["اسم البنك", "Bank Name"])

    # بريد الموظف (من قسم الطرف الثاني)
    if sec_second:
        ems2 = EMAIL_RE.findall(sec_second)
        if ems2:
            out["بريد الموظف"] = ems2[0]

    # جوال (خط خاص عشان صيغة "0590... 966")
    mobile_line = find_line_containing(text, "رقم الجوال")
    if mobile_line:
        out["رقم الجوال"] = normalize_mobile_from_line(mobile_line)

    # ---- Contract terms (مدة/بدء/انتهاء/مباشرة/تجربة) ----
    # مدة هذا العقد 1 سنة / 6 شهر
    m = re.search(r"مدة هذا العقد\s+(\d+)\s*(سنة|سنوات|شهر|أشهر)", text)
    if m:
        out["مدة العقد"] = m.group(1)

    # يبدأ من تاريخ YYYY-MM-DD وينتهي في ,YYYY-MM-DD
    m = re.search(r"يبدأ من تاريخ\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2}).*?وينتهي في\s*[,،]?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})", text)
    if m:
        out["بدء العقد"] = format_date_any(m.group(1))
        out["انتهاء العقد"] = format_date_any(m.group(2))

    # تاريخ مباشرة ... هو .YYYY-MM-DD
    m = re.search(r"تاريخ مباشرة.*?هو\s*[\.،,]?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})", text)
    if m:
        out["تاريخ المباشرة الفعلية"] = format_date_any(m.group(1))

    # فترة تجربة مدتها 90 يومًا
    m = re.search(r"فترة تجربة مدتها\s*(\d+)\s*يوم", text)
    if m:
        out["فترة التجربة"] = m.group(1)

    # أيام/ساعات
    m = re.search(r"تحدد أيام العمل العادية بـ\s*(\d+)\s*أيام", text)
    if m:
        out["أيام العمل الأسبوعية"] = m.group(1)

    m = re.search(r"تحدد ساعات العمل بـ\s*(\d+)\s*يومي", text)
    if m:
        out["ساعات العمل اليومية"] = m.group(1)

    # أجر الساعة الإضافية %50
    m = re.search(r"مضافًا إليه\s*٪\s*(\d+)", text)
    if m:
        out["أجر الساعة الإضافية"] = m.group(1)

    # ---- Money fields ----
    # الراتب الأساسي: "أجرًا أساسي قدره 2,000.00 ريال سعودي"
    m = re.search(r"أجرًا\s*أساسي\s*قدره\s*([0-9][0-9,]*(?:\.[0-9]+)?)", text)
    if m:
        out["الراتب الأساسي"] = clean_amount_to_int(m.group(1))

    # بدل السكن: "أجر 500.00 ريال سعودي, بدل سكن"
    m = re.search(r"أجر\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*ريال\s*سعودي\s*[,،]?\s*بدل\s*سكن", text)
    if m:
        out["بدل السكن"] = clean_amount_to_int(m.group(1))

    # الإجازة السنوية: "إجازة سنوية مدتها 21 يومًا"
    m = re.search(r"إجازة\s*سنوية\s*مدتها\s*(\d+)\s*يوم", text)
    if m:
        out["الإجازة السنوية"] = m.group(1)

    # التعويض: "تعويضًا ... قدره 2,500.00 ريال سعودي"
    m = re.search(r"تعويضًا.*?قدره\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*ريال\s*سعودي", text)
    if m:
        out["التعويض عند الفسخ بدون سبب"] = clean_amount_to_int(m.group(1))

    # تنظيف نهائي (أي None -> "")
    for k, v in list(out.items()):
        out[k] = "" if v is None else str(v).strip()

    return out
