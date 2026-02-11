# -*- coding: utf-8 -*-
import re
import unicodedata
import pdfplumber

ARABIC_RE = re.compile(r"[\u0600-\u06FF]")
LATIN_RE  = re.compile(r"[A-Za-z]")
EMAIL_RE  = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")

HEADERS = [
    "رقم العقد","تاريخ العقد","شركة/مؤسسة","الرقم الوطني الموحد","رقم المنشأة","السجل التجاري","عنوان الشركة","مكان العمل",
    "بريد الشركة","المسؤول الموقع","الصفة","اسم الموظف","رقم الهوية","نوع الهوية","تاريخ الميلاد","تاريخ انتهاء الهوية",
    "الجنسية","الجنس","الديانة","الحالة الاجتماعية","المؤهل العلمي","التخصص","المهنة","الرقم الوظيفي","رقم الآيبان",
    "اسم البنك","بريد الموظف","رقم الجوال","بدء العقد","انتهاء العقد","تاريخ المباشرة الفعلية","مدة العقد","فترة التجربة",
    "أيام العمل الأسبوعية","ساعات العمل اليومية","الراتب الأساسي","بدل السكن","الإجازة السنوية","أجر الساعة الإضافية",
    "التعويض عند الفسخ بدون سبب"
]

def normalize_text(t: str) -> str:
    t = unicodedata.normalize("NFKC", t or "")
    t = t.replace("\u200f", "").replace("\u200e", "")
    return t

def extract_text_from_pdf_path(pdf_path: str) -> str:
    parts = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            parts.append(normalize_text(page.extract_text() or ""))
    return "\n".join(parts)

def extract_text_from_pdf_bytes(pdf_bytes: bytes) -> str:
    parts = []
    with pdfplumber.open(pdf_bytes) as pdf:
        for page in pdf.pages:
            parts.append(normalize_text(page.extract_text() or ""))
    return "\n".join(parts)

def digits_only(s: str) -> str:
    if not s:
        return ""
    return "".join(re.findall(r"\d+", str(s)))

def clean_amount(s: str) -> str:
    if not s:
        return ""
    m = re.search(r"(\d[\d,]*)(?:\.\d+)?", s)
    if not m:
        return ""
    return m.group(1).replace(",", "")

def format_date(s: str) -> str:
    if not s:
        return ""
    s = str(s).strip()
    m = re.search(r"(\d{2,4})[-/](\d{1,2})[-/](\d{1,2})", s)
    if not m:
        return ""
    a, b, c = m.group(1), m.group(2), m.group(3)
    if len(a) == 4:
        year = int(a); month = int(b); day = int(c)
    else:
        day = int(a); month = int(b); year = int(c)
        if year < 100:
            year += 2000
    return f"{day:02d}/{month:02d}/{year:04d}"

def normalize_mobile(text: str) -> str:
    if not text:
        return ""
    raw = str(text)
    raw_no_spaces = re.sub(r"\s+", "", raw)

    groups = re.findall(r"\d+", raw)
    if not groups:
        return ""

    has_966 = any(g == "966" or g.startswith("966") for g in groups)

    local = None
    for g in groups:
        if len(g) == 10 and g.startswith("05"):
            local = g
            break
        if len(g) == 9 and g.startswith("5"):
            local = "0" + g
            break

    if has_966 and local:
        return "966" + local[1:]  # drop 0 after 966

    if raw_no_spaces.startswith("9660"):
        raw_no_spaces = "966" + raw_no_spaces[4:]

    return digits_only(raw_no_spaces)

def fix_rtl(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    if "@" in s or "SA" in s or re.fullmatch(r"[0-9\-/]+", s):
        return s
    ar = len(ARABIC_RE.findall(s))
    la = len(LATIN_RE.findall(s))
    if ar > la and ar >= 3:
        return s[::-1]
    return s

def find_field(patterns, text, flags=re.MULTILINE | re.IGNORECASE):
    for pat in patterns:
        m = re.search(pat, text, flags)
        if m:
            return m.group(1).strip()
    return ""

def find_before_label(label, text):
    m = re.search(rf"([^\n:]+)\s*:\s*{label}", text)
    return m.group(1).strip() if m else ""

def find_after_label(label, text):
    m = re.search(rf"{label}\s*:\s*([^\n]+)", text)
    return m.group(1).strip() if m else ""

def get_bi(label_std_patterns, label_rev_patterns, text):
    for lab in label_std_patterns:
        v = find_after_label(lab, text)
        if v:
            return v
    for lab in label_rev_patterns:
        v = find_before_label(lab, text)
        if v:
            return v
    return ""

def first_number(line: str) -> str:
    m = re.search(r"([0-9][0-9,\.]*)", line)
    return m.group(1) if m else ""

def parse_amount_by_keywords(text: str, keywords):
    for ln in text.splitlines():
        if all(k in ln for k in keywords):
            num = first_number(ln)
            if num:
                return clean_amount(num)
    return ""

def parse_contract(text: str) -> dict:
    # Party 1
    contract_no = find_field([
        r"رقم العقد:\s*([0-9]+)",
        r"([0-9]+)\s*:\s*دقعلا\s*مقر",
        r"ID\s*Contract\s*[:\-]?\s*([0-9]+)",
    ], text)

    contract_date = format_date(find_field([
        r"في يوم.*?\(?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})\s*\)?",
        r"\(\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})\s*\)",
        r"on\s*\(?([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})\)?",
    ], text))

    company = fix_rtl(get_bi(["شركة/مؤسسة","Corporation/Company"], ["ةسسؤم/ةكرش"], text).split("\n")[0].strip(" :"))
    nun     = get_bi(["الرقم الوطني الموحد","National Unified Number"], ["دحوملا ينطولا مقرلا"], text)
    estab   = get_bi(["رقم المنشأة","Establishment Number"], ["ةأشنملا مقر"], text)
    cr      = get_bi(["السجل التجاري","Commercial Registration"], ["يراجتلا لجسلا"], text)
    address = fix_rtl(get_bi(["العنوان","Address"], ["ناونعلا"], text).split("\n")[0].strip(" :"))
    workloc = fix_rtl(get_bi(["مكان العمل","Work Location"], ["لمعلا ناكم"], text).split("\n")[0].strip(" :"))

    emails = EMAIL_RE.findall(text)
    comp_email = emails[0] if emails else ""

    signatory = fix_rtl(find_field([
        r"ويمثلها بالتوقيع:\s*([^\n]+)",
        r"([^\n:]+)\s*:\s*عيقوتلاب اهلثميو",
    ], text))

    position  = fix_rtl(find_field([
        r"بصفته\s*([^\n]+)",
        r"([^\n]+)\s*هتفصب",
    ], text).split("\n")[0].strip(" :"))

    # Party 2
    m = re.search(r"Name:\s*([^\n]+)", text, re.IGNORECASE)
    emp_name = m.group(1).strip() if m else get_bi(["الاسم","االسم"], ["مسلاا"], text).split("\n")[0].strip(" :")
    emp_name = fix_rtl(emp_name.replace(":مسلاا","").strip())

    emp_id = digits_only(get_bi(["رقم الهوية","Identity Number"], ["ةيوهلا مقر"], text))
    id_type = fix_rtl(get_bi(["نوع الهوية","ID Type"], ["ةيوهلا عون"], text).split("\n")[0].strip(" :"))

    dob   = format_date(get_bi(["تاريخ الميالد","Date of Birth"], ["دلايملا خيرات"], text))
    idexp = format_date(get_bi(["تاريخ اإلنتهاء","تاريخ الانتهاء","ID Expiry Date"], ["ءاهتنلإا خيرات"], text))

    nationality = fix_rtl(get_bi(["الجنسية","Nationality"], ["ةيسنجلا"], text).split("\n")[0].strip(" :"))
    gender      = fix_rtl(get_bi(["الجنس","Gender"], ["سنجلا"], text).split("\n")[0].strip(" :"))
    religion    = fix_rtl(get_bi(["الديانة","Religion"], ["ةنايدلا"], text).split("\n")[0].strip(" :"))
    marital     = fix_rtl(get_bi(["الحالة االجتماعية","Marital Status"], ["ةيعامتجلاa ةلاحلا","ةيعامتجلاا ةلاحلا"], text).split("\n")[0].strip(" :"))
    education   = fix_rtl(get_bi(["المؤهل العلمي","Education"], ["يملاعلا لهؤملا"], text).split("\n")[0].strip(" :"))
    speciality  = fix_rtl(get_bi(["التخصص","Speciality"], ["صصختلا"], text).split("\n")[0].strip(" :"))

    profession  = fix_rtl(get_bi(["المهنة","Profession"], ["ةنهملا"], text).split("\n")[0].strip(" :"))
    emp_no      = digits_only(get_bi(["الرقم الوظيفي","Employee Number"], ["يفيظولا مقرلا"], text))

    iban = get_bi(["رقم اآليبان","Iban"], ["نابيلآا مقر"], text)
    iban = re.sub(r"\s+", "", iban).replace("null","").strip()

    bank = fix_rtl(get_bi(["اسم البنك","Bank Name"], ["كنبلا مسا"], text).split("\n")[0].strip(" :"))

    emp_email = ""
    m = re.search(r"(?:الطرف الثاني|SECOND PARTY|يناثلا فرطلا).*?(" + EMAIL_RE.pattern + r")", text, re.IGNORECASE|re.DOTALL)
    if m:
        emp_email = m.group(1)
    elif len(emails) >= 2:
        emp_email = emails[1]

    mobile = normalize_mobile(get_bi(["رقم الجوال","Mobile Number"], ["لاوجلا مقر"], text))

    # Terms
    duration = find_field([
        r"مدة هذا العقد\s*([0-9]+)\s*سنة",
        r"دقعلا اذه ةدم\s*([0-9]+)\s*ةنس",
        r"The contract[’']?s duration is\s*([0-9]+)\s*year",
    ], text)
    if not duration:
        duration = find_field([
            r"مدة هذا العقد\s*([0-9]+)\s*شهر",
            r"دقعلا اذه ةدم\s*([0-9]+)\s*رهش",
            r"duration is\s*([0-9]+)\s*month",
        ], text)

    start = format_date(find_field([
        r"يبدأ من تاريخ\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})",
        r"starting from\s*([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})",
    ], text))

    end = format_date(find_field([
        r"وينتهي في\s*,?\s*([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})",
        r"ends in\s*([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})",
    ], text))

    join = format_date(find_field([
        r"تاريخ مباشرة.*?هو\s*\.?([0-9]{4}[-/][0-9]{2}[-/][0-9]{2})",
        r"joining date.*?is\s*([0-9]{2}[-/][0-9]{2}[-/][0-9]{4})",
    ], text, flags=re.IGNORECASE | re.DOTALL))

    trial = find_field([
        r"فترة تجربة مدتها\s*([0-9]+)\s*يوم",
        r"ةبرجت\s*ةرتف\s*اهتدم\s*([0-9]+)\s*موي",
        r"trial period of\s*([0-9]+)\s*days",
    ], text)

    work_days  = find_field([r"أيام العمل العادية بـ\s*([0-9]+)", r"days per week.*?([0-9]+)"], text)
    work_hours = find_field([r"ساعات العمل بـ\s*([0-9]+)\s*يوميًا", r"daily hours.*?([0-9]+)"], text)

    overtime = find_field([
        r"٪\s*([0-9]+)\s*من أجره األساسي",
        r"plus\s*([0-9]+)\%",
    ], text)

    # amounts by lines
    basic_salary = parse_amount_by_keywords(text, ["يساسأ", "لاير"]) or parse_amount_by_keywords(text, ["أجر", "أساسي"])
    housing      = parse_amount_by_keywords(text, ["نكس", "لاير"]) or parse_amount_by_keywords(text, ["بدل", "سكن"])

    annual_leave = find_field([
        r"إجازة سنوية مدتها\s*([0-9]+)\s*يوم",
        r"annual leave of\s*([0-9]+)\s*days",
        r"ةيونس\s*ةزاجإ\s*ةدم\s*([0-9]+)\s*موي",
    ], text)

    termination = parse_amount_by_keywords(text, ["تعويض", "لاير"]) or parse_amount_by_keywords(text, ["ضيوعت", "لاير"])
    if not termination:
        m = re.search(r"(?:تعويض|ضيوعت)[^0-9]{0,80}([0-9][0-9\.,]+)", text)
        if m:
            termination = clean_amount(m.group(1))

    return {
        "رقم العقد": contract_no,
        "تاريخ العقد": contract_date,
        "شركة/مؤسسة": company,
        "الرقم الوطني الموحد": nun,
        "رقم المنشأة": estab,
        "السجل التجاري": cr,
        "عنوان الشركة": address,
        "مكان العمل": workloc,
        "بريد الشركة": comp_email,
        "المسؤول الموقع": signatory,
        "الصفة": position,

        "اسم الموظف": emp_name,
        "رقم الهوية": emp_id,
        "نوع الهوية": id_type,
        "تاريخ الميلاد": dob,
        "تاريخ انتهاء الهوية": idexp,
        "الجنسية": nationality,
        "الجنس": gender,
        "الديانة": religion,
        "الحالة الاجتماعية": marital,
        "المؤهل العلمي": education,
        "التخصص": speciality,
        "المهنة": profession,
        "الرقم الوظيفي": emp_no,
        "رقم الآيبان": iban,
        "اسم البنك": bank,
        "بريد الموظف": emp_email,
        "رقم الجوال": mobile,

        "بدء العقد": start,
        "انتهاء العقد": end,
        "تاريخ المباشرة الفعلية": join,
        "مدة العقد": digits_only(duration),
        "فترة التجربة": digits_only(trial),
        "أيام العمل الأسبوعية": digits_only(work_days),
        "ساعات العمل اليومية": digits_only(work_hours),
        "الراتب الأساسي": digits_only(basic_salary),
        "بدل السكن": digits_only(housing),
        "الإجازة السنوية": digits_only(annual_leave),
        "أجر الساعة الإضافية": digits_only(overtime),
        "التعويض عند الفسخ بدون سبب": digits_only(termination),
    }
