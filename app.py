# -*- coding: utf-8 -*-
import io
import time
import shutil
import tempfile
from datetime import datetime

import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from pdf_contracts import HEADERS, parse_contract, extract_text_from_pdf_bytes

APP_TITLE = "📄 تحويل عقود PDF إلى Excel (نسخة قوية + Debug)"
OUTPUT_FILE_NAME = "Employees_Data.xlsx"
SHEET_MAIN = "الموظفين"
SHEET_LOGS = "Logs"

# =========================
# UI
# =========================
st.set_page_config(page_title="PDF → Excel (عقود الموظفين)", page_icon="📄", layout="wide")
st.title(APP_TITLE)
st.write(
    "ارفع ملفات PDF (نصية). سيتم استخراج البيانات ووضع كل موظف في سطر واحد داخل Excel.\n"
    "أي حقل غير موجود في العقد سيبقى فارغ. وإذا ملف واحد فيه مشكلة، العملية تكمل للباقي."
)

with st.expander("⚙️ إعدادات متقدمة", expanded=False):
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        include_logs_sheet = st.checkbox("إضافة ورقة Logs", value=True)
    with col2:
        enable_debug = st.checkbox("Debug: عرض النص بعد التطبيع", value=True)
    with col3:
        show_preview_table = st.checkbox("عرض جدول جودة الملفات", value=True)
    with col4:
        cleanup_delay = st.slider("ثواني قبل تنظيف الملفات المؤقتة", 0, 8, 2)

uploaded = st.file_uploader("ارفع ملفات PDF هنا", type=["pdf"], accept_multiple_files=True)

# =========================
# Excel helpers
# =========================
def _auto_width(ws, max_width=55, min_width=10):
    for col_idx in range(1, ws.max_column + 1):
        header = ws.cell(row=1, column=col_idx).value or ""
        max_len = len(str(header))
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(min_width, max_len + 2), max_width)

def build_excel_bytes(rows, logs=None, include_logs=True):
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_MAIN

    # Header
    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data
    for row in rows:
        ws.append([row.get(h, "") if row.get(h, "") is not None else "" for h in HEADERS])

    ws.freeze_panes = "A2"
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    _auto_width(ws)

    # Logs
    if include_logs:
        ws2 = wb.create_sheet(SHEET_LOGS)
        ws2.append(["timestamp", "file_name", "status", "filled_fields", "total_fields", "quality_%", "note"])
        for c in range(1, 8):
            cell = ws2.cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        if logs:
            for item in logs:
                ws2.append([
                    item.get("timestamp", ""),
                    item.get("file_name", ""),
                    item.get("status", ""),
                    item.get("filled_fields", ""),
                    item.get("total_fields", ""),
                    item.get("quality_pct", ""),
                    item.get("note", ""),
                ])

        ws2.freeze_panes = "A2"
        for row_cells in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=7):
            for cell in row_cells:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        _auto_width(ws2, max_width=80)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

# =========================
# Quality + Debug helpers
# =========================
def quality_score(row: dict) -> tuple[int, int, float]:
    total = len(HEADERS)
    filled = sum(1 for h in HEADERS if str(row.get(h, "")).strip() != "")
    pct = round((filled / total) * 100, 1) if total else 0.0
    return filled, total, pct

def safe_snip(s: str, max_lines=120) -> str:
    if not s:
        return ""
    lines = s.splitlines()
    return "\n".join(lines[:max_lines])

# =========================
# Main processing
# =========================
def process_files(files):
    rows = []
    logs = []
    debug_items = []  # list of dicts for debug UI

    total_files = len(files)
    progress = st.progress(0)
    status = st.empty()

    for i, f in enumerate(files, start=1):
        status.write(f"جارٍ معالجة الملف {i}/{total_files}: **{f.name}**")
        ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

        try:
            pdf_bytes = f.read()

            if not pdf_bytes or len(pdf_bytes) < 50:
                row = {h: "" for h in HEADERS}
                rows.append(row)

                filled, total, pct = quality_score(row)
                logs.append({
                    "timestamp": ts,
                    "file_name": f.name,
                    "status": "SKIPPED",
                    "filled_fields": filled,
                    "total_fields": total,
                    "quality_pct": pct,
                    "note": "File seems empty/too small"
                })
                debug_items.append({
                    "file": f.name,
                    "raw_preview": "",
                    "normalized_preview": "",
                    "note": "Empty/too small"
                })

            else:
                # 1) normalized text (extract_text_from_pdf_bytes already normalizes)
                normalized_text = extract_text_from_pdf_bytes(pdf_bytes)

                # 2) parse
                data = parse_contract(normalized_text) or {}
                row = {h: (data.get(h, "") if data.get(h, "") is not None else "") for h in HEADERS}
                rows.append(row)

                filled, total, pct = quality_score(row)

                # status label
                st_label = "OK"
                note = "Parsed successfully"
                if pct < 20:
                    st_label = "LOW_QUALITY"
                    note = "Very low filled fields; likely RTL/extraction layout issues."

                logs.append({
                    "timestamp": ts,
                    "file_name": f.name,
                    "status": st_label,
                    "filled_fields": filled,
                    "total_fields": total,
                    "quality_pct": pct,
                    "note": note
                })

                # Debug previews
                if enable_debug:
                    # raw preview isn't available here unless we re-extract without normalization
                    # (but we can still show the normalized which matters most)
                    debug_items.append({
                        "file": f.name,
                        "raw_preview": "(raw not captured in this build)",
                        "normalized_preview": safe_snip(normalized_text, 120),
                        "note": f"Quality {pct}%"
                    })

        except Exception as e:
            row = {h: "" for h in HEADERS}
            rows.append(row)

            filled, total, pct = quality_score(row)
            logs.append({
                "timestamp": ts,
                "file_name": f.name,
                "status": "ERROR",
                "filled_fields": filled,
                "total_fields": total,
                "quality_pct": pct,
                "note": f"{type(e).__name__}: {str(e)}"
            })
            debug_items.append({
                "file": f.name,
                "raw_preview": "",
                "normalized_preview": "",
                "note": f"ERROR: {type(e).__name__}: {str(e)}"
            })

        progress.progress(int(i / total_files * 100))

    status.write("✅ انتهت المعالجة.")
    return rows, logs, debug_items

# =========================
# Run
# =========================
if uploaded:
    st.info(f"عدد الملفات المرفوعة: **{len(uploaded)}**")
    run = st.button("⚙️ تحويل إلى Excel", type="primary")
else:
    run = False

if run:
    temp_dir = tempfile.mkdtemp(prefix="pdf_to_excel_")

    try:
        rows, logs, debug_items = process_files(uploaded)

        # Show quality table
        if show_preview_table:
            st.subheader("📊 جودة الاستخراج لكل ملف")
            # Simple table without pandas
            for item in logs:
                st.write(
                    f"- **{item['file_name']}** | "
                    f"Status: `{item['status']}` | "
                    f"Filled: {item['filled_fields']}/{item['total_fields']} | "
                    f"Quality: **{item['quality_pct']}%** | "
                    f"{item['note']}"
                )

        excel_bytes = build_excel_bytes(rows, logs=logs, include_logs=include_logs_sheet)

        st.success("✅ تم تجهيز ملف Excel بنجاح!")
        st.download_button(
            label=f"⬇️ تنزيل {OUTPUT_FILE_NAME}",
            data=excel_bytes,
            file_name=OUTPUT_FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        # Debug: show normalized text per file
        if enable_debug:
            st.subheader("🧪 Debug (النص بعد التطبيع لكل ملف)")
            st.caption("هذا القسم مهم جدًا لمعرفة لماذا حقل معيّن ما تطلع قيمته. لو النص طبيعي هنا، الاستخراج بيكون دقيق.")
            for d in debug_items:
                with st.expander(f"📄 {d['file']} — {d.get('note','')}", expanded=False):
                    st.text_area("Normalized Text Preview (first 120 lines)", d.get("normalized_preview",""), height=260)

        st.info("🧹 سيتم حذف الملفات المؤقتة تلقائيًا.")
        time.sleep(int(cleanup_delay))

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
