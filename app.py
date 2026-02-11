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

APP_TITLE = "📄 تحويل عقود PDF إلى Excel"
OUTPUT_FILE_NAME = "Employees_Data.xlsx"
SHEET_MAIN = "الموظفين"
SHEET_LOGS = "Logs"

st.set_page_config(page_title="PDF → Excel (عقود الموظفين)", page_icon="📄", layout="wide")

st.title(APP_TITLE)
st.write(
    "ارفع ملفات PDF (نصية) وسيتم استخراج البيانات وتنظيمها في ملف Excel واحد. "
    "أي بيانات غير موجودة في العقد سيتم تركها فارغة. "
    "وفي النهاية يتم حذف الملفات المؤقتة تلقائيًا."
)

with st.expander("⚙️ إعدادات (اختياري)", expanded=False):
    col1, col2, col3 = st.columns(3)
    with col1:
        add_logs_sheet = st.checkbox("إضافة ورقة Logs للأخطاء", value=True)
    with col2:
        show_preview = st.checkbox("عرض معاينة سريعة للنتائج", value=False)
    with col3:
        sleep_after_ready = st.slider("ثواني انتظار قبل تنظيف الملفات المؤقتة", 0, 5, 1)

uploaded = st.file_uploader(
    "ارفع ملفات PDF هنا",
    type=["pdf"],
    accept_multiple_files=True
)

def _auto_width(ws, max_width=45, min_width=10):
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

    # Rows: أي شيء غير موجود يطلع فاضي
    for row in rows:
        ws.append([row.get(h, "") if row.get(h, "") is not None else "" for h in HEADERS])

    ws.freeze_panes = "A2"

    # تنسيق بقية الخلايا
    for row_cells in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row_cells:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    _auto_width(ws)

    # Logs sheet
    if include_logs:
        ws2 = wb.create_sheet(SHEET_LOGS)
        ws2.append(["timestamp", "file_name", "status", "note"])
        for c in range(1, 5):
            cell = ws2.cell(row=1, column=c)
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        if logs:
            for item in logs:
                ws2.append([
                    item.get("timestamp", ""),
                    item.get("file_name", ""),
                    item.get("status", ""),
                    item.get("note", "")
                ])

        ws2.freeze_panes = "A2"
        for row_cells in ws2.iter_rows(min_row=2, max_row=ws2.max_row, max_col=4):
            for cell in row_cells:
                cell.alignment = Alignment(vertical="top", wrap_text=True)
        _auto_width(ws2, max_width=60)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

def process_files(files):
    rows = []
    logs = []

    total = len(files)
    progress = st.progress(0)
    status = st.empty()

    for i, f in enumerate(files, start=1):
        status.write(f"جارٍ معالجة الملف {i}/{total}: **{f.name}**")
        ts = datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC")

        try:
            pdf_bytes = f.read()
            if not pdf_bytes or len(pdf_bytes) < 50:
                # ملف فاضي تقريبًا
                rows.append({})
                logs.append({
                    "timestamp": ts,
                    "file_name": f.name,
                    "status": "SKIPPED",
                    "note": "File seems empty or too small"
                })
            else:
                text = extract_text_from_pdf_bytes(pdf_bytes)

                if not text.strip():
                    # PDF نصي لكن ما طلع نص (أو صفحة فاضية)
                    rows.append({})
                    logs.append({
                        "timestamp": ts,
                        "file_name": f.name,
                        "status": "OK_WITH_EMPTY_TEXT",
                        "note": "No extractable text found (empty result). Kept row blank."
                    })
                else:
                    data = parse_contract(text) or {}
                    # ضمان وجود كل الأعمدة (أي ناقص يبقى فاضي)
                    cleaned = {h: (data.get(h, "") if data.get(h, "") is not None else "") for h in HEADERS}
                    rows.append(cleaned)

                    logs.append({
                        "timestamp": ts,
                        "file_name": f.name,
                        "status": "OK",
                        "note": "Parsed successfully"
                    })

        except Exception as e:
            # لا نوقف — نكمل والباقي
            rows.append({})
            logs.append({
                "timestamp": ts,
                "file_name": f.name,
                "status": "ERROR",
                "note": f"{type(e).__name__}: {str(e)}"
            })

        progress.progress(int(i / total * 100))

    status.write("✅ انتهت المعالجة.")
    return rows, logs

if uploaded:
    st.info(f"عدد الملفات المرفوعة: **{len(uploaded)}**")

colA, colB = st.columns([1, 1])
with colA:
    run = st.button("⚙️ تحويل إلى Excel", disabled=not uploaded)
with colB:
    st.caption("ملاحظة: أي ملف فيه مشكلة لن يوقف العملية، وسيظهر في Logs.")

if run:
    # مجلد مؤقت (حتى لو ما خزّنا شيء، نخليه كحماية/تنظيف)
    temp_dir = tempfile.mkdtemp(prefix="pdf_to_excel_")

    try:
        rows, logs = process_files(uploaded)

        excel_bytes = build_excel_bytes(
            rows=rows,
            logs=logs,
            include_logs=add_logs_sheet
        )

        st.success("✅ تم تجهيز ملف Excel بنجاح!")

        st.download_button(
            label=f"⬇️ تنزيل {OUTPUT_FILE_NAME}",
            data=excel_bytes,
            file_name=OUTPUT_FILE_NAME,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # معاينة سريعة
        if show_preview:
            st.subheader("👀 معاينة سريعة (أول 5 صفوف)")
            # عرض بسيط بدون pandas
            preview_rows = rows[:5]
            for idx, r in enumerate(preview_rows, start=1):
                st.write(f"**Row {idx}**")
                st.json({k: r.get(k, "") for k in HEADERS[:12]})  # جزء من الحقول للعرض

        # عرض ملخص Logs
        if add_logs_sheet and logs:
            ok = sum(1 for x in logs if x["status"] == "OK")
            err = sum(1 for x in logs if x["status"] == "ERROR")
            other = len(logs) - ok - err
            st.write(f"📌 ملخص: OK={ok} | ERROR={err} | Other={other}")

        st.info("🧹 سيتم حذف الملفات المؤقتة تلقائيًا.")
        time.sleep(int(sleep_after_ready))

    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
