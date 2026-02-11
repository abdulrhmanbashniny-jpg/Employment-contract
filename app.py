# -*- coding: utf-8 -*-
import io
import time
import shutil
import tempfile
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

from pdf_contracts import HEADERS, parse_contract, extract_text_from_pdf_bytes

st.set_page_config(page_title="PDF → Excel (عقود الموظفين)", page_icon="📄", layout="wide")

st.title("📄 تحويل عقود PDF إلى Excel")
st.write("ارفع ملفات PDF (أي عدد) وسيتم استخراج البيانات وتنظيمها في ملف Excel واحد، ثم حذف الملفات المؤقتة تلقائيًا.")

uploaded = st.file_uploader(
    "ارفع ملفات PDF هنا",
    type=["pdf"],
    accept_multiple_files=True
)

def build_excel_bytes(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "الموظفين"

    # Header row
    ws.append(HEADERS)
    for c in range(1, len(HEADERS) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Data rows
    for row in rows:
        ws.append([row.get(h, "") for h in HEADERS])

    ws.freeze_panes = "A2"

    # Column widths
    for col_idx, h in enumerate(HEADERS, start=1):
        max_len = len(h)
        for r in range(2, ws.max_row + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(10, max_len + 2), 45)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=len(HEADERS)):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()

if st.button("⚙️ تحويل إلى Excel", disabled=not uploaded):
    # مجلد مؤقت (يُحذف في النهاية)
    temp_dir = tempfile.mkdtemp(prefix="pdf_to_excel_")

    try:
        rows = []
        progress = st.progress(0)
        status = st.empty()

        total = len(uploaded)
        for i, f in enumerate(uploaded, start=1):
            status.write(f"جارٍ معالجة الملف {i}/{total}: **{f.name}**")
            pdf_bytes = f.read()

            text = extract_text_from_pdf_bytes(pdf_bytes)
            rows.append(parse_contract(text))

            progress.progress(int(i / total * 100))

        excel_bytes = build_excel_bytes(rows)

        st.success("✅ تم التحويل بنجاح!")
        st.download_button(
            label="⬇️ تنزيل Employees_Data.xlsx",
            data=excel_bytes,
            file_name="Employees_Data.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info("🧹 سيتم حذف الملفات المؤقتة تلقائيًا بعد ثوانٍ قليلة.")
        time.sleep(2)

    finally:
        # حذف كل شيء مؤقت (PDF/Excel) — لا يتم الاحتفاظ بأي ملفات
        shutil.rmtree(temp_dir, ignore_errors=True)
