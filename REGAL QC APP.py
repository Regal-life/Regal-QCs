from datetime import datetime
import io
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import streamlit as st

st.set_page_config(
    page_title="Regal QC Generator", page_icon="📋", layout="centered"
)

st.title("📋 QC Report Generator")
st.subheader("Generate Daily & Monthly QC Excel Reports")

if "qc_entries" not in st.session_state:
    st.session_state.qc_entries = []

# --- SECTION 1: MOBILE ENTRY FORM ---
st.markdown("### ➕ Log New QC Incident")

with st.form("qc_entry_form", clear_on_submit=True):
    col1, col2 = st.columns(2)
    with col1:
        date_opened = st.date_input("Date Opened", value=datetime.today())
        product_code = st.text_input(
            "Product Code", placeholder="e.g. IC118-4-1"
        )
        quantity = st.text_input("Quantity", value="1")
        time_taken = st.text_input("Time Taken", value="1 Hour")

    with col2:
        supplier = st.text_input(
            "Supplier / Location", placeholder="e.g. Accentronix CC"
        )
        resolved = st.selectbox("Resolved?", ["Y", "N"])
        date_completed = st.date_input("Date Completed", value=datetime.today())

    product_desc = st.text_area("Product Description")
    issue_desc = st.text_area("Background / Issue Description")
    corrective_action = st.text_area("Corrective Action / Outcome")

    # Camera Input Component
    st.markdown("#### 📷 Product Photo")
    camera_photo = st.camera_input("Take a photo of the product/defect")

    submitted = st.form_submit_button("Add Entry to Batch")

    if submitted:
        photo_bytes = camera_photo.getvalue() if camera_photo else None

        entry = {
            "no": len(st.session_state.qc_entries) + 1,
            "date_opened": datetime.combine(date_opened, datetime.min.time()),
            "product_code": product_code,
            "quantity": quantity,
            "photo_data": photo_bytes,
            "product_desc": product_desc,
            "issue_desc": issue_desc,
            "corrective_action": corrective_action,
            "resolved": resolved,
            "date_completed": datetime.combine(
                date_completed, datetime.min.time()
            )
            if resolved == "Y"
            else None,
            "time_taken": time_taken,
            "supplier": supplier,
        }
        st.session_state.qc_entries.append(entry)
        st.success(f"Entry #{entry['no']} added successfully!")

# --- SECTION 2: BATCH PREVIEW ---
if st.session_state.qc_entries:
    st.markdown("### 📝 Current Batch Entries")
    for item in st.session_state.qc_entries:
        st.write(
            f"**#{item['no']} - {item['product_code']}** ({item['supplier']})"
        )
        if item["photo_data"]:
            st.image(item["photo_data"], width=150)

    if st.button("Clear All Entries"):
        st.session_state.qc_entries = []
        st.rerun()


# --- SECTION 3: EXCEL GENERATION ENGINE ---
def build_excel_report(month_year_str, entries):
    wb = openpyxl.Workbook()

    fill_header = PatternFill(
        start_color="1F497D", end_color="1F497D", fill_type="solid"
    )
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=16, bold=True)
    font_kpi_label = Font(name="Calibri", size=9, bold=True, color="595959")
    border_grid = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    # 1. Summary Sheet
    ws_summary = wb.active
    ws_summary.title = "QC Presentation"
    ws_summary.views.sheetView[0].showGridLines = True
    ws_summary["A1"] = f"{month_year_str.upper()} QUALITY CONTROL — SUMMARY"
    ws_summary["A1"].font = font_title

    # 2. Register Sheet
    ws_reg = wb.create_sheet(title=f"QC {month_year_str}")
    ws_reg.views.sheetView[0].showGridLines = True
    ws_reg["A1"] = f"QUALITY CONTROL REGISTER — {month_year_str.upper()}"
    ws_reg["A1"].font = font_title

    headers = [
        "No.",
        "Date Opened/Checked",
        "Product Code",
        "Quantity",
        "Picture of Product",
        "Product Description",
        "Background / Issue Description",
        "Corrective Action",
        "Resolved (Y/N)",
        "Date Completed",
        "Time Taken",
        "Supplier",
    ]

    for col_idx, header in enumerate(headers, 1):
        cell = ws_reg.cell(row=3, column=col_idx, value=header)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(
            horizontal="center", vertical="center", wrap_text=True
        )

    for row_idx, item in enumerate(entries, 4):
        # Row height for photo fitting
        ws_reg.row_dimensions[row_idx].height = 65 if item["photo_data"] else 20

        row_vals = [
            item["no"],
            item["date_opened"],
            item["product_code"],
            item["quantity"],
            "",  # Image embedded directly into cell
            item["product_desc"],
            item["issue_desc"],
            item["corrective_action"],
            item["resolved"],
            item["date_completed"],
            item["time_taken"],
            item["supplier"],
        ]

        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_reg.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border_grid
            if col_idx in [2, 10] and isinstance(val, datetime):
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center")
            elif col_idx in [1, 4, 9]:
                cell.alignment = Alignment(horizontal="center")

        # Embed Image in Column 5 ("E")
        if item["photo_data"]:
            img_bytes = io.BytesIO(item["photo_data"])
            img = Image.open(img_bytes)
            img.thumbnail((80, 80))

            img_out = io.BytesIO()
            img.save(img_out, format="PNG")
            img_out.seek(0)

            xl_img = OpenPyxlImage(img_out)
            ws_reg.add_image(xl_img, f"E{row_idx}")

    for col in ws_reg.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws_reg.column_dimensions[col_letter].width = min(
            max(max_len + 3, 12), 40
        )

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


# --- SECTION 4: EXPORT ---
st.markdown("---")
month_select = st.text_input("Report Month/Year", value="August 2026")

if st.button("Build Excel File"):
    if not st.session_state.qc_entries:
        st.warning("Please add at least one entry before exporting.")
    else:
        excel_data = build_excel_report(
            month_select, st.session_state.qc_entries
        )
        st.download_button(
            label="⬇️ Download Monthly QC Report (.xlsx)",
            data=excel_data,
            file_name=f"Monthly_QC_Report_{month_select.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
