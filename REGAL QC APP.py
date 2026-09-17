from datetime import datetime
import io
import sqlite3
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import streamlit as st

# Page Configuration
st.set_page_config(
    page_title="Regal QC Manager",
    page_icon="📋",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# --- CUSTOM CSS UI STYLING ---
st.markdown(
    """
    <style>
    /* Global Container Styling */
    .main .block-container {
        padding-top: 2rem;
        padding-bottom: 3rem;
        max-width: 800px;
    }
    
    /* Header Bar */
    .app-header {
        background: linear-gradient(135deg, #1F497D 0%, #112948 100%);
        padding: 1.5rem;
        border-radius: 12px;
        color: white;
        text-align: center;
        margin-bottom: 1.5rem;
        box-shadow: 0 4px 12px rgba(0, 0, 0, 0.1);
    }
    .app-header h1 {
        margin: 0;
        font-size: 1.8rem;
        font-weight: 700;
        color: #ffffff;
    }
    .app-header p {
        margin: 0.3rem 0 0 0;
        font-size: 0.95rem;
        opacity: 0.85;
    }

    /* Metric Cards */
    .metric-card {
        background-color: #ffffff;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .metric-value {
        font-size: 1.6rem;
        font-weight: 700;
        color: #1F497D;
    }
    .metric-label {
        font-size: 0.8rem;
        color: #64748B;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }

    /* Status Badges */
    .badge-resolved {
        background-color: #DEF7EC;
        color: #03543F;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 600;
        display: inline-block;
    }
    .badge-open {
        background-color: #FDE8E8;
        color: #9B1C1C;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
        font-weight: 600;
        display: inline-block;
    }

    /* Tabs Styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
    }
    .stTabs [data-baseweb="tab"] {
        height: 45px;
        border-radius: 8px;
        font-weight: 600;
    }
    </style>
""",
    unsafe_allow_html=True,
)


# --- DATABASE SETUP ---
def init_db():
    conn = sqlite3.connect("qc_logs.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS qc_entries (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            date_opened TEXT,
            product_code TEXT,
            quantity TEXT,
            time_taken TEXT,
            supplier TEXT,
            resolved TEXT,
            date_completed TEXT,
            product_desc TEXT,
            issue_desc TEXT,
            corrective_action TEXT,
            photo_data BLOB,
            timestamp DATETIME DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()


init_db()


def add_qc_entry(entry):
    conn = sqlite3.connect("qc_logs.db")
    cursor = conn.cursor()
    cursor.execute(
        """
        INSERT INTO qc_entries (
            date_opened, product_code, quantity, time_taken, supplier, 
            resolved, date_completed, product_desc, issue_desc, corrective_action, photo_data
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """,
        (
            entry["date_opened"],
            entry["product_code"],
            entry["quantity"],
            entry["time_taken"],
            entry["supplier"],
            entry["resolved"],
            entry["date_completed"],
            entry["product_desc"],
            entry["issue_desc"],
            entry["corrective_action"],
            entry["photo_data"],
        ),
    )
    conn.commit()
    conn.close()


def get_all_entries():
    conn = sqlite3.connect("qc_logs.db")
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, date_opened, product_code, quantity, photo_data, product_desc, issue_desc, corrective_action, resolved, date_completed, time_taken, supplier FROM qc_entries ORDER BY id DESC"
    )
    rows = cursor.fetchall()
    conn.close()
    return rows


def delete_entry(entry_id):
    conn = sqlite3.connect("qc_logs.db")
    cursor = conn.cursor()
    cursor.execute("DELETE FROM qc_entries WHERE id = ?", (entry_id,))
    conn.commit()
    conn.close()


# --- APP HEADER ---
st.markdown(
    """
    <div class="app-header">
        <h1>📋 Regal Quality Control</h1>
        <p>Mobile Inspection & Report Management</p>
    </div>
""",
    unsafe_allow_html=True,
)

# Navigation Tabs
tab1, tab2, tab3 = st.tabs(
    ["➕ New Case", "📜 History Logs", "📊 Generate Report"]
)

# --- TAB 1: NEW INCIDENT ---
with tab1:
    st.caption("Fill in the form below to record a new Quality Control case.")
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
            resolved = st.selectbox("Status", ["Y", "N"], format_func=lambda x: "Resolved (Y)" if x == "Y" else "Open (N)")
            date_completed = st.date_input(
                "Date Completed", value=datetime.today()
            )

        st.divider()
        product_desc = st.text_area("Product Description", placeholder="Enter item details...")
        issue_desc = st.text_area("Issue Description", placeholder="Describe the defect or query...")
        corrective_action = st.text_area("Corrective Action", placeholder="Action taken to resolve...")

        st.divider()
        st.markdown("📷 **Product Photo**")
        camera_photo = st.camera_input("Capture product image")

        submitted = st.form_submit_button("💾 Save Case Entry", use_container_width=True, type="primary")

        if submitted:
            if not product_code:
                st.error("Please enter a Product Code.")
            else:
                photo_bytes = camera_photo.getvalue() if camera_photo else None
                entry = {
                    "date_opened": date_opened.strftime("%Y-%m-%d"),
                    "product_code": product_code,
                    "quantity": quantity,
                    "time_taken": time_taken,
                    "supplier": supplier,
                    "resolved": resolved,
                    "date_completed": date_completed.strftime("%Y-%m-%d") if resolved == "Y" else "",
                    "product_desc": product_desc,
                    "issue_desc": issue_desc,
                    "corrective_action": corrective_action,
                    "photo_data": photo_bytes,
                }
                add_qc_entry(entry)
                st.toast("Case saved to database!", icon="✅")
                st.success(f"Case logged for product {product_code}.")


# --- TAB 2: HISTORY & LOGS ---
with tab2:
    records = get_all_entries()

    # Dashboard Metrics Banner
    total_cases = len(records)
    resolved_cases = sum(1 for r in records if r[8] == "Y")
    open_cases = total_cases - resolved_cases

    m1, m2, m3 = st.columns(3)
    with m1:
        st.markdown(f'<div class="metric-card"><div class="metric-value">{total_cases}</div><div class="metric-label">Total Cases</div></div>', unsafe_allow_html=True)
    with m2:
        st.markdown(f'<div class="metric-card"><div class="metric-value" style="color: #059669;">{resolved_cases}</div><div class="metric-label">Resolved</div></div>', unsafe_allow_html=True)
    with m3:
        st.markdown(f'<div class="metric-card"><div class="metric-value" style="color: #DC2626;">{open_cases}</div><div class="metric-label">Open</div></div>', unsafe_allow_html=True)

    st.write("")

    if not records:
        st.info("No recorded logs in the database.")
    else:
        for rec in records:
            (
                rec_id,
                d_open,
                p_code,
                qty,
                photo,
                p_desc,
                issue,
                action,
                res,
                d_comp,
                t_taken,
                supp,
            ) = rec

            badge_html = (
                '<span class="badge-resolved">Resolved</span>'
                if res == "Y"
                else '<span class="badge-open">Open</span>'
            )

            with st.expander(f"Case #{rec_id} — {p_code}"):
                st.markdown(f"**Supplier:** {supp} | **Status:** {badge_html}", unsafe_allow_html=True)
                st.write("")
                
                c1, c2 = st.columns([2, 1])
                with c1:
                    st.markdown(f"**Date Opened:** {d_open}")
                    st.markdown(f"**Quantity:** {qty} | **Time:** {t_taken}")
                    if p_desc:
                        st.markdown(f"**Description:** {p_desc}")
                    if issue:
                        st.markdown(f"**Issue:** {issue}")
                    if action:
                        st.markdown(f"**Action:** {action}")
                with c2:
                    if photo:
                        st.image(photo, use_container_width=True)

                st.divider()
                if st.button(f"🗑️ Delete Entry", key=f"del_{rec_id}", type="secondary"):
                    delete_entry(rec_id)
                    st.toast(f"Deleted Case #{rec_id}", icon="🗑️")
                    st.rerun()


# --- TAB 3: EXCEL GENERATOR ENGINE ---
def build_excel_report(month_year_str, records):
    wb = openpyxl.Workbook()

    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=16, bold=True)
    border_grid = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws_summary = wb.active
    ws_summary.title = "QC Presentation"
    ws_summary.views.sheetView[0].showGridLines = True
    ws_summary["A1"] = f"{month_year_str.upper()} QUALITY CONTROL — SUMMARY"
    ws_summary["A1"].font = font_title

    ws_reg = wb.create_sheet(title=f"QC {month_year_str}")
    ws_reg.views.sheetView[0].showGridLines = True
    ws_reg["A1"] = f"QUALITY CONTROL REGISTER — {month_year_str.upper()}"
    ws_reg["A1"].font = font_title

    headers = [
        "No.", "Date Opened/Checked", "Product Code", "Quantity", 
        "Picture of Product", "Product Description", "Background / Issue Description", 
        "Corrective Action", "Resolved (Y/N)", "Date Completed", "Time Taken", "Supplier"
    ]

    for col_idx, header in enumerate(headers, 1):
        cell = ws_reg.cell(row=3, column=col_idx, value=header)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_idx, rec in enumerate(records, 4):
        (rec_id, d_open, p_code, qty, photo, p_desc, issue, action, res, d_comp, t_taken, supp) = rec
        ws_reg.row_dimensions[row_idx].height = 65 if photo else 20

        row_vals = [
            rec_id, d_open, p_code, qty, "", p_desc, issue, action, res, d_comp, t_taken, supp
        ]

        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_reg.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border_grid
            if col_idx in [1, 4, 9]:
                cell.alignment = Alignment(horizontal="center")

        if photo:
            img_bytes = io.BytesIO(photo)
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
        ws_reg.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 40)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output


with tab3:
    st.subheader("Export Excel Workbook")
    st.caption("Compile stored entries into formatted monthly reports.")

    month_select = st.text_input("Report Month & Year", value="August 2026")

    if st.button("📊 Compile Report File", type="primary", use_container_width=True):
        all_logs = get_all_entries()
        if not all_logs:
            st.warning("No records stored in the database to export.")
        else:
            excel_data = build_excel_report(month_select, all_logs)
            st.download_button(
                label="⬇️ Download .XLSX Report File",
                data=excel_data,
                file_name=f"Monthly_QC_Report_{month_select.replace(' ', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
