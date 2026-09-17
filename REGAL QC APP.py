from datetime import datetime
import io
import json
import os
import sqlite3

from fpdf import FPDF
import openpyxl
from openpyxl.drawing.image import Image as OpenPyxlImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from PIL import Image
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# Optional Barcode Reader library
try:
    from pyzbar.pyzbar import decode as decode_barcode

    HAS_PYZBAR = True
except ImportError:
    HAS_PYZBAR = False

# --- LOCAL STORAGE FOLDER SETUP ---
STORAGE_DIR = "qc_local_storage"
PDF_DIR = os.path.join(STORAGE_DIR, "pdf_certificates")
EXCEL_DIR = os.path.join(STORAGE_DIR, "excel_reports")
IMAGES_DIR = os.path.join(STORAGE_DIR, "uploaded_images")

for folder in [PDF_DIR, EXCEL_DIR, IMAGES_DIR]:
    os.makedirs(folder, exist_ok=True)

# --- PAGE CONFIGURATION ---
st.set_page_config(
    page_title="Regal QC Portal",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# --- ADVANCED CUSTOM CSS & UI STYLING ---
st.markdown(
    """
    <style>
    .main .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 1200px; }
    .hero-header { background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #1F497D 100%); padding: 1.8rem 2rem; border-radius: 16px; color: white; margin-bottom: 2rem; box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25); display: flex; justify-content: space-between; align-items: center; }
    .hero-title { font-size: 2.1rem; font-weight: 800; margin: 0; letter-spacing: -0.025em; color: #FFFFFF; }
    .hero-subtitle { font-size: 0.95rem; color: #94A3B8; margin-top: 0.3rem; font-weight: 400; }
    .hero-badge { background: rgba(255, 255, 255, 0.1); border: 1px solid rgba(255, 255, 255, 0.2); padding: 0.5rem 1rem; border-radius: 30px; font-size: 0.85rem; color: #E2E8F0; backdrop-filter: blur(4px); }
    .kpi-card { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 1.25rem; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05); }
    .kpi-title { font-size: 0.8rem; font-weight: 600; color: #64748B; text-transform: uppercase; letter-spacing: 0.05em; }
    .kpi-value { font-size: 2.2rem; font-weight: 800; margin-top: 0.2rem; color: #0F172A; }
    .kpi-sub { font-size: 0.8rem; margin-top: 0.2rem; font-weight: 500; }
    .status-badge-resolved { background-color: #DCFCE7; color: #15803D; padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 700; display: inline-block; }
    .status-badge-open { background-color: #FEE2E2; color: #B91C1C; padding: 4px 12px; border-radius: 20px; font-size: 0.8rem; font-weight: 700; display: inline-block; }
    .section-label { font-size: 1rem; font-weight: 700; color: #1E293B; margin-bottom: 0.75rem; display: flex; align-items: center; gap: 0.5rem; }
    </style>
""",
    unsafe_allow_html=True,
)


# --- DATABASE INITIALIZATION ---
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
            photos_json TEXT,
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
            resolved, date_completed, product_desc, issue_desc, corrective_action, photo_data, photos_json
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
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
            entry["photos_json"],
        ),
    )
    conn.commit()
    conn.close()


def get_all_entries():
    conn = sqlite3.connect("qc_logs.db")
    cursor = conn.cursor()
    cursor.execute(
        "SELECT id, date_opened, product_code, quantity, photo_data, product_desc, issue_desc, corrective_action, resolved, date_completed, time_taken, supplier, photos_json FROM qc_entries ORDER BY id DESC"
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


# --- LOCAL IMAGE SAVER HELPER ---
def save_image_locally(image_bytes, prefix="img"):
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    filename = f"{prefix}_{timestamp}.png"
    filepath = os.path.join(IMAGES_DIR, filename)
    with open(filepath, "wb") as f:
        f.write(image_bytes)
    return filepath


# --- PDF CLEARANCE CERTIFICATE ENGINE ---
def generate_pdf_certificate(rec, save_local=True):
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
        p_json,
    ) = rec

    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 10, "REGAL DISTRIBUTORS SA", ln=True, align="C")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(
        0,
        6,
        "QUALITY CONTROL INSPECTION & CLEARANCE CERTIFICATE",
        ln=True,
        align="C",
    )
    pdf.line(10, 28, 200, 28)
    pdf.ln(10)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(45, 8, "Inspection Ref:", border=0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"#{rec_id}", ln=True)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(45, 8, "Product Code:", border=0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"{p_code}", ln=True)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(45, 8, "Supplier / Location:", border=0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"{supp if supp else 'N/A'}", ln=True)

    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(45, 8, "Approval Status:", border=0)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(
        0,
        8,
        f"{'APPROVED FOR SALE / PASSED' if res == 'Y' else 'REJECTED / ACTION REQUIRED'}",
        ln=True,
    )

    pdf.ln(6)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Product Description:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 6, p_desc if p_desc else "N/A")

    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Defect / Issue Observed:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 6, issue if issue else "N/A")

    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Corrective Action / Resolution:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 6, action if action else "N/A")

    pdf.ln(12)
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(
        0,
        6,
        f"Generated automatically via Regal QC System on {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        ln=True,
        align="R",
    )

    pdf_bytes = bytes(pdf.output())

    # Save to local disk folder
    if save_local:
        pdf_filename = f"QC_Certificate_Case_{rec_id}.pdf"
        local_pdf_path = os.path.join(PDF_DIR, pdf_filename)
        with open(local_pdf_path, "wb") as f:
            f.write(pdf_bytes)

    return pdf_bytes


# --- HEADER BANNER ---
st.markdown(
    """
    <div class="hero-header">
        <div>
            <div class="hero-title">🛡️ Regal QC Portal</div>
            <div class="hero-subtitle">Quality Assurance, Defect Tracking & Analytics Control Center</div>
        </div>
        <div class="hero-badge">⚡ System Active</div>
    </div>
""",
    unsafe_allow_html=True,
)

tab_dash, tab_new, tab_logs, tab_export = st.tabs(
    ["📊 Visual Dashboard", "➕ New Inspection", "📜 Inspection Logs", "📄 Export Reports"]
)

# --- TAB 1: VISUAL DASHBOARD ---
with tab_dash:
    records = get_all_entries()
    total_count = len(records)
    resolved_count = sum(1 for r in records if r[8] == "Y")
    open_count = total_count - resolved_count
    pass_rate = round((resolved_count / total_count) * 100, 1) if total_count > 0 else 0

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    with kpi1:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Total Inspected</div><div class="kpi-value">{total_count}</div></div>', unsafe_allow_html=True)
    with kpi2:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Passed / Resolved</div><div class="kpi-value" style="color: #16A34A;">{resolved_count}</div></div>', unsafe_allow_html=True)
    with kpi3:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Open / Blocked</div><div class="kpi-value" style="color: #DC2626;">{open_count}</div></div>', unsafe_allow_html=True)
    with kpi4:
        st.markdown(f'<div class="kpi-card"><div class="kpi-title">Pass Rate</div><div class="kpi-value" style="color: #2563EB;">{pass_rate}%</div></div>', unsafe_allow_html=True)

    if not records:
        st.info("No inspection records logged yet.")
    else:
        c1, c2 = st.columns([3, 2])
        with c1:
            suppliers = [r[11] if r[11] else "Unassigned" for r in records]
            supp_counts = {s: suppliers.count(s) for s in set(suppliers)}
            fig_supp = px.bar(x=list(supp_counts.keys()), y=list(supp_counts.values()), labels={"x": "Supplier", "y": "Issues"})
            st.plotly_chart(fig_supp, use_container_width=True)
        with c2:
            fig_pie = px.pie(names=["Passed", "Open Issues"], values=[resolved_count, open_count], color_discrete_sequence=["#16A34A", "#DC2626"])
            st.plotly_chart(fig_pie, use_container_width=True)

# --- TAB 2: NEW INSPECTION ENTRY ---
with tab_new:
    st.markdown("### 📝 Record New Quality Control Entry")

    with st.expander("📷 Open Barcode / QR Code Scanner"):
        cam_code = st.camera_input("Position product barcode in front of camera", key="cam_scanner")
        if cam_code and HAS_PYZBAR:
            img = Image.open(cam_code)
            decoded = decode_barcode(img)
            if decoded:
                scanned_code = decoded[0].data.decode("utf-8")
                st.session_state["scanned_code"] = scanned_code
                st.success(f"Barcode Scanned Successfully: **{scanned_code}**")

    scanned_val = st.session_state.get("scanned_code", "")
    st.markdown("<div class='section-label'>📷 Product / Defect Photo</div>", unsafe_allow_html=True)
    live_photo = st.camera_input("Take Live Photo", key="live_qc_photo")

    with st.form("qc_entry_form", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            date_opened = st.date_input("Inspection Date", value=datetime.today())
            product_code = st.text_input("Product Code", value=scanned_val, placeholder="e.g. IC118-4-1")
            quantity = st.text_input("Quantity Checked", value="1")
            time_taken = st.text_input("Inspection Time Spent", value="1 Hour")
        with col2:
            supplier = st.text_input("Supplier / Location", placeholder="e.g. Accentronix CC")
            resolved = st.selectbox("Inspection Status", ["Y", "N"], format_func=lambda x: "Passed / Resolved (Y)" if x == "Y" else "Failed / Open Issue (N)")
            date_completed = st.date_input("Completion Date", value=datetime.today())

        st.divider()
        product_desc = st.text_area("Product Description")
        issue_desc = st.text_area("Issue / Defect Description")
        corrective_action = st.text_area("Corrective Action Taken")
        st.divider()
        uploaded_photos = st.file_uploader("Attach Extra Supporting Images", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

        submitted = st.form_submit_button("💾 Save Inspection Record", use_container_width=True, type="primary")

        if submitted:
            if not product_code:
                st.error("Product Code is required.")
            else:
                encoded_photos = []
                first_photo = None

                if live_photo:
                    first_photo = live_photo.getvalue()
                    save_image_locally(first_photo, prefix=f"{product_code}_live")
                    encoded_photos.append(first_photo.hex())

                if uploaded_photos:
                    for p in uploaded_photos:
                        val_bytes = p.getvalue()
                        if not first_photo:
                            first_photo = val_bytes
                        save_image_locally(val_bytes, prefix=f"{product_code}_file")
                        encoded_photos.append(val_bytes.hex())

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
                    "photo_data": first_photo,
                    "photos_json": json.dumps(encoded_photos),
                }
                add_qc_entry(entry)
                st.toast("Inspection logged & images stored locally!", icon="✅")

# --- TAB 3: LOGS & HISTORY ---
with tab_logs:
    st.markdown("### 📜 Inspection Case Records")
    records = get_all_entries()

    if not records:
        st.info("No records present in database.")
    else:
        for rec in records:
            (rec_id, d_open, p_code, qty, photo, p_desc, issue, action, res, d_comp, t_taken, supp, p_json) = rec
            badge_html = '<span class="status-badge-resolved">PASSED / RESOLVED</span>' if res == "Y" else '<span class="status-badge-open">FAILED / OPEN ISSUE</span>'

            with st.expander(f"Case #{rec_id} — Item: {p_code} ({d_open})"):
                st.markdown(f"**Supplier:** {supp if supp else 'N/A'} | Status: {badge_html}", unsafe_allow_html=True)
                col_b1, col_b2 = st.columns(2)
                with col_b1:
                    pdf_bytes = generate_pdf_certificate(rec, save_local=True)
                    st.download_button(
                        label="📄 Download PDF Clearance Certificate",
                        data=pdf_bytes,
                        file_name=f"QC_Certificate_Case_{rec_id}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                with col_b2:
                    if st.button("🗑️ Delete Record", key=f"del_{rec_id}", use_container_width=True):
                        delete_entry(rec_id)
                        st.rerun()

# --- TAB 4: EXCEL EXPORT ENGINE ---
def build_excel_report(month_year_str, records, save_local=True):
    wb = openpyxl.Workbook()
    fill_header = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_title = Font(name="Calibri", size=16, bold=True)
    border_grid = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"), top=Side(style="thin", color="D9D9D9"), bottom=Side(style="thin", color="D9D9D9"))

    ws_summary = wb.active
    ws_summary.title = "QC Summary"
    ws_summary["A1"] = f"{month_year_str.upper()} QUALITY CONTROL SUMMARY"
    ws_summary["A1"].font = font_title

    ws_reg = wb.create_sheet(title="QC Register")
    ws_reg["A1"] = f"QUALITY CONTROL REGISTER — {month_year_str.upper()}"
    ws_reg["A1"].font = font_title

    headers = ["No.", "Date Opened", "Product Code", "Quantity", "Photo", "Description", "Issue Details", "Corrective Action", "Status (Y/N)", "Date Completed", "Time Spent", "Supplier"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws_reg.cell(row=3, column=col_idx, value=header)
        cell.fill = fill_header
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, rec in enumerate(records, 4):
        (rec_id, d_open, p_code, qty, photo, p_desc, issue, action, res, d_comp, t_taken, supp, p_json) = rec
        row_vals = [rec_id, d_open, p_code, qty, "", p_desc, issue, action, res, d_comp, t_taken, supp]
        for col_idx, val in enumerate(row_vals, 1):
            cell = ws_reg.cell(row=row_idx, column=col_idx, value=val)
            cell.border = border_grid

    output = io.BytesIO()
    wb.save(output)
    output_bytes = output.getvalue()

    # Save Excel report locally on disk
    if save_local:
        file_name = f"QC_Report_{month_year_str.replace(' ', '_')}.xlsx"
        local_excel_path = os.path.join(EXCEL_DIR, file_name)
        with open(local_excel_path, "wb") as f:
            f.write(output_bytes)

    output.seek(0)
    return output

with tab_export:
    st.markdown("### 📄 Export Monthly Excel Reports & Backups")
    e_col1, e_col2 = st.columns(2)
    with e_col1:
        month_select = st.text_input("Report Period", value="September 2026")
        if st.button("Generate Excel Report (.XLSX)", type="primary", use_container_width=True):
            all_logs = get_all_entries()
            if all_logs:
                excel_data = build_excel_report(month_select, all_logs, save_local=True)
                st.download_button(
                    label="⬇️ Download .XLSX File",
                    data=excel_data,
                    file_name=f"QC_Report_{month_select.replace(' ', '_')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
                st.success(f"Saved locally to `{EXCEL_DIR}`!")
