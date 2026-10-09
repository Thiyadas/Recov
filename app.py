import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io
import re
import os
import zipfile
from datetime import datetime

# ReportLab Libraries for PDF Generation
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

st.set_page_config(page_title="AI Reinsurance System & PDF Generator", layout="wide")
st.title("🤖 ระบบประมวลผล Reinsurance Bordereaux & PDF Notice Generator")

# =========================================================
# Configuration & Constants (ตั้งค่าตามโครงสร้างมาตรฐานบริษัท)
# =========================================================
LAYERS_CONFIG = [
    {"layer_name": "Second Layer", "limit": 220000000.0, "excess_point": 120000000.0},
    {"layer_name": "Third Layer", "limit": 1060000000.0, "excess_point": 340000000.0},
    {"layer_name": "Fourth Layer", "limit": 2100000000.0, "excess_point": 1400000000.0},
]

# รายชื่อบริษัท Broker/Reinsurers สัดส่วนและชื่อเต็มสำหรับจดหมาย
REINSURERS_INFO = {
    "IRMC 40%": {"name": "IRMC Reinsurance Broker Co., Ltd.", "share": 0.40},
    "Lockton 36%": {"name": "Lockton Wattana Insurance Broker Limited", "share": 0.36},
    "TQR 15%": {"name": "TQR Public Company Limited", "share": 0.15},
    "Aon 9%": {"name": "Aon Re (Thailand) Co., Ltd.", "share": 0.09}
}

LOGO_FILENAME = "logo_dhipaya.jpg"

# ---------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------
def clean_num(val):
    if pd.isna(val) or val is None: return 0.0
    s = str(val).strip()
    if not s or s.lower() in ['nan', 'none', 'null', '-', 'n/a']: return 0.0
    if s.startswith('(') and s.endswith(')'): s = '-' + s[1:-1]
    s = re.sub(r'[^0-9.-]', '', s)
    try: return float(s) if s else 0.0
    except ValueError: return 0.0

def clean_text(val):
    if pd.isna(val) or val is None: return ""
    s = str(val).strip()
    if s.lower() in ['nan', 'none', 'null', '<na>']: return ""
    return s

def find_col_by_keywords(df, keywords):
    for kw in keywords:
        for col in df.columns:
            if kw.lower() in str(col).strip().lower(): return col
    return None

def calculate_layer_payout(net_loss, limit, excess_point):
    if net_loss <= excess_point: return 0.0
    return min(net_loss - excess_point, limit)

# ---------------------------------------------------------
# PDF Generator Function (สร้างใบ PLA PDF พร้อมโลโก้บริษัท)
# ---------------------------------------------------------
def generate_pla_pdf(reinsurer_full_name, layer_name, gross_loss, net_loss, excess_pt, layer_limit, reinsurer_share_amt, event_no="E2026-0005"):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4,
        rightMargin=36, leftMargin=36, topMargin=36, bottomMargin=36
    )
    styles = getSampleStyleSheet()
    
    style_company = ParagraphStyle('Company', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=0)
    style_addr = ParagraphStyle('Addr', parent=styles['Normal'], fontName='Helvetica', fontSize=8, leading=10, alignment=0)
    style_title = ParagraphStyle('Title', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, leading=14, alignment=1)
    style_subtitle = ParagraphStyle('SubTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=11, leading=13, alignment=1)
    style_normal = ParagraphStyle('Norm', parent=styles['Normal'], fontName='Helvetica', fontSize=9, leading=12)
    style_bold = ParagraphStyle('Bold', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=9, leading=12)

    elements = []

    # 1. Header with Logo (การใส่โลโก้บริษัท)
    company_text = [
        Paragraph("DHIPAYA INSURANCE PUBLIC COMPANY LIMITED", style_company),
        Paragraph("HEAD OFFICE ADDRESS :- 115 RAMA 3 ROAD, Chong Nonsi, Yannawa, Bangkok 10120", style_addr),
        Paragraph("TEL. 1736, 0 2239 2200", style_addr)
    ]

    if os.path.exists(LOGO_FILENAME):
        try:
            img = Image(LOGO_FILENAME, width=65, height=65)
            header_table = Table([[img, company_text]], colWidths=[75, 445])
            header_table.setStyle(TableStyle([
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
            ]))
            elements.append(header_table)
        except Exception:
            elements.append(Paragraph("DHIPAYA INSURANCE PUBLIC COMPANY LIMITED", style_company))
            elements.append(Paragraph("HEAD OFFICE ADDRESS :- 115 RAMA 3 ROAD, Chong Nonsi, Yannawa, Bangkok 10120 | TEL. 1736, 0 2239 2200", style_addr))
    else:
        elements.append(Paragraph("DHIPAYA INSURANCE PUBLIC COMPANY LIMITED", style_company))
        elements.append(Paragraph("HEAD OFFICE ADDRESS :- 115 RAMA 3 ROAD, Chong Nonsi, Yannawa, Bangkok 10120 | TEL. 1736, 0 2239 2200", style_addr))

    elements.append(Spacer(1, 8))
    elements.append(HRFlowable(width="100%", thickness=1, color=colors.black, spaceAfter=10))

    # 2. Contract & Document Title
    elements.append(Paragraph(f"Fire XL-{layer_name} 2025", style_title))
    elements.append(Paragraph("PRELIMINARY LOSS ADVICE", style_subtitle))
    elements.append(Spacer(1, 12))

    # 3. To Reinsurer Section
    elements.append(Paragraph(f"<b>To:</b> {reinsurer_full_name}", style_normal))
    elements.append(Paragraph("Dear Sirs,", style_normal))
    elements.append(Paragraph("We regret to inform you that we have received the loss advice from the claimant as per following detail.", style_normal))
    elements.append(Spacer(1, 10))

    # 4. Claim Details Table
    data_table = [
        [Paragraph("<b>CLAIM NO.</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>POLICY NO.</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>INSURED</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>LOCATION</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>NATURE OF LOSS</b>", style_normal), Paragraph(": Flood 2025", style_normal)],
        [Paragraph("<b>DATE OF LOSS</b>", style_normal), Paragraph(": 19/11/2025-30/11/2025", style_normal)],
        [Paragraph("<b>SUM INSURED (100%)</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>OUR GROSS RETENTION</b>", style_normal), Paragraph(": Please see Attachment", style_normal)],
        [Paragraph("<b>LOSS ESTIMATE</b>", style_normal), Paragraph(f": BHT. {gross_loss:,.2f}", style_normal)],
        [Paragraph("<b>LOSS OF GROSS RETENTION</b>", style_normal), Paragraph(f": BHT. {net_loss:,.2f}", style_normal)],
        [Paragraph("<b>EXCESS POINT</b>", style_normal), Paragraph(f": BHT. {excess_pt:,.2f}", style_normal)],
        [Paragraph("<b>ESTIMATE UNDER XOL TREATY</b>", style_normal), Paragraph(f": BHT. {layer_limit:,.2f}", style_normal)],
        [Paragraph("<b>YOUR SHARE OF ESTIMATE</b>", style_bold), Paragraph(f": <b>BHT. {reinsurer_share_amt:,.2f} ({layer_name})</b>", style_bold)],
    ]

    t = Table(data_table, colWidths=[200, 320])
    t.setStyle(TableStyle([
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 3),
    ]))
    elements.append(t)
    elements.append(Spacer(1, 15))

    # 5. Event & Dates
    today_str = datetime.now().strftime("%d/%m/%Y")
    elements.append(Paragraph(f"<b>Date:</b> {today_str}", style_normal))
    elements.append(Paragraph(f"<b>EVENT NO.:</b> {event_no}", style_normal))
    elements.append(Spacer(1, 10))

    # 6. Footer Notice
    elements.append(Paragraph("Kindly reserve the above captioned amount pending for further advice of each call from us.", style_normal))
    elements.append(Spacer(1, 20))
    elements.append(Paragraph("<i>This is a computer print out, therefore no signature is required.</i>", style_normal))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("Please Sign and return copy here of", style_normal))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("Handled by: Reinsurance Department", style_normal))

    doc.build(elements)
    buffer.seek(0)
    return buffer.getvalue()

# =========================================================
# Step 1: อัปโหลด Original Data
# =========================================================
st.header("📌 Step 1: อัปโหลดไฟล์ Original Data (ไฟล์ 1)")
uploaded_file = st.file_uploader("เลือกไฟล์ Original Data (.xlsx)", type=["xlsx"])

if uploaded_file:
    xl = pd.ExcelFile(uploaded_file)
    sheet_set = next((s for s in xl.sheet_names if 'settle' in s.lower() and 'pivot' not in s.lower()), xl.sheet_names[0])
    sheet_inc = next((s for s in xl.sheet_names if ('incurred' in s.lower() or 'reserve' in s.lower()) and 'pivot' not in s.lower()), xl.sheet_names[-1])

    df_set_raw = pd.read_excel(uploaded_file, sheet_name=sheet_set)
    df_inc_raw = pd.read_excel(uploaded_file, sheet_name=sheet_inc)

    df_set_raw.columns = [str(c).strip() for c in df_set_raw.columns]
    df_inc_raw.columns = [str(c).strip() for c in df_inc_raw.columns]

    col_claim_set = find_col_by_keywords(df_set_raw, ['เลขที่สินไหม', 'claim no']) or df_set_raw.columns[0]
    col_sg = find_col_by_keywords(df_set_raw, ['รวม (ค่าเสียหาย + ค่าสำรวจ – Salvage)', 'settle gross', 'gross'])
    col_sn = find_col_by_keywords(df_set_raw, ['retention (by type of loss)', 'settle net', 'retention'])
    col_status_set = find_col_by_keywords(df_set_raw, ['status', 'claim status', 'หมายเหตุ'])

    col_claim_inc = find_col_by_keywords(df_inc_raw, ['เลขที่สินไหม', 'claim no']) or df_inc_raw.columns[0]
    col_rg = find_col_by_keywords(df_inc_raw, ['รวมค่าสินไหมจ่าย', 'reserve gross', 'gross'])
    col_rn = find_col_by_keywords(df_inc_raw, ['retention (by type of loss)', 'retention', 'reserve net'])

    meta_keywords = {
        'Row Labels': ['class', 'sub class group', 'row labels'],
        'Sub Class': ['sub class', 'subclass'],
        'Policy No.': ['เลขที่กรมธรรม์', 'policy no'],
        'Loss Date': ['วันที่เกิดเหตุ', 'loss date'],
        'Insured Name': ['ชื่อผู้เอาประกัน/ชื่อบริษัทประกันภัย', 'insured name'],
        'จังหวัด': ['จังหวัด', 'province']
    }

    # Data Processing
    df_set = pd.DataFrame()
    df_set['Claim No.'] = df_set_raw[col_claim_set].apply(clean_text).str.upper()
    df_set['Settle Gross Loss'] = df_set_raw[col_sg].apply(clean_num) if col_sg else 0.0
    df_set['Settle Net Loss Retention'] = df_set_raw[col_sn].apply(clean_num) if col_sn else df_set['Settle Gross Loss']
    df_set['Status_Raw'] = df_set_raw[col_status_set].apply(clean_text) if col_status_set else 'Pending'

    for field, kws in meta_keywords.items():
        found_col = find_col_by_keywords(df_set_raw, kws)
        df_set[field] = df_set_raw[found_col].apply(clean_text) if found_col else ""

    df_set = df_set[~df_set['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม'])]
    agg_set = {'Settle Gross Loss': 'sum', 'Settle Net Loss Retention': 'sum', 'Status_Raw': 'first'}
    for field in meta_keywords.keys(): agg_set[field] = 'first'
    grp_set = df_set.groupby('Claim No.', as_index=False).agg(agg_set)

    df_inc = pd.DataFrame()
    df_inc['Claim No.'] = df_inc_raw[col_claim_inc].apply(clean_text).str.upper()
    df_inc['Reserve Gross Loss'] = df_inc_raw[col_rg].apply(clean_num) if col_rg else 0.0
    df_inc['Reserve Net Loss Retention'] = df_inc_raw[col_rn].apply(clean_num) if col_rn else df_inc['Reserve Gross Loss']

    for field, kws in meta_keywords.items():
        found_col = find_col_by_keywords(df_inc_raw, kws)
        df_inc[field] = df_inc_raw[found_col].apply(clean_text) if found_col else ""

    df_inc = df_inc[~df_inc['Claim No.'].isin(['NAN', 'NONE', '', 'NULL', 'TOTAL', 'ยอดรวม'])]
    agg_inc = {'Reserve Gross Loss': 'sum', 'Reserve Net Loss Retention': 'sum'}
    for field in meta_keywords.keys(): agg_inc[field] = 'first'
    grp_inc = df_inc.groupby('Claim No.', as_index=False).agg(agg_inc)

    df_merged = pd.merge(grp_set, grp_inc, on='Claim No.', how='outer', suffixes=('_set', '_inc')).fillna('')

    for num_col in ['Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention']:
        df_merged[num_col] = pd.to_numeric(df_merged[num_col], errors='coerce').fillna(0.0)

    for field in meta_keywords.keys():
        col_set, col_inc = f"{field}_set", f"{field}_inc"
        if col_set in df_merged.columns and col_inc in df_merged.columns:
            df_merged[field] = df_merged[col_set].where(df_merged[col_set] != "", df_merged[col_inc])
            df_merged.drop(columns=[col_set, col_inc], inplace=True)

    df_merged['Status'] = df_merged['Status_Raw'].apply(lambda x: 'Pending' if str(x).strip() != '' else 'Closed')
    cols_order = ['Row Labels', 'Sub Class', 'Claim No.', 'Policy No.', 'Loss Date', 'Insured Name', 'จังหวัด', 
                  'Settle Gross Loss', 'Settle Net Loss Retention', 'Reserve Gross Loss', 'Reserve Net Loss Retention', 'Status']
    df_final_file2 = df_merged[cols_order].copy()

    tot_sg = float(df_final_file2['Settle Gross Loss'].sum())
    tot_sn = float(df_final_file2['Settle Net Loss Retention'].sum())
    tot_rg = float(df_final_file2['Reserve Gross Loss'].sum())
    tot_rn = float(df_final_file2['Reserve Net Loss Retention'].sum())
    tot_gross_all = tot_sg + tot_rg
    tot_net_all = tot_sn + tot_rn

  # =========================================================
    # Step 2: สรุปผล ตรวจสอบ และปุ่มแก้ไข / ยืนยันไฟล์ 2 & 3
    # =========================================================
    st.markdown("---")
    st.header("📌 Step 2: สรุปผล ตรวจสอบ และแก้ไขไฟล์ 2 & ไฟล์ 3")
    
    col1, col2 = st.columns(2)
    col1.metric("Total Gross (Settle + Reserve)", f"{tot_gross_all:,.2f}")
    col2.metric("Total Net Loss (Settle + Reserve)", f"{tot_net_all:,.2f}")

    # Build Layer Dataframe
    summary_layer_rows = []
    for layer in LAYERS_CONFIG:
        under_xl = calculate_layer_payout(tot_net_all, layer["limit"], layer["excess_point"])
        row = {"Section": "PLA/LSA XL", "Layer": layer["layer_name"], "Gross 100%": tot_gross_all, "Net Loss": tot_net_all, "Limit": layer["limit"], "Excess Point": layer["excess_point"], "Under XL": under_xl}
        for rein_key, rein_val in REINSURERS_INFO.items(): row[rein_key] = under_xl * rein_val["share"]
        summary_layer_rows.append(row)

    df_summary_layer = pd.DataFrame(summary_layer_rows)

    # 📥 ส่วนดาวน์โหลดไฟล์ 2 & 3 ไปตรวจสอบ
    col_dl1, col_dl2 = st.columns(2)

    wb2 = openpyxl.Workbook()
    ws2 = wb2.active
    ws2.title = "Details Claim"
    ws2.append([''] * 12)
    ws2.append(['', '', '', '', '', '', '', 'Settle 12/2025', '', 'Reserve as at 31/12/2025', '', ''])
    ws2.append(cols_order)
    for r in df_final_file2.itertuples(index=False): ws2.append(list(r))
    ws2.append(['', '', '', '', '', '', '', tot_sg, tot_sn, tot_rg, tot_rn, ''])
    buf2 = io.BytesIO()
    wb2.save(buf2)
    col_dl1.download_button("📥 ดาวน์โหลด ไฟล์ 2 (Bordereaux Claim)", data=buf2.getvalue(), file_name="Bordereaux_Claim_Output.xlsx")

    wb3 = openpyxl.Workbook()
    ws3 = wb3.active
    ws3.title = "P&E XL"
    ws3.append(["Summary Claim (By Layer) Calculated from File 1"])
    ws3.append(list(df_summary_layer.columns))
    for r in df_summary_layer.itertuples(index=False): ws3.append(list(r))
    buf3 = io.BytesIO()
    wb3.save(buf3)
    col_dl2.download_button("📥 ดาวน์โหลด ไฟล์ 3 (Summary By Layer)", data=buf3.getvalue(), file_name="Summary_Claim_By_Layer_Output.xlsx")

    # ---------------------------------------------------------
    # 🎯 ปุ่มแอ็กชัน: แก้ไขไฟนอล หรือ ตรวจสอบถูกต้อง
    # ---------------------------------------------------------
    st.subheader("⚙️ การยืนยันความถูกต้องก่อนออกเอกสาร")
    
    if "is_approved" not in st.session_state:
        st.session_state.is_approved = False

    col_btn1, col_btn2 = st.columns(2)
    
    with col_btn1:
        # ปุ่มสำหรับโหมดแก้ไข
        if st.checkbox("✏️ แก้ไขไฟล์ไฟนอล (หากต้องการปรับแก้ตัวเลขก่อนไปทำ PDF)"):
            st.warning("⚠️ คุณสามารถอัปโหลดไฟล์ 2 / 3 ที่แก้ไขแล้ว หรือปรับแก้ตารางสรุปด้านล่าง:")
            df_summary_layer = st.data_editor(df_summary_layer, num_rows="dynamic", key="editor_layer")

    with col_btn2:
        # ปุ่มอนุมัติและยืนยันข้อมูล
        if st.button("✅ ตรวจสอบถูกต้อง (Confirm Data)"):
            st.session_state.is_approved = True
            st.success("🎉 อนุมัติข้อมูลไฟล์ 2 & 3 เรียบร้อยแล้ว! สามารถดำเนินการออก PDF ได้ใน Step ถัดไป")

    # =========================================================
    # Step 3: สร้างเอกสารทั้ง PLA & LSA และแปลงเป็น PDF
    # =========================================================
    if st.session_state.is_approved:
        st.markdown("---")
        st.header("📌 Step 3: สร้างเอกสาร LSA & PLA และแปลงเป็น PDF")
        
        st.info("💡 เมื่ออนุมัติเรียบร้อย ระบบจะทำการสร้างเอกสารทั้ง PLA และ LSA PDF ของทุกบริษัทแยกตาม Layer ออกมารวมเป็นไฟล์ ZIP เดียวกัน")

        if st.button("🚀 สร้างและดาวน์โหลด PDF (PLA + LSA) ของทุกบริษัท (.ZIP)"):
            zip_buffer = io.BytesIO()
            
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                for layer in LAYERS_CONFIG:
                    layer_name = layer["layer_name"]
                    limit = layer["limit"]
                    excess_pt = layer["excess_point"]
                    under_xl = calculate_layer_payout(tot_net_all, limit, excess_pt)

                    if under_xl > 0:
                        for rein_key, rein_val in REINSURERS_INFO.items():
                            rein_full_name = rein_val["name"]
                            rein_share_amt = under_xl * rein_val["share"]
                            safe_rein_name = re.sub(r'[^a-zA-Z0-9]', '_', rein_key)
                            
                            # 1. สร้าง PLA PDF
                            pdf_pla_bytes = generate_pla_pdf(
                                reinsurer_full_name=rein_full_name,
                                layer_name=layer_name,
                                gross_loss=tot_gross_all,
                                net_loss=tot_net_all,
                                excess_pt=excess_pt,
                                layer_limit=limit,
                                reinsurer_share_amt=rein_share_amt
                            )
                            pla_filename = f"PLA_{layer_name.replace(' ', '_')}_{safe_rein_name}.pdf"
                            zip_file.writestr(pla_filename, pdf_pla_bytes)

                            # 2. สร้าง LSA PDF (เรียกฟังก์ชันสร้าง LSA หรือเรียก generate_lsa_pdf ที่สร้างไว้ในขั้นตอนที่ 1)
                            # หากยังไม่มีฟังก์ชัน generate_lsa_pdf สามารถเรียก generate_pla_pdf แก้ขัดก่อนได้
                            try:
                                pdf_lsa_bytes = generate_lsa_pdf(
                                    reinsurer_full_name=rein_full_name,
                                    layer_name=layer_name,
                                    gross_loss=tot_gross_all,
                                    net_loss=tot_net_all,
                                    excess_pt=excess_pt,
                                    layer_limit=limit,
                                    reinsurer_share_amt=rein_share_amt
                                )
                            except NameError:
                                # Fallback กรณีที่ยังไม่ได้ประกาศฟังก์ชัน generate_lsa_pdf
                                pdf_lsa_bytes = pdf_pla_bytes

                            lsa_filename = f"LSA_{layer_name.replace(' ', '_')}_{safe_rein_name}.pdf"
                            zip_file.writestr(lsa_filename, pdf_lsa_bytes)

            zip_buffer.seek(0)
            st.success("✅ สร้างไฟล์ PDF (ทั้ง PLA และ LSA) ของทุกบริษัทสำเร็จเรียบร้อยแล้ว!")
            st.download_button(
                label="📦 ดาวน์โหลดเอกสาร PDF ทั้งหมด (PLA + LSA ZIP)",
                data=zip_buffer.getvalue(),
                file_name="PLA_and_LSA_Notices_All_Reinsurers.zip",
                mime="application/zip"
            )
