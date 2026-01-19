"""
🚀 OTOMASI SUMMARY EVENT PROMO
Streamlit Web Application
"""

import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

# ============================================================
# KONFIGURASI HALAMAN
# ============================================================
st.set_page_config(
    page_title="Otomasi Summary Promo",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# FUNGSI-FUNGSI UTAMA
# ============================================================

def clean_promo_name(nama_promo):
    """
    Membersihkan tanggal dari nama promo
    Contoh: 'PB HOMECARE FAIR 1-31 JANUARI 2026' → 'PB HOMECARE FAIR'
    """
    if not nama_promo:
        return nama_promo
    
    # Daftar bulan dalam bahasa Indonesia dan Inggris
    bulan = r'(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)'
    
    # Pattern untuk berbagai format tanggal
    date_patterns = [
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s+{bulan}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s+{bulan}\s*\d{{4}}$',
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s*{bulan}\s*\d{{4}}',
    ]
    
    result = nama_promo
    for pattern in date_patterns:
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)
    
    result = ' '.join(result.split()).strip()
    result = re.sub(r'[\s-]+$', '', result).strip()
    
    return result


def extract_promo_info(header_text):
    """
    Mengekstrak informasi promo dari teks header di row 1
    """
    if pd.isna(header_text):
        return None, None
    
    header_text = str(header_text)
    
    promo_match = re.search(r'\d+\s*-\s*([^,]+)', header_text)
    nama_promo = promo_match.group(1).strip() if promo_match else ""
    nama_promo = clean_promo_name(nama_promo)
    
    periode_match = re.search(r',\s*(\d{1,2}\s*(?:-\s*\d{1,2})?\s*\w+\s*\d{4}(?:\s*-\s*\d{1,2}\s*\w+\s*\d{4})?)\s*$', header_text)
    if periode_match:
        periode_text = periode_match.group(1).strip()
    else:
        parts = header_text.rsplit(',', 1)
        periode_text = parts[-1].strip() if len(parts) > 1 else ""
    
    return nama_promo, periode_text


def extract_mekanisme(mek_text):
    """
    Mengekstrak mekanisme promo dari row 2
    """
    if pd.isna(mek_text):
        return ""
    mek_text = str(mek_text)
    mek_text = re.sub(r'^\d+\.\s*', '', mek_text)
    return mek_text.strip()


def parse_periode(periode_text):
    """
    Parse periode promo
    """
    if not periode_text:
        return ""
    return ' '.join(str(periode_text).strip().split())


def safe_convert_number(value):
    """
    Konversi nilai ke angka dengan aman
    """
    if pd.isna(value):
        return None
    try:
        num = float(value)
        return num if num != 0 or str(value) == '0' else None
    except:
        return None


def detect_column_type(df):
    """
    Mendeteksi tipe kolom berdasarkan header
    """
    try:
        row3 = df.iloc[3]
        row4 = df.iloc[4]
        
        for col in range(len(row3)):
            val3 = str(row3.iloc[col]) if pd.notna(row3.iloc[col]) else ""
            val4 = str(row4.iloc[col]) if pd.notna(row4.iloc[col]) else ""
            
            if 'Sales' in val3 or 'Sales' in val4:
                return 'sales_amount'
        
        return 'bonus_qty'
    except:
        return 'sales_amount'


def process_sheet(df, sheet_name):
    """
    Memproses satu sheet dan mengekstrak data summary
    """
    result = {
        'Nama Promo': '',
        'Mekanisme Promo': '',
        'Periode Promo': '',
        'All Count': None,
        'All Claim': None,
        'Sales Amount': None,
        'Amount': None,
        'Left': None
    }
    
    try:
        if len(df) < 7:
            return None
        
        header_text = df.iloc[1, 0]
        mek_text = df.iloc[2, 0]
        
        nama_promo, periode_text = extract_promo_info(header_text)
        mekanisme = extract_mekanisme(mek_text)
        periode = parse_periode(periode_text)
        
        if not nama_promo:
            return None
        
        result['Nama Promo'] = nama_promo
        result['Mekanisme Promo'] = mekanisme
        result['Periode Promo'] = periode
        
        summary_row = df.iloc[6]
        num_cols = len(df.columns)
        
        result['All Count'] = safe_convert_number(summary_row.iloc[3])
        result['All Claim'] = safe_convert_number(summary_row.iloc[4])
        
        col_type = detect_column_type(df)
        
        if col_type == 'sales_amount' and num_cols >= 17:
            result['Sales Amount'] = safe_convert_number(summary_row.iloc[12])
            result['Amount'] = safe_convert_number(summary_row.iloc[13])
            result['Left'] = safe_convert_number(summary_row.iloc[16])
        elif num_cols >= 16:
            result['Sales Amount'] = None
            result['Amount'] = safe_convert_number(summary_row.iloc[12])
            result['Left'] = safe_convert_number(summary_row.iloc[15])
        else:
            result['Sales Amount'] = None
            result['Amount'] = None
            result['Left'] = None
        
        return result
        
    except Exception as e:
        return None


def generate_summary(uploaded_file):
    """
    Membaca semua sheet dan menghasilkan summary
    """
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error membaca file: {str(e)}")
        return None
    
    sheet_names = xl.sheet_names
    results = []
    processed_sheets = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, sheet_name in enumerate(sheet_names):
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            
            if df.empty:
                continue
            
            result = process_sheet(df, sheet_name)
            
            if result and result['Nama Promo']:
                result['No.'] = len(results) + 1
                results.append(result)
                processed_sheets.append(f"✅ {result['Nama Promo'][:50]}")
            
        except Exception as e:
            processed_sheets.append(f"❌ Sheet '{sheet_name}': Error")
        
        # Update progress
        progress_bar.progress((idx + 1) / len(sheet_names))
        status_text.text(f"Memproses sheet {idx + 1} dari {len(sheet_names)}...")
    
    progress_bar.empty()
    status_text.empty()
    
    if len(results) == 0:
        return None, processed_sheets
    
    summary_df = pd.DataFrame(results)
    
    cols = ['No.', 'Nama Promo', 'Mekanisme Promo', 'Periode Promo', 
            'All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    for col in cols:
        if col not in summary_df.columns:
            summary_df[col] = None
    
    return summary_df[cols], processed_sheets


def format_preview_display(df):
    """
    Format DataFrame untuk preview yang rapi
    """
    display_df = df.copy()
    numeric_cols = ['All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    for col in numeric_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(
                lambda x: '{:,.0f}'.format(x) if pd.notna(x) else ''
            )
    
    return display_df


def save_summary_to_excel(df):
    """
    Menyimpan summary ke file Excel dengan format rapi
    Returns: BytesIO object
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    # Styles
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    alt_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin'), right=Side(style='thin'),
        top=Side(style='thin'), bottom=Side(style='thin')
    )
    
    headers = ['No.', 'Nama Promo', 'Mekanisme Promo', 'Periode Promo', 
               'All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    # Header row
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    # Data rows
    for row_idx, row_data in df.iterrows():
        excel_row = row_idx + 2
        
        for col_idx, header in enumerate(headers):
            value = row_data.get(header, '')
            
            if header in ['All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']:
                if pd.notna(value) and value is not None:
                    try:
                        cell = ws.cell(row=excel_row, column=col_idx + 1, value=int(value))
                        cell.number_format = '#,##0'
                    except:
                        cell = ws.cell(row=excel_row, column=col_idx + 1, value='')
                else:
                    cell = ws.cell(row=excel_row, column=col_idx + 1, value='')
            else:
                cell = ws.cell(row=excel_row, column=col_idx + 1, value=value)
            
            cell.border = thin_border
            
            if col_idx == 0:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [4, 5, 6, 7, 8]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            
            if row_idx % 2 == 1:
                cell.fill = alt_fill
    
    # Column widths
    widths = {'A': 5, 'B': 35, 'C': 50, 'D': 28, 'E': 12, 'F': 12, 'G': 18, 'H': 18, 'I': 18}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    
    ws.row_dimensions[1].height = 25
    ws.freeze_panes = 'A2'
    
    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output


# ============================================================
# TAMPILAN APLIKASI
# ============================================================

# Header
st.title("🚀 Otomasi Summary Event Promo")
st.markdown("---")

# Sidebar info
with st.sidebar:
    st.header("📌 Panduan Penggunaan")
    st.markdown("""
    1. Upload file Excel mentah (.xlsx)
    2. Klik tombol **Proses File**
    3. Lihat preview hasil
    4. Download file summary
    """)
    
    st.markdown("---")
    
    st.header("📋 Format Input")
    st.markdown("""
    - File Excel dengan multiple sheets
    - Setiap sheet berisi 1 promo
    - Row 2: Header (nama & periode)
    - Row 3: Mekanisme promo
    - Row 7: Data summary (total)
    """)
    
    st.markdown("---")
    st.caption("Version 4.0 | 2026")

# Main content
col1, col2 = st.columns([2, 1])

with col1:
    st.header("📤 Upload File Excel")
    uploaded_file = st.file_uploader(
        "Pilih file Excel mentah",
        type=['xlsx', 'xls'],
        help="Upload file Excel dengan format yang sesuai"
    )

if uploaded_file is not None:
    st.success(f"✅ File berhasil diupload: **{uploaded_file.name}**")
    
    # Process button
    if st.button("🚀 Proses File", type="primary", use_container_width=True):
        with st.spinner("Memproses file..."):
            result = generate_summary(uploaded_file)
            
            if result[0] is not None:
                summary_df, processed_sheets = result
                
                # Store in session state
                st.session_state['summary_df'] = summary_df
                st.session_state['processed_sheets'] = processed_sheets
                st.session_state['original_filename'] = uploaded_file.name
                
                st.success(f"✅ Berhasil memproses {len(summary_df)} promo!")
            else:
                st.error("❌ Gagal memproses file. Periksa format file input.")

# Display results if available
if 'summary_df' in st.session_state:
    summary_df = st.session_state['summary_df']
    
    st.markdown("---")
    st.header("📊 Preview Hasil")
    
    # Display formatted preview
    display_df = format_preview_display(summary_df)
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
    # Download button
    st.markdown("---")
    st.header("📥 Download Hasil")
    
    excel_file = save_summary_to_excel(summary_df)
    
    original_name = st.session_state.get('original_filename', 'file')
    output_filename = f"Summary_{original_name.replace('.xlsx', '').replace('.xls', '')}_Output.xlsx"
    
    st.download_button(
        label="📥 Download Summary Excel",
        data=excel_file,
        file_name=output_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True
    )
    
    # Show processing details
    with st.expander("📋 Detail Proses"):
        for item in st.session_state.get('processed_sheets', []):
            st.text(item)

# Footer
st.markdown("---")
st.caption("Otomasi Summary Event Promo")
