"""
🚀 OTOMASI SUMMARY EVENT PROMO
Streamlit Web Application
Version - Improved Robustness
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
# FUNGSI-FUNGSI UTAMA (IMPROVED VERSION)
# ============================================================

def clean_promo_name(nama_promo):
    """
    Membersihkan tanggal dari nama promo
    Contoh: 'PB HOMECARE FAIR 1-31 JANUARI 2026' → 'PB HOMECARE FAIR'
    """
    if not nama_promo:
        return nama_promo
    
    bulan = r'(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)'
    
    date_patterns = [
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s+{bulan}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}\s+{bulan}\s*\d{{4}}',
        rf'\s+{bulan}\s*\d{{4}}$',
        rf'\s*\d{{1,2}}\s*-\s*\d{{1,2}}\s*{bulan}\s*\d{{4}}',
        rf'\s*\d{{1,2}}/{bulan}/\d{{4}}',
        rf'\s*\d{{1,2}}/\d{{1,2}}/\d{{4}}',
    ]
    
    result = nama_promo
    for pattern in date_patterns:
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)
    
    result = ' '.join(result.split()).strip()
    result = re.sub(r'[\s-]+$', '', result).strip()
    
    return result

def extract_nama_promo(text):
    if not text or not isinstance(text, str):
        return ""

    s = text.upper().strip()

    # Hapus metadata (Last Claim, dll)
    s = re.sub(r'\(.*?\)', '', s)

    # Ambil teks setelah "NN -"
    m = re.search(r'\b\d+\s*[-–]\s*(.+)', s)
    if not m:
        return ""

    s = m.group(1)

    # Bersihkan tanggal di belakang pakai fungsi Anda
    s = clean_promo_name(s)

    # Rapikan
    s = re.sub(r'[,–\-]+$', '', s)
    s = re.sub(r'\s+', ' ', s).strip()

    return s

def extract_promo_info_flexible(df):
    nama_promo = ""
    periode_text = ""
    mekanisme = ""

    for row_idx in range(min(6, len(df))):
        for col_idx in range(min(3, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx]
            if pd.isna(cell_value):
                continue

            cell_str = str(cell_value).strip()

            if not nama_promo:
                extracted = extract_nama_promo(cell_str)
                if extracted and extracted.startswith("PROMO"):
                    nama_promo = extracted
                    continue

            if not periode_text:
                periode_match = re.search(
                    r'(\d{1,2}\s*[-–]\s*\d{1,2}\s+\w+\s+\d{4})',
                    cell_str, re.IGNORECASE
                )
                if periode_match:
                    periode_text = periode_match.group(1)

            if row_idx > 0 and not mekanisme:
                if any(kw in cell_str.lower() for kw in [
                    'beli', 'min', 'gratis', 'disc', 'potongan',
                    'bonus', 'cashback', 'free'
                ]):
                    mekanisme = cell_str

    if not mekanisme and len(df) > 2:
        mekanisme = str(df.iloc[2, 0]).strip()

    return nama_promo, periode_text, mekanisme
    
def find_header_row(df):
    """
    Mencari row yang berisi header kolom data
    Biasanya berisi 'No', 'Customer', 'Count', 'Claim', dll
    """
    header_keywords = ['no', 'customer', 'count', 'claim', 'sales', 'amount', 'left', 'bonus', 'qty', 'total']
    
    for row_idx in range(min(10, len(df))):
        try:
            row_values = [str(v).lower().strip() for v in df.iloc[row_idx] if pd.notna(v)]
            matches = sum(1 for kw in header_keywords if any(kw in val for val in row_values))
            if matches >= 3:
                return row_idx
        except Exception:
            continue
    
    return None


def find_summary_row(df, header_row):
    """
    Mencari row yang berisi data summary (total)
    Biasanya row pertama setelah header yang memiliki nilai numerik
    """
    if header_row is None:
        # Fallback ke row 6 (0-indexed)
        return 6 if len(df) > 6 else None
    
    # Cari row dengan 'All' atau 'Total' atau row pertama dengan data numerik
    for row_idx in range(header_row + 1, min(header_row + 5, len(df))):
        try:
            row_values = df.iloc[row_idx]
            first_val = str(row_values.iloc[0]).lower() if pd.notna(row_values.iloc[0]) else ""
            
            if 'all' in first_val or 'total' in first_val:
                return row_idx
            
            # Cek apakah row ini memiliki beberapa nilai numerik
            numeric_count = sum(1 for v in row_values if pd.notna(v) and isinstance(v, (int, float)))
            if numeric_count >= 3:
                return row_idx
        except Exception:
            continue
    
    return header_row + 1 if header_row is not None else 6


def find_column_by_keywords(df, header_row, keywords):
    """
    Mencari indeks kolom berdasarkan keywords di header
    """
    if header_row is None or header_row >= len(df):
        return None
    
    try:
        # Cek di header_row dan header_row+1 (untuk merged headers)
        for check_row in [header_row, header_row - 1]:
            if check_row < 0 or check_row >= len(df):
                continue
            
            for col_idx in range(len(df.columns)):
                cell_value = df.iloc[check_row, col_idx]
                if pd.isna(cell_value):
                    continue
                
                cell_str = str(cell_value).lower().strip()
                if any(kw in cell_str for kw in keywords):
                    return col_idx
    except Exception:
        pass
    
    return None


def safe_convert_number(value):
    """
    Konversi nilai ke angka dengan aman
    """
    if pd.isna(value):
        return None
    try:
        if isinstance(value, str):
            # Hapus karakter non-numerik kecuali titik dan minus
            value = re.sub(r'[^\d.\-]', '', value.replace(',', ''))
        num = float(value)
        return num if not np.isnan(num) else None
    except (ValueError, TypeError):
        return None


def process_sheet_robust(df, sheet_name):
    """
    Memproses satu sheet dengan deteksi yang lebih robust
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
        if len(df) < 5:
            return None, f"❌ Sheet '{sheet_name}': Terlalu sedikit baris"
        
        # Extract info promo
        nama_promo, periode_text, mekanisme = extract_promo_info_flexible(df)
        
        if not nama_promo:
            return None, f"❌ Sheet '{sheet_name}': Nama promo tidak ditemukan"
        
        result['Nama Promo'] = nama_promo
        result['Mekanisme Promo'] = mekanisme
        result['Periode Promo'] = periode_text
        
        # Cari header row dan summary row
        header_row = find_header_row(df)
        summary_row_idx = find_summary_row(df, header_row)
        
        if summary_row_idx is None or summary_row_idx >= len(df):
            return None, f"❌ Sheet '{sheet_name}': Summary row tidak ditemukan"
        
        summary_row = df.iloc[summary_row_idx]
        num_cols = len(df.columns)
        
        # Metode 1: Cari kolom berdasarkan keyword di header
        count_col = find_column_by_keywords(df, header_row, ['count', 'jumlah customer', 'total customer'])
        claim_col = find_column_by_keywords(df, header_row, ['claim', 'klaim'])
        sales_col = find_column_by_keywords(df, header_row, ['sales amount', 'sales amt', 'nilai penjualan', 'sales'])
        amount_col = find_column_by_keywords(df, header_row, ['amount', 'bonus amount', 'nilai bonus', 'bonus amt'])
        left_col = find_column_by_keywords(df, header_row, ['left', 'sisa', 'remaining'])
        
        # Metode 2: Fallback ke posisi tetap jika tidak ditemukan
        if count_col is not None:
            result['All Count'] = safe_convert_number(summary_row.iloc[count_col])
        elif num_cols > 3:
            result['All Count'] = safe_convert_number(summary_row.iloc[3])
        
        if claim_col is not None:
            result['All Claim'] = safe_convert_number(summary_row.iloc[claim_col])
        elif num_cols > 4:
            result['All Claim'] = safe_convert_number(summary_row.iloc[4])
        
        # Untuk Sales Amount, Amount, dan Left - coba berbagai pendekatan
        if sales_col is not None:
            result['Sales Amount'] = safe_convert_number(summary_row.iloc[sales_col])
        
        if amount_col is not None:
            result['Amount'] = safe_convert_number(summary_row.iloc[amount_col])
        
        if left_col is not None:
            result['Left'] = safe_convert_number(summary_row.iloc[left_col])
        
        # Fallback berdasarkan jumlah kolom (untuk kompatibilitas dengan format lama)
        if result['Amount'] is None or result['Left'] is None:
            # Deteksi apakah ada kolom Sales Amount
            has_sales = sales_col is not None or any(
                'sales' in str(v).lower() 
                for v in df.iloc[header_row] if pd.notna(v)
            ) if header_row is not None else False
            
            if has_sales and num_cols >= 17:
                if result['Sales Amount'] is None:
                    result['Sales Amount'] = safe_convert_number(summary_row.iloc[12])
                if result['Amount'] is None:
                    result['Amount'] = safe_convert_number(summary_row.iloc[13])
                if result['Left'] is None:
                    result['Left'] = safe_convert_number(summary_row.iloc[16])
            elif num_cols >= 16:
                if result['Amount'] is None:
                    result['Amount'] = safe_convert_number(summary_row.iloc[12])
                if result['Left'] is None:
                    result['Left'] = safe_convert_number(summary_row.iloc[15])
            elif num_cols >= 13:
                # Untuk sheet dengan kolom lebih sedikit
                if result['Amount'] is None:
                    result['Amount'] = safe_convert_number(summary_row.iloc[-3])
                if result['Left'] is None:
                    result['Left'] = safe_convert_number(summary_row.iloc[-1])
        
        return result, f"✅ {nama_promo[:50]}"
        
    except Exception as e:
        return None, f"❌ Sheet '{sheet_name}': Error - {str(e)[:50]}"


def generate_summary(uploaded_file):
    """
    Membaca semua sheet dan menghasilkan summary
    """
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"❌ Error membaca file: {str(e)}")
        return None, []
    
    sheet_names = xl.sheet_names
    results = []
    processed_sheets = []
    skipped_sheets = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, sheet_name in enumerate(sheet_names):
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name, header=None)
            
            if df.empty or len(df) < 3:
                skipped_sheets.append(f"⏭️ Sheet '{sheet_name}': Kosong atau terlalu sedikit data")
                continue
            
            result, message = process_sheet_robust(df, sheet_name)
            
            if result and result['Nama Promo']:
                result['No.'] = len(results) + 1
                results.append(result)
                processed_sheets.append(f"✅ [{len(results):02d}] {result['Nama Promo'][:50]}")
            else:
                skipped_sheets.append(message)
            
        except Exception as e:
            skipped_sheets.append(f"❌ Sheet '{sheet_name}': Exception - {str(e)[:50]}")
        
        progress_bar.progress((idx + 1) / len(sheet_names))
        status_text.text(f"Memproses sheet {idx + 1} dari {len(sheet_names)}...")
    
    progress_bar.empty()
    status_text.empty()
    
    all_logs = processed_sheets + skipped_sheets
    
    if len(results) == 0:
        return None, all_logs
    
    summary_df = pd.DataFrame(results)
    
    cols = ['No.', 'Nama Promo', 'Mekanisme Promo', 'Periode Promo', 
            'All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    for col in cols:
        if col not in summary_df.columns:
            summary_df[col] = None
    
    return summary_df[cols], all_logs


def format_preview_display(df):
    """
    Format DataFrame untuk preview yang rapi
    """
    display_df = df.copy()
    numeric_cols = ['All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    for col in numeric_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(
                lambda x: '{:,.0f}'.format(x) if pd.notna(x) and x is not None else ''
            )
    
    return display_df


def save_summary_to_excel(df):
    """
    Menyimpan summary ke file Excel dengan format rapi
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
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
    
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_alignment
        cell.border = thin_border
    
    for row_idx, row_data in df.iterrows():
        excel_row = row_idx + 2
        
        for col_idx, header in enumerate(headers):
            value = row_data.get(header, '')
            
            if header in ['All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']:
                if pd.notna(value) and value is not None:
                    try:
                        cell = ws.cell(row=excel_row, column=col_idx + 1, value=int(float(value)))
                        cell.number_format = '#,##0'
                    except (ValueError, TypeError):
                        cell = ws.cell(row=excel_row, column=col_idx + 1, value='')
                else:
                    cell = ws.cell(row=excel_row, column=col_idx + 1, value='')
            else:
                cell = ws.cell(row=excel_row, column=col_idx + 1, value=value if pd.notna(value) else '')
            
            cell.border = thin_border
            
            if col_idx == 0:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif col_idx in [4, 5, 6, 7, 8]:
                cell.alignment = Alignment(horizontal="right", vertical="center")
            else:
                cell.alignment = Alignment(vertical="center", wrap_text=True)
            
            if row_idx % 2 == 1:
                cell.fill = alt_fill
    
    widths = {'A': 5, 'B': 35, 'C': 50, 'D': 28, 'E': 12, 'F': 12, 'G': 18, 'H': 18, 'I': 18}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    
    ws.row_dimensions[1].height = 25
    ws.freeze_panes = 'A2'
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output


# ============================================================
# TAMPILAN APLIKASI
# ============================================================

st.title("🚀 Otomasi Summary Event Promo")
st.markdown("---")

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
    - Row awal: Header (nama & periode)
    - Row berikutnya: Mekanisme promo
    - Ada baris dengan data summary/total
    """)
    
    st.markdown("---")
    
    st.header("🔄 Fitur v5.0")
    st.markdown("""
    - ✅ Deteksi kolom otomatis
    - ✅ Fleksibel untuk berbagai format
    - ✅ Error handling lebih baik
    - ✅ Log proses detail
    """)
    
    st.markdown("---")
    st.caption("Version 5.0 | 2026")

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
    
    if st.button("🚀 Proses File", type="primary", use_container_width=True):
        with st.spinner("Memproses file..."):
            summary_df, process_logs = generate_summary(uploaded_file)
            
            if summary_df is not None:
                st.session_state['summary_df'] = summary_df
                st.session_state['process_logs'] = process_logs
                st.session_state['original_filename'] = uploaded_file.name
                
                success_count = len(summary_df)
                total_sheets = len(process_logs)
                st.success(f"✅ Berhasil memproses {success_count} dari {total_sheets} sheet!")
            else:
                st.session_state['process_logs'] = process_logs
                st.error("❌ Gagal memproses file. Periksa format file input.")

if 'summary_df' in st.session_state:
    summary_df = st.session_state['summary_df']
    
    st.markdown("---")
    st.header("📊 Preview Hasil")
    
    display_df = format_preview_display(summary_df)
    st.dataframe(display_df, use_container_width=True, hide_index=True)
    
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

# Selalu tampilkan detail proses jika ada
if 'process_logs' in st.session_state:
    with st.expander("📋 Detail Proses", expanded=False):
        logs = st.session_state.get('process_logs', [])
        success_logs = [l for l in logs if l.startswith('✅')]
        failed_logs = [l for l in logs if not l.startswith('✅')]
        
        if success_logs:
            st.markdown("**Berhasil diproses:**")
            for item in success_logs:
                st.text(item)
        
        if failed_logs:
            st.markdown("**Dilewati/Error:**")
            for item in failed_logs:
                st.text(item)

st.markdown("---")
st.caption("Otomasi Summary Event Promo")
