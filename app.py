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
# FUNGSI-FUNGSI UTAMA (REVISI V2)
# ============================================================

def clean_promo_name(nama_promo):
    """
    Membersihkan tanggal dari nama promo secara agresif.
    Mendukung format tahun 2 digit (25) dan 4 digit (2025).
    """
    if not nama_promo:
        return nama_promo
    
    # Daftar bulan lengkap (Indo & English) + Singkatan
    bulan = r'(?:JANUARI|FEBRUARI|MARET|APRIL|MEI|JUNI|JULI|AGUSTUS|SEPTEMBER|OKTOBER|NOVEMBER|DESEMBER|JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC|JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)'
    
    # Regex Patterns diperbarui untuk support 2 digit tahun (\d{2,4})
    date_patterns = [
        # Format: 20-26 AUG 25 atau 20-26 AGUSTUS 2025
        rf'\s*\d{{1,2}}\s*[-–]\s*\d{{1,2}}\s+{bulan}\s*\d{{2,4}}',
        # Format: 1 JAN - 31 JAN 25
        rf'\s*\d{{1,2}}\s+{bulan}\s*[-–]\s*\d{{1,2}}\s+{bulan}\s*\d{{2,4}}',
        # Format: AGUSTUS 2025
        rf'\s+{bulan}\s*\d{{2,4}}$',
        # Format Slash: 01/01/25
        rf'\s*\d{{1,2}}/\d{{1,2}}/\d{{2,4}}',
        # Format di dalam kurung: (20-25 AUG)
        rf'\s*\(\d.*?\)$'
    ]
    
    result = str(nama_promo)
    for pattern in date_patterns:
        result = re.sub(pattern, '', result, flags=re.IGNORECASE)
    
    # Bersihkan sisa karakter aneh di ujung string
    result = re.sub(r'[,–\-\.]+$', '', result)
    result = ' '.join(result.split()).strip()
    
    return result

def extract_nama_promo(text):
    if not text or not isinstance(text, str):
        return ""

    s = text.upper().strip()

    # Hapus metadata (Last Claim, dll) dan teks dalam kurung
    s = re.sub(r'\(.*?\)', '', s)

    # Pola: Ambil teks setelah angka urut dan dash (misal: "382 - NAMA PROMO")
    # Atau ambil langsung jika tidak ada dash
    if " - " in s or " – " in s:
        parts = re.split(r'\s*[-–]\s*', s, maxsplit=1)
        if len(parts) > 1:
            s = parts[1]
    
    # Bersihkan tanggal menggunakan fungsi clean_promo_name
    s = clean_promo_name(s)

    return s

def clean_mekanisme(text):
    """Menghapus penomoran di awal (contoh: '1. Beli' -> 'Beli')"""
    if not isinstance(text, str):
        return text
    # Hapus angka diikuti titik/kurung di awal string
    return re.sub(r'^\s*\d+[\.\)\-]\s*', '', text).strip()

def extract_promo_info_flexible(df):
    nama_promo = ""
    periode_text = ""
    mekanisme = ""

    # Scanning 6 baris pertama dan 3 kolom pertama
    for row_idx in range(min(8, len(df))):
        for col_idx in range(min(4, len(df.columns))):
            cell_value = df.iloc[row_idx, col_idx]
            if pd.isna(cell_value):
                continue

            cell_str = str(cell_value).strip()

            # 1. Cari Nama Promo
            if not nama_promo:
                # Logika: Jika mengandung kata PROMO atau huruf kapital semua yang panjang
                temp_name = extract_nama_promo(cell_str)
                if "PROMO" in temp_name or (len(temp_name) > 10 and row_idx < 2):
                    nama_promo = temp_name
                    # Jangan continue, karena di sel yang sama mungkin ada periode (jarang tapi mungkin)

            # 2. Cari Periode Promo (Revisi Regex untuk tahun 2 digit)
            if not periode_text:
                # Pola: tanggal - tanggal bulan tahun (2/4 digit)
                periode_match = re.search(
                    r'(\d{1,2}\s*[-–]\s*\d{1,2}\s+\w+\s+\d{2,4})|(\d{1,2}\s+\w+\s*[-–]\s*\d{1,2}\s+\w+\s+\d{2,4})',
                    cell_str, re.IGNORECASE
                )
                if periode_match:
                    periode_text = periode_match.group(0)

            # 3. Cari Mekanisme
            if row_idx > 0 and not mekanisme:
                lower_str = cell_str.lower()
                keywords = ['beli', 'min', 'gratis', 'disc', 'potongan', 'bonus', 'cashback', 'free', 'mendapatkan']
                if any(kw in lower_str for kw in keywords):
                    mekanisme = clean_mekanisme(cell_str)

    # Fallback mekanisme: ambil sel non-kosong di baris ke-3 jika belum ketemu
    if not mekanisme and len(df) > 2:
        mekanisme = clean_mekanisme(str(df.iloc[2, 0]))

    return nama_promo, periode_text, mekanisme
    
def find_header_row(df):
    """Mencari row header"""
    header_keywords = ['no', 'customer', 'count', 'claim', 'sales', 'amount', 'left', 'bonus', 'qty', 'total']
    
    for row_idx in range(min(12, len(df))):
        try:
            row_values = [str(v).lower().strip() for v in df.iloc[row_idx] if pd.notna(v)]
            matches = sum(1 for kw in header_keywords if any(kw in val for val in row_values))
            if matches >= 2: # Cukup 2 match agar lebih sensitif
                return row_idx
        except Exception:
            continue
    return None

def find_summary_row(df, header_row):
    """
    Mencari row summary.
    REVISI: Jika tidak ketemu kata 'Total'/'All', ambil baris TERAKHIR yang memiliki data.
    """
    start_search = (header_row + 1) if header_row is not None else 5
    
    # 1. Cari eksplisit kata "Total" atau "All"
    for row_idx in range(start_search, len(df)):
        try:
            first_val = str(df.iloc[row_idx, 0]).lower() if pd.notna(df.iloc[row_idx, 0]) else ""
            if 'all' in first_val or 'total' in first_val or 'grand' in first_val:
                return row_idx
        except:
            continue

    # 2. Fallback Ultimate: Ambil baris terakhir dari dataframe
    # Asumsinya report Excel pasti row paling bawah adalah total
    return len(df) - 1

def find_column_by_keywords(df, header_row, keywords):
    """Mencari indeks kolom berdasarkan keywords"""
    if header_row is None:
        return None
    
    # Cek header_row dan baris sebelumnya (handling merged cell)
    for check_row in [header_row, header_row - 1]:
        if check_row < 0 or check_row >= len(df): continue
        
        for col_idx in range(len(df.columns)):
            cell_val = df.iloc[check_row, col_idx]
            if pd.isna(cell_val): continue
            
            if any(kw in str(cell_val).lower() for kw in keywords):
                return col_idx
    return None

def safe_convert_number(value):
    if pd.isna(value): return None
    try:
        val_str = str(value)
        # Hapus 'Rp', titik sebagai ribuan (jika format indo), dll.
        # Asumsi input Excel standar (titik/koma tergantung locale, kita ambil digitnya saja)
        clean_val = re.sub(r'[^\d,\.\-]', '', val_str)
        if ',' in clean_val and '.' in clean_val: 
            clean_val = clean_val.replace(',', '') # format US 1,000.00 -> 1000.00
        elif ',' in clean_val:
            clean_val = clean_val.replace(',', '.') # format Indo 1000,00 -> 1000.00
            
        return float(clean_val)
    except:
        return None

def process_sheet_robust(df, sheet_name):
    result = {
        'Nama Promo': '', 'Mekanisme Promo': '', 'Periode Promo': '',
        'All Count': None, 'All Claim': None, 'Sales Amount': None, 
        'Amount': None, 'Left': None
    }
    
    try:
        if len(df) < 2: return None, "❌ Terlalu sedikit baris"
        
        # 1. Extract Info Teks
        nama, periode, mekanisme = extract_promo_info_flexible(df)
        
        # JIKA NAMA KOSONG, GUNAKAN NAMA SHEET SEBAGAI CADANGAN
        if not nama:
            nama = clean_promo_name(sheet_name)
        
        result['Nama Promo'] = nama
        result['Mekanisme Promo'] = mekanisme
        result['Periode Promo'] = periode
        
        # 2. Cari Data Angka
        header_row = find_header_row(df)
        summary_row_idx = find_summary_row(df, header_row)
        
        summary_row = df.iloc[summary_row_idx]
        
        # Mapping Kolom
        count_col = find_column_by_keywords(df, header_row, ['count', 'jumlah cust', 'total cust'])
        claim_col = find_column_by_keywords(df, header_row, ['claim', 'klaim'])
        sales_col = find_column_by_keywords(df, header_row, ['sales amount', 'sales amt', 'penjualan'])
        amount_col = find_column_by_keywords(df, header_row, ['amount', 'bonus amt', 'nilai bonus'])
        left_col = find_column_by_keywords(df, header_row, ['left', 'sisa'])
        
        # Ambil Value (dengan Fallback posisi kolom jika header tidak ketemu)
        cols_cnt = len(df.columns)
        
        if count_col is not None: result['All Count'] = safe_convert_number(summary_row.iloc[count_col])
        elif cols_cnt > 3: result['All Count'] = safe_convert_number(summary_row.iloc[3])
            
        if claim_col is not None: result['All Claim'] = safe_convert_number(summary_row.iloc[claim_col])
        elif cols_cnt > 4: result['All Claim'] = safe_convert_number(summary_row.iloc[4])
            
        if sales_col is not None: result['Sales Amount'] = safe_convert_number(summary_row.iloc[sales_col])
        
        if amount_col is not None: result['Amount'] = safe_convert_number(summary_row.iloc[amount_col])
        elif cols_cnt > 12: result['Amount'] = safe_convert_number(summary_row.iloc[cols_cnt-3]) # Seringkali di ujung
            
        if left_col is not None: result['Left'] = safe_convert_number(summary_row.iloc[left_col])
        elif cols_cnt > 1: result['Left'] = safe_convert_number(summary_row.iloc[cols_cnt-1]) # Seringkali kolom terakhir
        
        return result, f"✅ {nama[:30]}..."
        
    except Exception as e:
        # PENTING: Jangan return None jika error, return apa yang ada
        # agar sheet tetap terhitung meski data angka kosong
        if result['Nama Promo']:
            return result, f"⚠️ Partial: {str(e)[:20]}"
        return None, f"❌ Error: {str(e)[:30]}"


def generate_summary(uploaded_file):
    try:
        xl = pd.ExcelFile(uploaded_file)
    except Exception as e:
        return None, [f"Fatal Error: {str(e)}"]
    
    results = []
    logs = []
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    sheet_names = xl.sheet_names
    total_sheets = len(sheet_names)
    
    for i, sheet in enumerate(sheet_names):
        # Update progress bar setiap 10 sheet agar tidak lag
        if i % 10 == 0:
            progress_bar.progress((i + 1) / total_sheets)
            status_text.text(f"Processing {i+1}/{total_sheets}: {sheet}")
            
        try:
            df = pd.read_excel(xl, sheet_name=sheet, header=None)
            res, msg = process_sheet_robust(df, sheet)
            
            if res:
                res['No.'] = len(results) + 1
                results.append(res)
                logs.append(msg)
            else:
                logs.append(f"⏭️ Skip {sheet}: {msg}")
                
        except Exception as e:
            logs.append(f"❌ Fail {sheet}: {e}")
            
    progress_bar.empty()
    status_text.empty()
    
    if not results:
        return None, logs
        
    summary_df = pd.DataFrame(results)
    cols = ['No.', 'Nama Promo', 'Mekanisme Promo', 'Periode Promo', 
            'All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    
    # Pastikan semua kolom ada
    for c in cols:
        if c not in summary_df.columns: summary_df[c] = None
            
    return summary_df[cols], logs

def format_preview_display(df):
    display_df = df.copy()
    numeric_cols = ['All Count', 'All Claim', 'Sales Amount', 'Amount', 'Left']
    for col in numeric_cols:
        if col in display_df.columns:
            display_df[col] = display_df[col].apply(lambda x: '{:,.0f}'.format(x) if pd.notna(x) else '')
    return display_df

def save_summary_to_excel(df):
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    
    # Styling
    header_style = Font(bold=True, color="FFFFFF")
    fill_style = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border_style = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    headers = list(df.columns)
    
    # Write Header
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.font = header_style
        cell.fill = fill_style
        cell.alignment = center_align
        cell.border = border_style
        
    # Write Data
    for row_num, row_data in enumerate(df.values, 2):
        for col_num, value in enumerate(row_data, 1):
            cell = ws.cell(row=row_num, column=col_num, value=value)
            cell.border = border_style
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            
            # Format number jika perlu
            if col_num > 4 and pd.notna(value): # Kolom angka (No, Nama, Mekanisme, Periode = 4 pertama)
                 try:
                     cell.value = float(value)
                     cell.number_format = '#,##0'
                 except: pass

    # Adjust Width
    ws.column_dimensions['B'].width = 40 # Nama Promo
    ws.column_dimensions['C'].width = 50 # Mekanisme
    ws.column_dimensions['D'].width = 20 # Periode
    
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ============================================================
# UI UTAMA
# ============================================================

st.title("🚀 Otomasi Summary Event Promo v2.0")
st.caption("Auto-clean dates, auto-detect summary rows")
st.markdown("---")

uploaded_file = st.file_uploader("Upload File Excel (.xlsx)", type=['xlsx'])

if uploaded_file:
    if st.button("Proses File", type="primary"):
        with st.spinner("Sedang memproses..."):
            df_result, logs = generate_summary(uploaded_file)
            
            if df_result is not None:
                st.session_state['data'] = df_result
                st.success(f"Selesai! Berhasil mengambil {len(df_result)} promo.")
            else:
                st.error("Gagal memproses file.")
                st.write(logs)

if 'data' in st.session_state:
    df = st.session_state['data']
    st.subheader("Preview Hasil")
    st.dataframe(format_preview_display(df), use_container_width=True, hide_index=True)
    
    excel_data = save_summary_to_excel(df)
    st.download_button(
        label="📥 Download Excel",
        data=excel_data,
        file_name="Summary_Promo_Clean.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
