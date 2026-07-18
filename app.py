import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
import plotly.io as pio
from io import BytesIO
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
import tempfile
import os

# -------------------------- Fungsi Bantu --------------------------
def normalisasi_nopin(val):
    if pd.isna(val): return ''
    return re.sub(r'[^A-Z0-9]', '', str(val).upper())

def normalisasi_kecamatan(val):
    if pd.isna(val) or str(val).strip() == '': return 'Tidak Diketahui'
    val = str(val).strip().title()
    if 'batam' in val.lower() and 'kota' in val.lower(): return 'Batam Kota'
    return val

def cari_kolom(daftar_kolom, kata_kunci):
    for col in daftar_kolom:
        if any(kw in str(col).upper() for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file Excel...")
def baca_semua_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('.xlsx') else None)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name, header=None)
            if not df.empty: sheets[name] = df
        except Exception: pass
    return sheets

# -------------------------- Parsing Waktu Fleksibel --------------------------
def parse_waktu(series):
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    if series.dtype == object:
        series = series.astype(str).str.strip().str.replace(r'\.', ':', regex=True)
    for fmt in ['%H:%M:%S', '%H:%M', '%H.%M.%S', '%I:%M:%S %p', '%I:%M %p', '%H.%M']:
        dt = pd.to_datetime(series, format=fmt, errors='coerce')
        if dt.notna().sum() > 0: return dt
    if pd.api.types.is_numeric_dtype(series):
        try:
            hours = series * 24
            td = pd.to_timedelta(hours, unit='h')
            base = pd.Timestamp('2026-01-01')
            return base + td
        except: pass
    return pd.to_datetime(series, errors='coerce', dayfirst=True)

def hitung_durasi(df_master):
    keywords_masuk = ['MASUK', 'JAM MASUK', 'WAKTU MASUK', 'TIMBANG MASUK', 'JAM_1', 'TIMBANG1',
                      'BERANGKAT', 'START', 'IN', 'TIME IN', 'JAM', 'WAKTU', 'ENTRY', 'JAM MASUK TPA']
    keywords_keluar = ['KELUAR', 'JAM KELUAR', 'WAKTU KELUAR', 'TIMBANG KELUAR', 'JAM_2', 'TIMBANG2',
                       'TIBA', 'END', 'OUT', 'TIME OUT', 'EXIT', 'JAM KELUAR TPA']

    col_masuk, col_keluar = None, None
    for col in df_master.columns:
        col_upper = str(col).upper()
        if not col_masuk and any(kw in col_upper for kw in keywords_masuk): col_masuk = col
        if not col_keluar and any(kw in col_upper for kw in keywords_keluar): col_keluar = col
        if col_masuk and col_keluar: break

    if not col_masuk or not col_keluar:
        st.info("ℹ️ Kolom jam masuk/keluar tidak ditemukan. Analisis durasi dilewati.")
        return df_master

    df_master['MASUK_ORI'] = df_master[col_masuk].astype(str)
    df_master['KELUAR_ORI'] = df_master[col_keluar].astype(str)

    st.write(f"✅ Kolom waktu terdeteksi: **{col_masuk}** (masuk), **{col_keluar}** (keluar)")
    with st.expander("🔍 Lihat 10 data waktu mentah (hanya 2 kolom)"):
        st.dataframe(df_master[[col_masuk, col_keluar]].head(10))

    dt_masuk = parse_waktu(df_master[col_masuk])
    dt_keluar = parse_waktu(df_master[col_keluar])
    valid_masuk = dt_masuk.notna().sum()
    valid_keluar = dt_keluar.notna().sum()
    st.write(f"Berhasil parsing: Masuk **{valid_masuk}**, Keluar **{valid_keluar}**")

    if valid_masuk > 0 and valid_keluar > 0:
        df_master['MASUK_DT'] = dt_masuk
        df_master['KELUAR_DT'] = dt_keluar
        df_master['DURASI_MENIT'] = (df_master['KELUAR_DT'] - df_master['MASUK_DT']).dt.total_seconds() / 60
        mask_valid = (df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)
        df_master = df_master[mask_valid].copy()
        try: df_master['JAM_INPUT'] = df_master['MASUK_DT'].dt.hour
        except: pass
        st.success(f"✅ Durasi berhasil dihitung. Tersisa {len(df_master)} baris dengan durasi valid.")
    else:
        st.warning("⚠️ Gagal mengonversi waktu. Periksa format di Excel (contoh: 08:30:00 atau 08.30).")
    return df_master

# -------------------------- Proses Data Utama --------------------------
def proses_data(sheets_dict, config, use_master=False, master_sheet=None):
    if use_master and master_sheet:
        if master_sheet not in sheets_dict: return None
        df_master = sheets_dict[master_sheet].copy()
        first_row = df_master.iloc[0].astype(str).str.upper().values
        if any('NOPIN' in str(x) or 'PINTU' in str(x) for x in first_row):
            df_master.columns = [str(c).strip().upper() for c in first_row]
            df_master = df_master.iloc[1:].reset_index(drop=True)
        col_nopin = cari_kolom(df_master.columns, ['NOPIN', 'PINTU'])
        col_plat = cari_kolom(df_master.columns, ['PLAT', 'NOPOL'])
        if not col_nopin: st.error("Kolom NOPIN tidak ditemukan."); return None
        df_master = df_master.rename(columns={col_nopin: 'NOPIN'})
        if col_plat: df_master = df_master.rename(columns={col_plat: 'NO_PLAT'})
        else: df_master['NO_PLAT'] = ''
        df_master = df_master.dropna(subset=['NOPIN'])
        df_master['NOPIN'] = df_master['NOPIN'].astype(str).str.strip().str.upper()
        df_master = df_master[~df_master['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_master = df_master[df_master['NOPIN'] != '']
        df_master['NOPIN'] = df_master['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
        if 'Kecamatan' in df_master.columns: df_master['Kecamatan'] = df_master['Kecamatan'].apply(normalisasi_kecamatan)
        else: df_master['Kecamatan'] = 'Tidak Diketahui'
        # Pastikan TANGGAL ada
        if 'TANGGAL' not in df_master.columns: df_master['TANGGAL'] = ''
        # Hapus duplikat hanya jika kolom tersedia
        key_cols = ['NOPIN', 'TANGGAL', 'NO_PLAT', 'Kecamatan']
        ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
        if ton_col: key_cols.append(ton_col)
        available_keys = [c for c in key_cols if c in df_master.columns]
        if available_keys: df_master.drop_duplicates(subset=available_keys, keep='first', inplace=True)
        if ton_col: df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
        col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col
        df_master = hitung_durasi(df_master)
        return hitung_agregasi(df_master, col_netto)

    # Mode Otomatis / Manual
    armada_sheet = config.get('armada_sheet')
    daily_sheets = config.get('daily_sheets', [])

    # 1. Membaca referensi List Armada
    ref_dict = {}
    if armada_sheet and armada_sheet in sheets_dict:
        df_arm = sheets_dict[armada_sheet].copy()
        header_arm = 1
        try:
            # Gunakan uploaded_file dari session state
            if 'uploaded_file' in st.session_state and st.session_state.uploaded_file:
                xls = pd.ExcelFile(st.session_state.uploaded_file)
                df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_arm)
            else:
                raise FileNotFoundError
        except:
            # fallback: ambil dari sheets_dict
            df_ref = df_arm.iloc[header_arm:].reset_index(drop=True)
            if header_arm > 0:
                header_row = df_arm.iloc[header_arm-1].astype(str).str.strip().str.upper()
                df_ref.columns = [str(c).strip().upper() for c in header_row]
            else:
                df_ref.columns = [str(c).strip().upper() for c in df_arm.iloc[0]]

        df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]
        col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
        col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
        col_kec = cari_kolom(df_ref.columns, ['LOKASI KECAMATAN', 'KECAMATAN', 'LOKASI', 'KEC'])
        col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
        col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
        if col_nopin and col_plat:
            df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
            df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
            for _, row in df_ref.iterrows():
                key = row['NOPIN']
                ref_dict[key] = {'NO.PLAT': row['NO.PLAT']}
                if col_kec: ref_dict[key]['Kecamatan'] = str(row[col_kec]).strip()
                if col_merk: ref_dict[key]['MERK'] = str(row[col_merk]).strip()
                if col_type: ref_dict[key]['TYPE'] = str(row[col_type]).strip()

    # 2. Proses sheet harian
    cleaned_sheets = {}
    skipped = []
    for sheet in daily_sheets:
        if sheet not in sheets_dict: skipped.append(sheet); continue
        df_raw = sheets_dict[sheet].copy()
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str:
                header_idx = idx; break
        if header_idx is None: skipped.append(sheet); continue
        df_hari = df_raw.iloc[header_idx+1:].reset_index(drop=True)
        header_row = df_raw.iloc[header_idx].astype(str).str.strip().str.upper()
        df_hari.columns = [str(c).strip().upper() for c in header_row]

        col_nopin_h = cari_kolom(df_hari.columns, ['PINTU', 'NOPIN'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT'])
        if not col_nopin_h or not col_plat_h: skipped.append(sheet); continue
        df_hari.rename(columns={col_nopin_h: 'NOPIN', col_plat_h: 'NO_PLAT'}, inplace=True)

        # Pembersihan
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
        df_hari = df_hari.reset_index(drop=True)

        # Sinkronisasi dengan master
        if ref_dict:
            ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN'})
            cols_ref = ['NOPIN', 'NO.PLAT']
            for c in ['Kecamatan', 'MERK', 'TYPE']:
                if c in ref_df.columns: cols_ref.append(c)
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns: df_hari.drop(columns=[col], inplace=True)
            df_hari = df_hari.merge(ref_df[cols_ref], on='NOPIN', how='left')
        else:
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns: df_hari[col] = ''

        # Normalisasi kecamatan
        if 'Kecamatan' in df_hari.columns:
            df_hari['Kecamatan'] = df_hari['Kecamatan'].apply(normalisasi_kecamatan)
        else: df_hari['Kecamatan'] = 'Tidak Diketahui'

        # Tanggal
        try: tgl = f"2026-06-{int(sheet):02d}"
        except: tgl = sheet
        df_hari['TANGGAL'] = tgl
        df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]
        cleaned_sheets[sheet] = df_hari

    if not cleaned_sheets: return None

    df_master = pd.concat(cleaned_sheets.values(), ignore_index=True, sort=False)

    # Pastikan kolom kunci ada
    if 'Kecamatan' not in df_master.columns: df_master['Kecamatan'] = 'Tidak Diketahui'
    if 'TANGGAL' not in df_master.columns: df_master['TANGGAL'] = ''

    # Hapus duplikat dengan subset yang ada
    key_cols = ['NOPIN', 'TANGGAL', 'NO_PLAT', 'Kecamatan']
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col: key_cols.append(ton_col)
    available_keys = [c for c in key_cols if c in df_master.columns]
    if available_keys: df_master.drop_duplicates(subset=available_keys, keep='first', inplace=True)

    if ton_col: df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    df_master = hitung_durasi(df_master)
    return hitung_agregasi(df_master, col_netto, cleaned_sheets, skipped)

def hitung_agregasi(df_master, col_netto, cleaned_sheets=None, skipped=None):
    df_armada, teraktif, tidak_efisien = None, None, None
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns: group_cols.append(c)
        df_armada = df_master.groupby(group_cols, dropna=False).agg(
            Total_Trip=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)
        teraktif = df_armada.iloc[0] if not df_armada.empty else None
        tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty and (df_armada['Total_Trip'] > 0).any() else None

    df_waktu = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']
        df_waktu = df_waktu.round(1)

    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)
        if 'DURASI_MENIT' in df_master.columns:
            durasi_kec = df_master.dropna(subset=['DURASI_MENIT']).groupby('Kecamatan', dropna=False)['DURASI_MENIT'].mean().reset_index()
            durasi_kec.columns = ['Kecamatan', 'Rata_Durasi_Menit']
            df_kec = df_kec.merge(durasi_kec, on='Kecamatan', how='left')

    df_type = pd.DataFrame()
    if 'TYPE' in df_master.columns and col_netto:
        df_type = df_master.groupby('TYPE', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)
        if 'DURASI_MENIT' in df_master.columns:
            durasi_type = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
            durasi_type.columns = ['TYPE', 'Rata_Durasi_Menit']
            df_type = df_type.merge(durasi_type, on='TYPE', how='left')

    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('TANGGAL')

    return {
        'df_master': df_master,
        'df_armada': df_armada,
        'teraktif': teraktif,
        'tidak_efisien': tidak_efisien,
        'df_waktu_jenis': df_waktu,
        'df_kec': df_kec,
        'df_type': df_type,
        'df_tren': df_tren,
        'col_netto': col_netto,
        'skipped': skipped if skipped else [],
        'cleaned_count': len(cleaned_sheets) if cleaned_sheets else 1
    }

# -------------------------- Ringkasan Eksekutif & PDF --------------------------
def buat_ringkasan_eksekutif(data):
    df = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']
    df_armada = data['df_armada']
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']

    if 'TANGGAL' in df.columns:
        df['TANGGAL_DT'] = pd.to_datetime(df['TANGGAL'], errors='coerce')
        valid_dates = df['TANGGAL_DT'].dropna()
        bulan_tahun = valid_dates.dt.strftime('%B %Y').iloc[0] if not valid_dates.empty else "Juni 2026"
    else: bulan_tahun = "Juni 2026"

    total_ritase = len(df)
    total_armada = df['NOPIN'].nunique()
    total_tonase_ton = round(df[col_netto].sum() / 1000, 2) if col_netto else 0

    kec_tertinggi = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else "N/A"
    tonase_tertinggi = round(df_kec.iloc[0]['Total_Tonase'] / 1000, 2) if not df_kec.empty else 0

    terbawah = df_armada.sort_values('Total_Trip', ascending=True).head(5) if not df_armada.empty else pd.DataFrame()
    list_terbawah = ", ".join([f"{row['NOPIN']} ({row['NO_PLAT']})" for _, row in terbawah.iterrows()]) if not terbawah.empty else "tidak tersedia"

    durasi_rata = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else 0
    teks = f"""
LAPORAN KESIMPULAN EKSEKUTIF - REKAP TONASE DLH BATAM ({bulan_tahun})
=========================================================================
1. RINGKASAN OPERASIONAL:
   - Sepanjang bulan {bulan_tahun}, tercatat sebanyak {total_armada} unit armada sampah
     aktif beroperasi di bawah naungan DLH Kota Batam.
   - Total frekuensi perjalanan (ritase) pengangkutan menuju TPA adalah {total_ritase} trip.
   - Total volume sampah yang berhasil dipindahkan dan ditimbang
     mencapai {total_tonase_ton:,.0f} Ton.
   - Rata-rata durasi pelayanan di jembatan timbang: {durasi_rata:.1f} menit.
   - Wilayah dengan beban pengangkutan tertinggi berada di Kecamatan {kec_tertinggi}
     dengan kontribusi muatan sebesar {tonase_tertinggi:,.0f} Ton.
   - Armada teraktif: {teraktif.get('NOPIN','-')} ({teraktif.get('NO_PLAT','-')}) dengan {int(teraktif.get('Total_Trip',0))} trip.
   - Armada paling tidak efisien: {tidak_efisien.get('NOPIN','-')} ({tidak_efisien.get('NO_PLAT','-')}) dengan {int(tidak_efisien.get('Total_Trip',0))} trip.

2. REKOMENDASI STRATEGIS MANAJEMEN:
   - [Optimasi Rute & Armada] Melakukan redistribusi atau penambahan unit armada
     pada wilayah kritis (Kecamatan {kec_tertinggi}) guna mencegah kelambatan
     pengangkutan sampah di area pemukiman padat.
   - [Jadwal Shift Kerja] Mengatur ulang jam operasional keberangkatan armada
     untuk memecah penumpukan truk pada jam-jam puncak kepadatan di jembatan timbang.
   - [Pemeliharaan Rutin] Memberikan perhatian maintenance berkala pada armada
     dengan ritase terendah, antara lain: {list_terbawah}, untuk menganalisis
     apakah unit tersebut mengalami kendala mekanis atau kekurangan kru lapangan.
=========================================================================
    """
    return teks

def generate_pdf_report(data, grafik_dict, ringkasan_teks):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=18, textColor=colors.HexColor("#1e3c72"), spaceAfter=12)
    heading_style = ParagraphStyle('Heading', parent=styles['Heading2'], fontSize=14, textColor=colors.HexColor("#1e3c72"), spaceBefore=12, spaceAfter=6)
    normal_style = styles['Normal']

    story.append(Paragraph("Laporan Analisis Armada DLH Kota Batam", title_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph("Ringkasan Eksekutif", heading_style))
    for baris in ringkasan_teks.split('\n'):
        if baris.strip(): story.append(Paragraph(baris, normal_style))
    story.append(Spacer(1, 12))

    df_kec = data['df_kec']
    if not df_kec.empty:
        story.append(Paragraph("5 Kecamatan dengan Aktivitas Tertinggi", heading_style))
        table_data = [['Kecamatan', 'Ritase', 'Tonase (Kg)', 'Armada', 'Rata Durasi']]
        for _, row in df_kec.head(5).iterrows():
            table_data.append([row['Kecamatan'], str(row['Total_Ritase']), f"{row['Total_Tonase']:,.0f}", str(row['Jumlah_Armada']),
                               f"{row.get('Rata_Durasi_Menit', '-'):.1f}" if 'Rata_Durasi_Menit' in row else '-'])
        t = Table(table_data, colWidths=[120, 60, 80, 60, 80])
        t.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e3c72")),
                               ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                               ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                               ('FONTSIZE', (0,0), (-1,-1), 8)]))
        story.append(t); story.append(Spacer(1, 12))

    df_type = data['df_type']
    if not df_type.empty:
        story.append(Paragraph("Ringkasan per Jenis Armada", heading_style))
        table_data2 = [['Type', 'Ritase', 'Tonase (Kg)', 'Armada', 'Rata Durasi']]
        for _, row in df_type.iterrows():
            table_data2.append([row['TYPE'], str(row['Total_Ritase']), f"{row['Total_Tonase']:,.0f}", str(row['Jumlah_Armada']),
                                f"{row.get('Rata_Durasi_Menit', '-'):.1f}" if 'Rata_Durasi_Menit' in row else '-'])
        t2 = Table(table_data2, colWidths=[100, 60, 80, 60, 80])
        t2.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e3c72")),
                                ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                                ('ALIGN', (0,0), (-1,-1), 'CENTER'), ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                                ('FONTSIZE', (0,0), (-1,-1), 8)]))
        story.append(t2); story.append(Spacer(1, 12))

    temp_files = []
    story.append(Paragraph("Visualisasi Data", heading_style))
    for key, fig in grafik_dict.items():
        if fig is not None:
            fig.update_layout(template='plotly_white', paper_bgcolor='white', plot_bgcolor='white')
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
                pio.write_image(fig, tmp.name, format='png', width=500, height=300, scale=2)
                temp_files.append(tmp.name)
                story.append(Image(tmp.name, width=450, height=250))
                story.append(Spacer(1, 12))

    doc.build(story)
    for f in temp_files:
        try: os.unlink(f)
        except: pass
    buffer.seek(0)
    return buffer

# -------------------------- SESSION STATE --------------------------
if "hasil" not in st.session_state: st.session_state.hasil = None
if "sheets" not in st.session_state: st.session_state.sheets = None
if "config" not in st.session_state: st.session_state.config = None
if "grafik" not in st.session_state: st.session_state.grafik = {}
if "uploaded_file" not in st.session_state: st.session_state.uploaded_file = None

# -------------------------- ANTARMUKA STREAMLIT --------------------------
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
st.markdown("Unggah file Excel, pilih mode **Otomatis**, **Manual**, atau **Gunakan Sheet Master Data**. Menampilkan data waktu masuk/keluar dan tanggal.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        st.session_state.sheets = baca_semua_sheet(uploaded_file)
        st.session_state.uploaded_file = uploaded_file
        if not st.session_state.sheets: st.error("File tidak memiliki sheet yang valid.")
        else: st.success(f"Terbaca {len(st.session_state.sheets)} sheet.")

    if st.session_state.sheets:
        st.markdown("---")
        st.header("⚙️ Mode Pemrosesan")
        mode = st.radio("Pilih mode", ["Otomatis", "Manual", "Gunakan Sheet Master Data"])
        if mode == "Manual":
            with st.expander("Pengaturan Manual"):
                sheet_names = list(st.session_state.sheets.keys())
                armada_sheet = st.selectbox("Sheet List Armada", sheet_names)
                daily_candidates = [s for s in sheet_names if s.isdigit()] or sheet_names
                daily_sheets = st.multiselect("Sheet Harian", daily_candidates, default=daily_candidates)
                st.session_state.config = {'armada_sheet': armada_sheet, 'daily_sheets': daily_sheets}
        elif mode == "Gunakan Sheet Master Data":
            sheet_names = list(st.session_state.sheets.keys())
            master_sheet = st.selectbox("Pilih Sheet Master Data", sheet_names)
            st.session_state.config = {}
            use_master = True
        else:
            st.session_state.config = {'armada_sheet': None, 'daily_sheets': []}

        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Memproses..."):
                if mode == "Gunakan Sheet Master Data":
                    hasil = proses_data(st.session_state.sheets, {}, use_master=True, master_sheet=master_sheet)
                elif mode == "Otomatis":
                    armada = next((s for s in st.session_state.sheets if 'list armada' in s.lower()), None)
                    if armada is None: armada = next((s for s in st.session_state.sheets if 'armada' in s.lower()), None)
                    daily = [s for s in st.session_state.sheets if s.isdigit()]
                    if not daily: daily = [s for s in st.session_state.sheets if s != armada and s not in ['Tugas', 'Master Data']]
                    st.session_state.config = {'armada_sheet': armada, 'daily_sheets': daily}
                    hasil = proses_data(st.session_state.sheets, st.session_state.config)
                else:
                    hasil = proses_data(st.session_state.sheets, st.session_state.config)

                if hasil is None: st.error("Gagal memproses data.")
                else:
                    st.session_state.hasil = hasil
                    st.success(f"✅ {hasil['cleaned_count']} sheet berhasil diolah. Data waktu + tanggal siap.")
                    st.balloons()

# Tampilkan hasil
if st.session_state.hasil is not None:
    data = st.session_state.hasil
    df_master = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']
    df_type = data['df_type']
    df_armada = data['df_armada']
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']
    df_waktu_jenis = data['df_waktu_jenis']
    df_tren = data['df_tren']

    st.header("📊 Analisis Keseluruhan (Global)")
    total_trip_global = len(df_master)
    total_armada_global = df_master['NOPIN'].nunique()
    total_tonase_global = df_master[col_netto].sum() / 1000 if col_netto else 0
    durasi_rata_global = df_master['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_master.columns else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Trip", total_trip_global)
    col2.metric("Armada Aktif", total_armada_global)
    col3.metric("Total Tonase (Ton)", f"{total_tonase_global:,.1f}")
    col4.metric("Rata² Durasi (menit)", f"{durasi_rata_global:.1f}" if durasi_rata_global else "-")

    st.markdown("---")
    st.header("📋 Master Data (Tanggal, Masuk, Keluar, Durasi)")
    cols_waktu = ['NOPIN', 'NO_PLAT', 'Kecamatan', 'TANGGAL']
    if col_netto: cols_waktu.append(col_netto)
    if 'MASUK_ORI' in df_master.columns: cols_waktu.append('MASUK_ORI')
    if 'KELUAR_ORI' in df_master.columns: cols_waktu.append('KELUAR_ORI')
    if 'DURASI_MENIT' in df_master.columns: cols_waktu.append('DURASI_MENIT')
    cols_ada = [c for c in cols_waktu if c in df_master.columns]
    st.dataframe(df_master[cols_ada].sort_values(['TANGGAL', 'MASUK_ORI']).head(20), use_container_width=True)

    st.subheader("📊 Ringkasan Seluruh Kecamatan")
    if not df_kec.empty:
        col1, col2 = st.columns([2, 1])
        with col1: st.dataframe(df_kec.style.format({'Total_Tonase': '{:,.0f}', 'Rata_Durasi_Menit': '{:.1f}'}), use_container_width=True)
        with col2: st.plotly_chart(px.pie(df_kec, names='Kecamatan', values='Total_Ritase', title='Distribusi Trip per Kecamatan', template='plotly_white'), use_container_width=True)
        fig_ton = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', color='Total_Tonase', color_continuous_scale='Viridis', title='Total Tonase per Kecamatan', template='plotly_white')
        fig_ton.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_ton, use_container_width=True)

    st.subheader("🚛 Analisis per Jenis Armada (TYPE)")
    if not df_type.empty:
        col1, col2 = st.columns([2, 1])
        with col1: st.dataframe(df_type.style.format({'Total_Tonase': '{:,.0f}', 'Rata_Durasi_Menit': '{:.1f}'}), use_container_width=True)
        with col2: st.plotly_chart(px.pie(df_type, names='TYPE', values='Total_Ritase', title='Distribusi Trip per Type', template='plotly_white'), use_container_width=True)
        fig_type_bar = px.bar(df_type, x='TYPE', y='Total_Tonase', color='Total_Tonase', color_continuous_scale='Blues', title='Total Tonase per Type', template='plotly_white')
        fig_type_bar.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_type_bar, use_container_width=True)

    st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien")
    if teraktif is not None:
        col_a, col_b = st.columns(2)
        with col_a: st.success(f"**Teraktif:** {teraktif.get('NOPIN','-')} ({teraktif.get('NO_PLAT','')}) – {int(teraktif.get('Total_Trip',0))} trip")
        with col_b:
            if tidak_efisien is not None: st.error(f"**Tidak Efisien:** {tidak_efisien.get('NOPIN','-')} ({tidak_efisien.get('NO_PLAT','')}) – {int(tidak_efisien.get('Total_Trip',0))} trip")

    st.markdown("---")
    st.header("⏱️ Analisis Waktu Pelayanan (Durasi Masuk & Keluar)")
    if 'DURASI_MENIT' in df_master.columns and df_master['DURASI_MENIT'].notna().sum() > 0:
        valid_waktu = df_master['DURASI_MENIT'].notna().sum()
        col_w1, col_w2, col_w3, col_w4 = st.columns(4)
        col_w1.metric("Data Waktu Valid", valid_waktu)
        col_w2.metric("Rata‑rata Durasi", f"{df_master['DURASI_MENIT'].mean():.1f} menit")
        col_w3.metric("Durasi Minimum", f"{df_master['DURASI_MENIT'].min():.1f} menit")
        col_w4.metric("Durasi Maksimum", f"{df_master['DURASI_MENIT'].max():.1f} menit")

        fig_hist = px.histogram(df_master, x='DURASI_MENIT', nbins=30, title='Distribusi Durasi Pelayanan (menit)', template='plotly_white')
        st.plotly_chart(fig_hist, use_container_width=True)
        st.session_state.grafik['hist_durasi'] = fig_hist

        if not df_kec.empty and 'Rata_Durasi_Menit' in df_kec.columns:
            fig_durasi_kec = px.bar(df_kec.dropna(subset=['Rata_Durasi_Menit']), x='Kecamatan', y='Rata_Durasi_Menit',
                                    color='Rata_Durasi_Menit', color_continuous_scale='Blues',
                                    title='Rata‑rata Durasi per Kecamatan', template='plotly_white')
            st.plotly_chart(fig_durasi_kec, use_container_width=True)
            st.session_state.grafik['durasi_kec'] = fig_durasi_kec

        if not df_waktu_jenis.empty:
            fig_durasi_type = px.bar(df_waktu_jenis, x='Jenis Armada', y='Rata2 Waktu Tempuh (menit)',
                                     color='Rata2 Waktu Tempuh (menit)', color_continuous_scale='Teal',
                                     title='Rata‑rata Waktu Tempuh per Jenis Armada', template='plotly_white')
            st.plotly_chart(fig_durasi_type, use_container_width=True)
            st.session_state.grafik['durasi_type'] = fig_durasi_type

    if not df_tren.empty:
        st.markdown("---")
        st.subheader("📈 Tren Harian")
        fig_tren = px.line(df_tren, x='TANGGAL', y='Total_Ritase', title='Tren Frekuensi Ritase Harian', markers=True, template='plotly_white')
        fig_tren.update_traces(line_color='#0D9488')
        st.plotly_chart(fig_tren, use_container_width=True)

    st.session_state.grafik['tren'] = fig_tren
    st.session_state.grafik['kec_ton'] = fig_ton
    st.session_state.grafik['type_bar'] = fig_type_bar
    st.session_state.grafik['type_pie'] = px.pie(df_type, names='TYPE', values='Total_Ritase', title='Distribusi Trip per Type', template='plotly_white') if not df_type.empty else None

    st.markdown("---")
    st.header("📍 Analisis per Kecamatan")
    if 'Kecamatan' in df_master.columns:
        daftar_kec = sorted(df_master['Kecamatan'].unique().tolist())
        kec_terpilih = st.selectbox("Pilih Kecamatan", daftar_kec, key="kec_global")
        df_kec_filter = df_master[df_master['Kecamatan'] == kec_terpilih]

        total_trip_kec = len(df_kec_filter)
        total_armada_kec = df_kec_filter['NOPIN'].nunique()
        total_tonase_kec = df_kec_filter[col_netto].sum() / 1000 if col_netto else 0
        durasi_rata_kec = df_kec_filter['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_kec_filter.columns else 0

        colA, colB, colC, colD = st.columns(4)
        colA.metric("Total Trip", total_trip_kec)
        colB.metric("Armada Aktif", total_armada_kec)
        colC.metric("Total Tonase (Ton)", f"{total_tonase_kec:,.1f}")
        colD.metric("Rata² Durasi (menit)", f"{durasi_rata_kec:.1f}" if durasi_rata_kec else "-")

        st.subheader(f"🚚 Daftar Armada di {kec_terpilih}")
        if col_netto:
            armada_kec = df_kec_filter.groupby(['NOPIN', 'NO_PLAT'], dropna=False).agg(
                Total_Trip=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
            ).reset_index().sort_values('Total_Trip', ascending=False)
            st.dataframe(armada_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True)
            st.plotly_chart(px.bar(armada_kec.head(10), x='NOPIN', y='Total_Trip', color='Total_Trip', color_continuous_scale='Blues',
                                   title=f'10 Armada Teraktif di {kec_terpilih}', template='plotly_white'), use_container_width=True)

    st.markdown("---")
    st.subheader("📝 Laporan Ringkasan Eksekutif (Poin 7.0)")
    ringkasan_teks = buat_ringkasan_eksekutif(data)
    st.markdown(f"```\n{ringkasan_teks}\n```")
    st.download_button("📄 Unduh Ringkasan Eksekutif (TXT)", ringkasan_teks.encode('utf-8'), "Ringkasan_Eksekutif_Poin7.txt")

    st.subheader("📑 Laporan PDF Lengkap (Warna Sesuai Dashboard)")
    if st.button("📥 Buat Laporan PDF"):
        with st.spinner("Membuat PDF..."):
            pdf_buffer = generate_pdf_report(data, st.session_state.grafik, ringkasan_teks)
            st.download_button(label="⬇️ Unduh Laporan PDF", data=pdf_buffer, file_name="Laporan_DLH_Armada.pdf", mime="application/pdf")

    st.subheader("📥 Unduh Data Hasil Analisis")
    @st.cache_data
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as w: dataframe.to_excel(w, index=False)
        return output.getvalue()

    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        st.download_button("📊 Master Data (Excel)", to_excel(df_master), "Master_Data.xlsx")
        st.download_button("📊 Statistik Armada (Excel)", to_excel(df_armada), "Statistik_Armada.xlsx")
    with col_u2:
        st.download_button("📊 Laporan per Kecamatan (Excel)", to_excel(df_kec), "Kecamatan.xlsx")
        st.download_button("📊 Laporan per Type (Excel)", to_excel(df_type), "Type_Armada.xlsx")
        st.download_button("📈 Tren Harian (Excel)", to_excel(df_tren), "Tren_Harian.xlsx")
    with col_u3:
        if not df_waktu_jenis.empty:
            st.download_button("⏱️ Waktu per Jenis (Excel)", to_excel(df_waktu_jenis), "Waktu_per_Jenis.xlsx")
        if 'DURASI_MENIT' in df_master.columns:
            cols_csv = ['NOPIN','NO_PLAT','TANGGAL','MASUK_ORI','KELUAR_ORI','DURASI_MENIT']
            cols_csv_ada = [c for c in cols_csv if c in df_master.columns]
            st.download_button("⏱️ Data Waktu & Durasi (CSV)", df_master[cols_csv_ada].to_csv(index=False).encode('utf-8'), "waktu_durasi.csv")

    with st.expander("🔎 Lihat Data Mentah Lengkap (200 baris pertama)"):
        st.dataframe(df_master.head(200))
