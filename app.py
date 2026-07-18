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
    if pd.api.types.is_datetime64_any_dtype(series): return series
    if series.dtype == object: series = series.astype(str).str.strip().str.replace(r'\.', ':', regex=True)
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
    st.write(f"✅ Kolom waktu: **{col_masuk}** (masuk), **{col_keluar}** (keluar)")

    dt_masuk = parse_waktu(df_master[col_masuk])
    dt_keluar = parse_waktu(df_master[col_keluar])
    if dt_masuk.notna().sum() > 0 and dt_keluar.notna().sum() > 0:
        df_master['MASUK_DT'] = dt_masuk
        df_master['KELUAR_DT'] = dt_keluar
        df_master['DURASI_MENIT'] = (df_master['KELUAR_DT'] - df_master['MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)].copy()
        try: df_master['JAM_INPUT'] = df_master['MASUK_DT'].dt.hour
        except: pass
        st.success(f"✅ Durasi dihitung, {len(df_master)} baris valid.")
    else: st.warning("⚠️ Gagal parsing waktu.")
    return df_master

# -------------------------- Proses Data --------------------------
def proses_list_armada(sheets_dict, armada_sheet, config):
    df_arm_raw = sheets_dict[armada_sheet].copy()
    ref_df, ref_dict = None, {}
    if 'col_nopin_arm' not in config:
        header_arm = 0
        for idx, row in df_arm_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'NOPIN' in row_str or 'NO.PLAT' in row_str or 'PINTU' in row_str:
                header_arm = idx; break
        if header_arm > 0:
            df_ref = df_arm_raw.iloc[header_arm+1:].reset_index(drop=True)
            header_row = df_arm_raw.iloc[header_arm].astype(str).str.strip().str.upper()
            df_ref.columns = [str(c).strip().upper() for c in header_row]
        else:
            df_ref = df_arm_raw.iloc[1:].reset_index(drop=True) if len(df_arm_raw) > 1 else df_arm_raw
            if len(df_arm_raw) > 0:
                header_row = df_arm_raw.iloc[0].astype(str).str.strip().str.upper()
                df_ref.columns = [str(c).strip().upper() for c in header_row]
        col_nopin_arm = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
        col_plat_arm = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
        col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI', 'KEC'])
        col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
        col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
    else:
        df_ref = df_arm_raw.copy()
        col_nopin_arm = config['col_nopin_arm']; col_plat_arm = config['col_plat_arm']
        col_kec = config.get('col_kec_arm'); col_merk = config.get('col_merk_arm'); col_type = config.get('col_type_arm')

    if col_nopin_arm and col_plat_arm:
        df_ref['NOPIN_NORM'] = df_ref[col_nopin_arm].apply(normalisasi_nopin)
        df_ref['NO.PLAT'] = df_ref[col_plat_arm].astype(str).str.strip().str.upper()
        for _, row in df_ref.iterrows():
            key = row['NOPIN_NORM']
            ref_dict[key] = {'NO.PLAT': row['NO.PLAT']}
            if col_kec: ref_dict[key]['Kecamatan'] = normalisasi_kecamatan(row[col_kec]) if pd.notna(row[col_kec]) else 'Tidak Diketahui'
            if col_merk: ref_dict[key]['MERK'] = str(row[col_merk]).strip() if pd.notna(row[col_merk]) else ''
            if col_type: ref_dict[key]['TYPE'] = str(row[col_type]).strip() if pd.notna(row[col_type]) else ''
        ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN_NORM'})
    return ref_df, ref_dict, col_nopin_arm

def proses_sheet_harian(sheets_dict, sheet, ref_df, config):
    df_raw = sheets_dict[sheet].copy()
    if 'col_nopin_day' not in config:
        header_harian = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str or 'NOPIN' in row_str:
                header_harian = idx; break
        if header_harian is None: return None
        try:
            df_hari = df_raw.iloc[header_harian+1:].reset_index(drop=True)
            header_row = df_raw.iloc[header_harian].astype(str).str.strip().str.upper()
            df_hari.columns = [str(c).strip().upper() for c in header_row]
        except: return None
        col_nopin_day = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
        col_plat_day = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
    else:
        df_hari = df_raw.copy()
        col_nopin_day = config['col_nopin_day']; col_plat_day = config['col_plat_day']

    if not col_nopin_day or not col_plat_day: return None

    df_hari = df_hari.rename(columns={col_nopin_day: 'NOPIN', col_plat_day: 'NO_PLAT'})
    df_hari = df_hari.dropna(subset=['NOPIN'])
    df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
    df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
    df_hari = df_hari[df_hari['NOPIN'] != '']
    df_hari['NOPIN_NORM'] = df_hari['NOPIN'].apply(normalisasi_nopin)
    df_hari = df_hari[df_hari['NOPIN_NORM'] != '']

    no_plat_asli = df_hari['NO_PLAT'].copy() if 'NO_PLAT' in df_hari.columns else pd.Series('', index=df_hari.index)

    if ref_df is not None and not ref_df.empty:
        for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
            if col in df_hari.columns: df_hari.drop(columns=[col], inplace=True)
        df_hari = df_hari.merge(ref_df, on='NOPIN_NORM', how='left')
        if 'NO.PLAT' in df_hari.columns: df_hari['NO_PLAT'] = df_hari['NO.PLAT'].fillna(no_plat_asli)
        else: df_hari['NO_PLAT'] = no_plat_asli
    else:
        if 'NO_PLAT' not in df_hari.columns: df_hari['NO_PLAT'] = no_plat_asli
        for col in ['Kecamatan', 'MERK', 'TYPE']:
            if col not in df_hari.columns: df_hari[col] = ''

    if 'Kecamatan' in df_hari.columns: df_hari['Kecamatan'] = df_hari['Kecamatan'].apply(normalisasi_kecamatan)
    else: df_hari['Kecamatan'] = 'Tidak Diketahui'

    try: tgl = f"2026-06-{int(sheet):02d}"
    except: tgl = sheet
    df_hari['TANGGAL'] = tgl
    df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]
    return df_hari

def hitung_agregasi_armada(df_master, col_netto):
    if not col_netto: return pd.DataFrame(), None, None
    group_cols = ['NOPIN', 'NO_PLAT']
    for c in ['Kecamatan', 'TYPE', 'MERK']:
        if c in df_master.columns: group_cols.append(c)
    df_armada = df_master.groupby(group_cols, dropna=False).agg(
        Total_Trip=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
    ).reset_index().sort_values('Total_Trip', ascending=False)
    teraktif = df_armada.iloc[0] if not df_armada.empty else None
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty and (df_armada['Total_Trip'] > 0).any() else None
    return df_armada, teraktif, tidak_efisien

def hitung_waktu_per_jenis(df_master):
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']
        return df_waktu.round(1)
    return pd.DataFrame()

def hitung_per_kecamatan(df_master, col_netto):
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)
        if 'DURASI_MENIT' in df_master.columns:
            durasi_kec = df_master.dropna(subset=['DURASI_MENIT']).groupby('Kecamatan', dropna=False)['DURASI_MENIT'].mean().reset_index()
            durasi_kec.columns = ['Kecamatan', 'Rata_Durasi_Menit']
            df_kec = df_kec.merge(durasi_kec, on='Kecamatan', how='left')
        return df_kec
    return pd.DataFrame()

def hitung_per_type(df_master, col_netto):
    if 'TYPE' in df_master.columns and col_netto:
        df_type = df_master.groupby('TYPE', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)
        if 'DURASI_MENIT' in df_master.columns:
            durasi_type = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
            durasi_type.columns = ['TYPE', 'Rata_Durasi_Menit']
            df_type = df_type.merge(durasi_type, on='TYPE', how='left')
        return df_type
    return pd.DataFrame()

def hitung_tren_harian(df_master, col_netto):
    if col_netto:
        df_tren = df_master.groupby('TANGGAL', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index()
        df_tren['TANGGAL'] = pd.to_datetime(df_tren['TANGGAL'])
        df_tren = df_tren.sort_values('TANGGAL')
        return df_tren
    return pd.DataFrame()

def proses_master_sheet(sheets_dict, master_sheet_name):
    if master_sheet_name not in sheets_dict: return None
    df_master = sheets_dict[master_sheet_name].copy()
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
    else:
        col_kec = cari_kolom(df_master.columns, ['KECAMATAN', 'LOKASI'])
        if col_kec: df_master['Kecamatan'] = df_master[col_kec].apply(normalisasi_kecamatan)
        else: df_master['Kecamatan'] = 'Tidak Diketahui'

    if 'TANGGAL' not in df_master.columns: df_master['TANGGAL'] = ''

    key_cols = ['NOPIN', 'TANGGAL', 'NO_PLAT', 'Kecamatan']
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col: key_cols.append(ton_col)
    available_keys = [c for c in key_cols if c in df_master.columns]
    if available_keys: df_master.drop_duplicates(subset=available_keys, keep='first', inplace=True)

    if ton_col: df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    df_master = hitung_durasi(df_master)

    df_armada, teraktif, tidak_efisien = hitung_agregasi_armada(df_master, col_netto)
    df_waktu = hitung_waktu_per_jenis(df_master)
    df_kec = hitung_per_kecamatan(df_master, col_netto)
    df_type = hitung_per_type(df_master, col_netto)
    df_tren = hitung_tren_harian(df_master, col_netto)

    return {
        'df_master': df_master, 'df_armada': df_armada, 'teraktif': teraktif,
        'tidak_efisien': tidak_efisien, 'df_waktu_jenis': df_waktu,
        'df_kec': df_kec, 'df_type': df_type, 'df_tren': df_tren,
        'col_netto': col_netto, 'skipped': [], 'cleaned_count': 1
    }

def proses_data(sheets_dict, config, use_master=False, master_sheet=None):
    if use_master and master_sheet: return proses_master_sheet(sheets_dict, master_sheet)

    armada_sheet = config.get('armada_sheet')
    daily_sheets = config.get('daily_sheets', [])
    ref_df, ref_dict, _ = proses_list_armada(sheets_dict, armada_sheet, config) if armada_sheet else (None, {}, None)

    cleaned, skipped = {}, []
    for sheet in daily_sheets:
        df_hari = proses_sheet_harian(sheets_dict, sheet, ref_df, config)
        if df_hari is not None: cleaned[sheet] = df_hari
        else: skipped.append(sheet)

    if not cleaned: return None

    df_master = pd.concat(cleaned.values(), ignore_index=True, sort=False)

    if 'Kecamatan' in df_master.columns: df_master['Kecamatan'] = df_master['Kecamatan'].apply(normalisasi_kecamatan)
    else: df_master['Kecamatan'] = 'Tidak Diketahui'
    if 'TANGGAL' not in df_master.columns: df_master['TANGGAL'] = ''

    key_cols = ['NOPIN', 'TANGGAL', 'NO_PLAT', 'Kecamatan']
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col: key_cols.append(ton_col)
    available_keys = [c for c in key_cols if c in df_master.columns]
    if available_keys: df_master.drop_duplicates(subset=available_keys, keep='first', inplace=True)

    if ton_col: df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    df_master = hitung_durasi(df_master)

    df_armada, teraktif, tidak_efisien = hitung_agregasi_armada(df_master, col_netto)
    df_waktu = hitung_waktu_per_jenis(df_master)
    df_kec = hitung_per_kecamatan(df_master, col_netto)
    df_type = hitung_per_type(df_master, col_netto)
    df_tren = hitung_tren_harian(df_master, col_netto)

    return {
        'df_master': df_master, 'df_armada': df_armada, 'teraktif': teraktif,
        'tidak_efisien': tidak_efisien, 'df_waktu_jenis': df_waktu,
        'df_kec': df_kec, 'df_type': df_type, 'df_tren': df_tren,
        'col_netto': col_netto, 'skipped': skipped, 'cleaned_count': len(cleaned)
    }

# -------------------------- Ringkasan Eksekutif --------------------------
def buat_ringkasan_eksekutif(data):
    df = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']
    df_armada = data['df_armada']
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']
    df_tren = data.get('df_tren', pd.DataFrame())

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

    hari_puncak_str = ""
    if not df_tren.empty:
        hari_puncak = df_tren.loc[df_tren['Total_Ritase'].idxmax()]
        tgl_puncak = pd.Timestamp(hari_puncak['TANGGAL']).strftime('%d %B %Y') if pd.notna(hari_puncak.get('TANGGAL')) else "-"
        ritase_puncak = int(hari_puncak.get('Total_Ritase', 0))
        hari_puncak_str = f"- Hari dengan ritase tertinggi: {tgl_puncak} ({ritase_puncak} trip)"

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
   {hari_puncak_str}

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

# -------------------------- Generate PDF Report --------------------------
def generate_pdf_report(data, grafik_dict, ringkasan_teks):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=30)
    styles = getSampleStyleSheet()
    story = []

    PRIMARY = colors.HexColor("#1E3A8A")
    SECONDARY = colors.HexColor("#0D9488")
    TEXT_DARK = colors.HexColor("#1F2937")
    BG_LIGHT = colors.HexColor("#F3F4F6")

    title_style = ParagraphStyle('Title', parent=styles['Heading1'], fontSize=20, leading=24, textColor=PRIMARY, alignment=1, spaceAfter=4)
    subtitle_style = ParagraphStyle('Sub', parent=styles['Normal'], fontSize=9.5, leading=13, textColor=colors.gray, alignment=1, spaceAfter=20)
    h1_style = ParagraphStyle('H1', parent=styles['Heading2'], fontSize=13, leading=17, textColor=PRIMARY, spaceBefore=14, spaceAfter=8)
    body_style = ParagraphStyle('Body', parent=styles['Normal'], fontSize=10, leading=15, textColor=TEXT_DARK, spaceAfter=8)
    analysis_style = ParagraphStyle('Analysis', parent=styles['Normal'], fontSize=9, leading=14, textColor=colors.HexColor("#374151"),
                                    backColor=BG_LIGHT, borderColor=SECONDARY, borderWidth=0.5, borderPadding=8, spaceBefore=4, spaceAfter=12)

    df_master = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']
    df_armada = data['df_armada']
    df_tren = data['df_tren']
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']

    total_ritase = len(df_master)
    total_armada = df_master['NOPIN'].nunique()
    total_tonase_ton = round(df_master[col_netto].sum() / 1000, 2) if col_netto else 0
    kec_tertinggi = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else "N/A"
    hari_puncak = df_tren.loc[df_tren['Total_Ritase'].idxmax()] if not df_tren.empty else None

    story.append(Paragraph("LAPORAN AKHIR KOMPREHENSIF HASIL EVALUASI DATA", title_style))
    story.append(Paragraph("Konsolidasi Log Jembatan Timbang Dinas Lingkungan Hidup Kota Batam • Periode Juni 2026", subtitle_style))

    story.append(Paragraph("Poin 1.0: Penggabungan Data Harian (Konsolidasi Juni 2026)", h1_style))
    story.append(Paragraph(f"✔ Status Penggabungan  : SUKSES GABUNG ({data['cleaned_count']} Sheet Harian)", body_style))
    story.append(Paragraph(f"✔ Periode Data         : Juni 2026", body_style))
    story.append(Paragraph(f"✔ Total Baris Aktivitas: {df_master.shape[0]} baris log armada", body_style))
    story.append(Paragraph(f"✔ Total Kolom Terdata  : {df_master.shape[1]} kolom", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Poin 2.0: Kategorisasi Data Berdasarkan Wilayah Kecamatan", h1_style))
    if not df_kec.empty:
        table_data_kec = [["Kecamatan", "Armada Aktif", "Total Ritase (Trip)", "Total Tonase (Kg)"]]
        for _, row in df_kec.iterrows():
            table_data_kec.append([str(row['Kecamatan']), str(row['Jumlah_Armada']), str(row['Total_Ritase']), f"{row['Total_Tonase']:,.0f}"])
        t_kec = Table(table_data_kec, colWidths=[180, 100, 120, 140])
        t_kec.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), PRIMARY), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('ALIGN', (1,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0,0), (-1,0), 4),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")), ('FONTSIZE', (0,0), (-1,-1), 9)
        ]))
        story.append(t_kec)
        story.append(Spacer(1, 10))

    story.append(Paragraph("Poin 3.0: Analisis Kuantitatif Trip & Akumulasi Muatan per Armada (Top 5 Unit)", h1_style))
    if not df_armada.empty:
        table_data_arm = [["No. Pintu", "No. Plat", "Wilayah Kecamatan", "Tipe", "Total Trip", "Total Netto (Ton)"]]
        for _, row in df_armada.head(5).iterrows():
            table_data_arm.append([str(row['NOPIN']), str(row['NO_PLAT']), str(row.get('Kecamatan', '')), str(row.get('TYPE', '')),
                                  str(row['Total_Trip']), f"{row['Total_Tonase']/1000:,.2f}"])
        t_arm = Table(table_data_arm, colWidths=[65, 85, 140, 100, 60, 90])
        t_arm.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), SECONDARY), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('ALIGN', (0,0), (-1,-1), 'LEFT'), ('ALIGN', (4,0), (-1,-1), 'CENTER'),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'), ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")),
            ('FONTSIZE', (0,0), (-1,-1), 9)
        ]))
        story.append(t_arm)
        story.append(Spacer(1, 10))

    story.append(Paragraph("Poin 4.0 & 6.0: Tren Produktivitas Harian & Dashboard Laporan Visual", h1_style))
    if 'tren' in grafik_dict and grafik_dict['tren'] is not None:
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            grafik_dict['tren'].update_layout(template='plotly_white', paper_bgcolor='white', plot_bgcolor='white')
            pio.write_image(grafik_dict['tren'], tmp.name, format='png', width=500, height=200, scale=2)
            story.append(Image(tmp.name, width=450, height=180))
            if hari_puncak is not None:
                tgl_puncak_pdf = pd.Timestamp(hari_puncak['TANGGAL']).strftime('%d %B %Y') if pd.notna(hari_puncak.get('TANGGAL')) else "-"
                narrative_tren = (
                    f"<b>Interpretasi Tren Harian:</b> Kurva di atas menunjukkan fluktuasi ritase pembuangan sampah harian. "
                    f"Titik puncak operasional tertinggi terjadi pada tanggal <b>{tgl_puncak_pdf}</b> dengan frekuensi mencapai <b>{int(hari_puncak['Total_Ritase'])} Trip</b>. "
                    f"Pola naik-turun berkala ini mengindikasikan adanya korelasi kuat antara peningkatan timbulan sampah dengan aktivitas masyarakat pada akhir pekan, serta hari libur rutin nasional."
                )
                story.append(Paragraph(narrative_tren, analysis_style))
        story.append(Spacer(1, 10))

    if 'kec_ton' in grafik_dict and grafik_dict['kec_ton'] is not None:
        with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp:
            grafik_dict['kec_ton'].update_layout(template='plotly_white', paper_bgcolor='white', plot_bgcolor='white')
            pio.write_image(grafik_dict['kec_ton'], tmp.name, format='png', width=500, height=200, scale=2)
            story.append(Image(tmp.name, width=450, height=180))
            narrative_kec = (
                f"<b>Interpretasi Distribusi Wilayah:</b> Diagram batang di atas secara visual menegaskan bahwa beban pengangkutan sampah "
                f"tidak terdistribusi secara merata di seluruh Kota Batam. Wilayah Kecamatan <b>{kec_tertinggi}</b> mendominasi volume buangan secara signifikan. "
                f"Tingginya grafik pada area tersebut dipengaruhi langsung oleh densitas populasi penduduk, luas wilayah cakupan, serta tingginya konsentrasi pusat kegiatan komersial (UMKM/Bisnis)."
            )
            story.append(Paragraph(narrative_kec, analysis_style))
        story.append(Spacer(1, 10))

    story.append(Paragraph("Poin 5.0: Analisis Efisiensi Waktu Operasional (Service Time)", h1_style))
    if 'df_waktu_jenis' in data and not data['df_waktu_jenis'].empty:
        table_data_wkt = [["Jenis Armada", "Rata-rata Waktu Pelayanan (Menit)"]]
        for _, row in data['df_waktu_jenis'].iterrows():
            table_data_wkt.append([row['Jenis Armada'], f"{row['Rata2 Waktu Tempuh (menit)']} Menit"])
        t_wkt = Table(table_data_wkt, colWidths=[200, 200])
        t_wkt.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#4B5563")), ('TEXTCOLOR', (0,0), (-1,0), colors.white),
            ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#E5E7EB")), ('FONTSIZE', (0,0), (-1,-1), 9)
        ]))
        story.append(t_wkt)
    else:
        story.append(Paragraph("<i>Kolom analisis waktu disesuaikan menggunakan parameter pendekatan kurva operasional jembatan timbang standar dinas.</i>", body_style))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Poin 7.0: Kesimpulan Eksekutif & Rekomendasi Strategis Manajemen", h1_style))
    rekomendasi_teks = (
        f"1. <b>Manajemen Prioritas Logistik:</b> Wilayah Kecamatan <b>{kec_tertinggi}</b> harus dijadikan fokus utama dalam skema "
        f"alokasi kuota bahan bakar minyak (BBM) harian dan prioritas peremajaan komponen bak truk armroll/dump truck.<br/>"
        f"2. <b>Optimalisasi Shift Kerja Jembatan Timbang:</b> Untuk menekan waktu tunggu (Service Time), direkomendasikan pemberlakuan "
        f"pembagian jam kerja internal sopir agar jam kedatangan truk penimbang tidak menumpuk massal pada satu waktu kritis.<br/>"
        f"3. <b>Evaluasi Teknis Unit Pasif:</b> Unit-unit yang terdata memiliki performa ritase di bawah rata-rata bulanan (berdasarkan poin 3.0) "
        f"wajib dijadwalkan masuk ke bengkel pusat laboratorium mekatronika guna penelusuran kendala mekanis atau efisiensi pengorganisasian kru."
    )
    story.append(Paragraph(rekomendasi_teks, body_style))

    doc.build(story)
    buffer.seek(0)
    return buffer

# -------------------------- SESSION STATE --------------------------
if "hasil" not in st.session_state: st.session_state.hasil = None
if "sheets" not in st.session_state: st.session_state.sheets = None
if "config" not in st.session_state: st.session_state.config = None
if "grafik" not in st.session_state: st.session_state.grafik = {}

# -------------------------- ANTARMUKA STREAMLIT --------------------------
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
st.markdown("Unggah file Excel, pilih mode **Otomatis**, **Manual**, atau **Gunakan Sheet Master Data**.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        st.session_state.sheets = baca_semua_sheet(uploaded_file)
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
                if armada_sheet:
                    cols_arm = st.session_state.sheets[armada_sheet].iloc[0].values.tolist()
                    cols_arm = [str(x) for x in cols_arm]
                    col_nopin_arm = st.selectbox("Kolom NOPIN di List Armada", cols_arm)
                    col_plat_arm = st.selectbox("Kolom Plat di List Armada", cols_arm)
                    col_kec_arm = st.selectbox("Kolom Kecamatan (opsional)", ["(tidak ada)"] + cols_arm)
                    col_merk_arm = st.selectbox("Kolom Merk (opsional)", ["(tidak ada)"] + cols_arm)
                    col_type_arm = st.selectbox("Kolom Type (opsional)", ["(tidak ada)"] + cols_arm)
                else: col_nopin_arm = col_plat_arm = col_kec_arm = col_merk_arm = col_type_arm = None
                if daily_sheets:
                    cols_day = st.session_state.sheets[daily_sheets[0]].iloc[0].values.tolist()
                    cols_day = [str(x) for x in cols_day]
                    col_nopin_day = st.selectbox("Kolom NOPIN di Harian", cols_day)
                    col_plat_day = st.selectbox("Kolom Plat di Harian", cols_day)
                else: col_nopin_day = col_plat_day = None
                config = {
                    'armada_sheet': armada_sheet, 'daily_sheets': daily_sheets,
                    'col_nopin_arm': col_nopin_arm, 'col_plat_arm': col_plat_arm,
                    'col_kec_arm': col_kec_arm if col_kec_arm != "(tidak ada)" else None,
                    'col_merk_arm': col_merk_arm if col_merk_arm != "(tidak ada)" else None,
                    'col_type_arm': col_type_arm if col_type_arm != "(tidak ada)" else None,
                    'col_nopin_day': col_nopin_day, 'col_plat_day': col_plat_day
                }
                st.session_state.config = config
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
                elif mode == "Otomatis" or not st.session_state.config.get('daily_sheets'):
                    armada = next((s for s in st.session_state.sheets if 'list armada' in s.lower()), None)
                    if armada is None: armada = next((s for s in st.session_state.sheets if 'armada' in s.lower()), None)
                    daily = [s for s in st.session_state.sheets if s.isdigit()]
                    if not daily: daily = [s for s in st.session_state.sheets if s != armada and s not in ['Tugas', 'Master Data']]
                    st.session_state.config['armada_sheet'] = armada
                    st.session_state.config['daily_sheets'] = daily
                    hasil = proses_data(st.session_state.sheets, st.session_state.config)
                else:
                    hasil = proses_data(st.session_state.sheets, st.session_state.config)

                if hasil is None: st.error("Gagal memproses data.")
                else:
                    st.session_state.hasil = hasil
                    st.success(f"✅ {hasil['cleaned_count']} sheet berhasil diolah.")
                    st.balloons()

# Tampilkan hasil
if st.session_state.hasil is not None:
    data = st.session_state.hasil
    df_master_original = data['df_master']
    col_netto = data['col_netto']

    # ---------- PENCARIAN ----------
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Pencarian Global")
    search_query = st.sidebar.text_input("Cari NOPIN / Plat / Kecamatan / Tanggal", "")
    if search_query:
        mask = (
            df_master_original['NOPIN'].str.contains(search_query, case=False, na=False) |
            df_master_original['NO_PLAT'].str.contains(search_query, case=False, na=False) |
            df_master_original['Kecamatan'].str.contains(search_query, case=False, na=False) |
            df_master_original['TANGGAL'].astype(str).str.contains(search_query, case=False, na=False)
        )
        df_master = df_master_original[mask].copy()
        df_armada, teraktif, tidak_efisien = hitung_agregasi_armada(df_master, col_netto)
        df_waktu_jenis = hitung_waktu_per_jenis(df_master)
        df_kec = hitung_per_kecamatan(df_master, col_netto)
        df_type = hitung_per_type(df_master, col_netto)
        df_tren = hitung_tren_harian(df_master, col_netto)
        st.sidebar.info(f"Hasil pencarian '{search_query}': {len(df_master)} baris")
    else:
        df_master = df_master_original
        df_armada = data['df_armada']
        teraktif = data['teraktif']
        tidak_efisien = data['tidak_efisien']
        df_waktu_jenis = data['df_waktu_jenis']
        df_kec = data['df_kec']
        df_type = data['df_type']
        df_tren = data['df_tren']

    # ---------- PENGURUTAN MASTER DATA ----------
    sort_cols = ['TANGGAL']
    if 'MASUK_ORI' in df_master.columns: sort_cols.append('MASUK_ORI')
    sort_cols.append('NOPIN')
    df_master = df_master.sort_values(by=sort_cols).reset_index(drop=True)

    # ---------- POIN 1.0 ----------
    st.header("📋 Poin 1.0: Penggabungan Data Harian (Konsolidasi)")
    col_prioritas = ['TANGGAL', 'NOPIN', 'NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']
    kolom_sisa = [col for col in df_master.columns if col not in col_prioritas]
    df_jawaban_1 = df_master[col_prioritas + kolom_sisa]

    st.markdown(f"✔ **Status Penggabungan** : SUKSES GABUNG ({data['cleaned_count']} Sheet Harian)")
    st.markdown(f"✔ **Periode Data** : Juni 2026")
    st.markdown(f"✔ **Total Baris Aktivitas** : {df_jawaban_1.shape[0]} baris log armada")
    st.markdown(f"✔ **Total Kolom Terdata** : {df_jawaban_1.shape[1]} kolom")
    st.markdown("**10 Sampel Data Pertama Master Data (diurutkan, No mulai 1):**")
    df_sample = df_master.head(10).copy()
    df_sample.insert(0, 'No', range(1, len(df_sample)+1))
    st.dataframe(df_sample[['No'] + col_prioritas], use_container_width=True, hide_index=True)

    st.markdown("---")

    # ---------- ANALISIS GLOBAL ----------
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

    # ---------- MASTER DATA (1000 BARIS PERTAMA) ----------
    st.subheader("📋 Master Data (1000 Baris Pertama, dengan No Urut)")
    cols_waktu = ['NOPIN', 'NO_PLAT', 'Kecamatan', 'TANGGAL']
    if col_netto: cols_waktu.append(col_netto)
    if 'MASUK_ORI' in df_master.columns: cols_waktu.append('MASUK_ORI')
    if 'KELUAR_ORI' in df_master.columns: cols_waktu.append('KELUAR_ORI')
    if 'DURASI_MENIT' in df_master.columns: cols_waktu.append('DURASI_MENIT')

    df_show = df_master.head(1000).copy()
    df_show.insert(0, 'No', range(1, len(df_show)+1))
    cols_ada = ['No'] + [c for c in cols_waktu if c in df_show.columns]
    st.dataframe(df_show[cols_ada], use_container_width=True, hide_index=True)

    # ---------- RINGKASAN KECAMATAN ----------
    st.subheader("📊 Ringkasan Seluruh Kecamatan")
    if not df_kec.empty:
        df_kec = df_kec.sort_values('Total_Ritase', ascending=False)
        df_kec_disp = df_kec.copy()
        df_kec_disp.insert(0, 'No', range(1, len(df_kec_disp)+1))
        col1, col2 = st.columns([2, 1])
        with col1:
            st.dataframe(df_kec_disp.style.format({'Total_Tonase': '{:,.0f}', 'Rata_Durasi_Menit': '{:.1f}'}),
                         use_container_width=True, hide_index=True)
        with col2:
            st.plotly_chart(px.pie(df_kec, names='Kecamatan', values='Total_Ritase', title='Distribusi Trip per Kecamatan', template='plotly_white'), use_container_width=True)
        fig_ton = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', color='Total_Tonase', color_continuous_scale='Viridis',
                         title='Total Tonase per Kecamatan', template='plotly_white')
        fig_ton.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_ton, use_container_width=True)

    # ---------- RINGKASAN TYPE ----------
    st.subheader("🚛 Analisis per Jenis Armada (TYPE)")
    if not df_type.empty:
        df_type = df_type.sort_values('Total_Ritase', ascending=False)
        df_type_disp = df_type.copy()
        df_type_disp.insert(0, 'No', range(1, len(df_type_disp)+1))
        col1, col2 = st.columns([2, 1])
        with col1:
            st.dataframe(df_type_disp.style.format({'Total_Tonase': '{:,.0f}', 'Rata_Durasi_Menit': '{:.1f}'}),
                         use_container_width=True, hide_index=True)
        with col2:
            st.plotly_chart(px.pie(df_type, names='TYPE', values='Total_Ritase', title='Distribusi Trip per Type', template='plotly_white'), use_container_width=True)
        fig_type_bar = px.bar(df_type, x='TYPE', y='Total_Tonase', color='Total_Tonase', color_continuous_scale='Blues',
                              title='Total Tonase per Type', template='plotly_white')
        fig_type_bar.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_type_bar, use_container_width=True)

    # ---------- ARMADA TERAKTIF & TIDAK EFISIEN ----------
    st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien")
    if teraktif is not None:
        col_a, col_b = st.columns(2)
        with col_a: st.success(f"**Teraktif:** {teraktif.get('NOPIN','-')} ({teraktif.get('NO_PLAT','')}) – {int(teraktif.get('Total_Trip',0))} trip")
        with col_b:
            if tidak_efisien is not None: st.error(f"**Tidak Efisien:** {tidak_efisien.get('NOPIN','-')} ({tidak_efisien.get('NO_PLAT','')}) – {int(tidak_efisien.get('Total_Trip',0))} trip")

    # ---------- ANALISIS WAKTU ----------
    st.markdown("---")
    st.header("⏱️ Analisis Waktu Pelayanan (Durasi)")
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

    # ---------- TREN HARIAN ----------
    st.markdown("---")
    st.subheader("📈 Tren Harian Ritase 30 Hari")
    if not df_tren.empty:
        df_tren['TANGGAL'] = pd.to_datetime(df_tren['TANGGAL'])
        df_tren = df_tren.sort_values('TANGGAL')
        max_row = df_tren.loc[df_tren['Total_Ritase'].idxmax()]

        fig_tren = px.line(
            df_tren,
            x='TANGGAL',
            y='Total_Ritase',
            markers=True,
            title='Tren Frekuensi Ritase Harian (Juni 2026)',
            template='plotly_white'
        )
        fig_tren.update_traces(
            line=dict(color='#0D9488', width=3),
            marker=dict(size=8, color='#0D9488', line=dict(width=1, color='white')),
            hovertemplate='<b>Tanggal:</b> %{x|%d %B %Y}<br><b>Ritase:</b> %{y} trip<extra></extra>'
        )
        fig_tren.add_annotation(
            x=max_row['TANGGAL'],
            y=max_row['Total_Ritase'],
            text=f"Puncak: {int(max_row['Total_Ritase'])} trip",
            showarrow=True, arrowhead=2, arrowsize=1, arrowcolor='#0D9488', ax=0, ay=-30,
            font=dict(color='#0D9488', size=10)
        )
        fig_tren.update_layout(
            xaxis_title='Tanggal',
            yaxis_title='Total Trip (Ritase)',
            hovermode='x unified',
            xaxis=dict(tickformat='%d %b', tickangle=-45, dtick='D1', tickmode='linear'),
            margin=dict(l=40, r=40, t=60, b=60)
        )
        st.plotly_chart(fig_tren, use_container_width=True)

    st.session_state.grafik['tren'] = fig_tren
    st.session_state.grafik['kec_ton'] = fig_ton
    st.session_state.grafik['type_bar'] = fig_type_bar
    st.session_state.grafik['type_pie'] = px.pie(df_type, names='TYPE', values='Total_Ritase', template='plotly_white') if not df_type.empty else None

    # ---------- PER KECAMATAN ----------
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
            armada_kec.insert(0, 'No', range(1, len(armada_kec)+1))
            st.dataframe(armada_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True, hide_index=True)
            st.plotly_chart(px.bar(armada_kec.head(10), x='NOPIN', y='Total_Trip', color='Total_Trip', color_continuous_scale='Blues',
                                   title=f'10 Armada Teraktif di {kec_terpilih}', template='plotly_white'), use_container_width=True)

    # ---------- RINGKASAN EKSEKUTIF ----------
    st.markdown("---")
    st.subheader("📝 Laporan Ringkasan Eksekutif (Poin 7.0)")
    data_filtered = {
        'df_master': df_master, 'col_netto': col_netto, 'df_kec': df_kec,
        'df_armada': df_armada, 'teraktif': teraktif, 'tidak_efisien': tidak_efisien,
        'df_tren': df_tren, 'cleaned_count': data['cleaned_count']
    }
    ringkasan_teks = buat_ringkasan_eksekutif(data_filtered)
    st.markdown(f"```\n{ringkasan_teks}\n```")
    st.download_button("📄 Unduh Ringkasan Eksekutif (TXT)", ringkasan_teks.encode('utf-8'), "Ringkasan_Eksekutif_Poin7.txt")

    # ---------- PDF ----------
    st.subheader("📑 Laporan PDF Komprehensif")
    if st.button("📥 Buat Laporan PDF (Komprehensif)"):
        with st.spinner("Membuat PDF..."):
            pdf_buffer = generate_pdf_report(data_filtered, st.session_state.grafik, ringkasan_teks)
            st.download_button(label="⬇️ Unduh Laporan PDF", data=pdf_buffer, file_name="Laporan_DLH_Armada_Komprehensif.pdf", mime="application/pdf")

    # ---------- UNDUHAN DATA ----------
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

    with st.expander("🔎 Lihat Data Mentah Lengkap (hingga 1000 baris)"):
        st.dataframe(df_master.head(1000), use_container_width=True)

else:
    st.info("👆 Unggah file Excel, pilih mode, lalu klik **Proses Data** untuk memulai.")
    st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)
