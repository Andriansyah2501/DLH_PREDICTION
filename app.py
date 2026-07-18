import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
from io import BytesIO

# -------------------------- Fungsi Bantu --------------------------
def normalisasi_nopin(nopin_str):
    """Hilangkan semua karakter selain huruf dan angka, lalu uppercase."""
    if pd.isna(nopin_str):
        return ''
    return re.sub(r'[^A-Z0-9]', '', str(nopin_str).upper())

def cari_kolom(daftar_kolom, kata_kunci):
    """Cari kolom yang mengandung salah satu kata kunci (case‑insensitive)."""
    for col in daftar_kolom:
        if any(kw in str(col).upper() for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file Excel...")
def baca_semua_sheet(uploaded_file):
    """Baca semua sheet sebagai DataFrame mentah (header=None) agar tidak ada header yang terlewat."""
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('.xlsx') else None)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name, header=None)
            if not df.empty:
                sheets[name] = df
        except:
            pass
    return sheets

def proses_data(sheets_dict, config):
    """
    Memproses data dengan konfigurasi yang diberikan (otomatis/manual).
    Seluruh sheet sudah dalam bentuk mentah (header=None), sehingga kita bebas mencari header sendiri.
    """
    armada_sheet = config.get('armada_sheet')
    daily_sheets = config.get('daily_sheets', [])

    ref_df = None
    ref_dict = {}

    # =================== 1. PROSES LIST ARMADA ===================
    if armada_sheet and armada_sheet in sheets_dict:
        df_arm_raw = sheets_dict[armada_sheet].copy()

        # --- Deteksi header armada ---
        if 'col_nopin_arm' not in config:   # Mode otomatis
            header_arm = 0
            for idx, row in df_arm_raw.iterrows():
                row_str = " ".join(row.astype(str).dropna().str.upper().values)
                if 'NOPIN' in row_str or 'NO.PLAT' in row_str or 'PINTU' in row_str:
                    header_arm = idx
                    break
            # Potong data setelah header
            if header_arm > 0:
                df_ref = df_arm_raw.iloc[header_arm+1:].reset_index(drop=True)
                # Baris header dijadikan nama kolom
                header_row = df_arm_raw.iloc[header_arm].astype(str).str.strip().str.upper()
                df_ref.columns = [str(c).strip().upper() for c in header_row]
            else:
                # Gunakan baris pertama sebagai header (fallback)
                df_ref = df_arm_raw.iloc[1:].reset_index(drop=True) if len(df_arm_raw) > 1 else df_arm_raw
                if len(df_arm_raw) > 0:
                    header_row = df_arm_raw.iloc[0].astype(str).str.strip().str.upper()
                    df_ref.columns = [str(c).strip().upper() for c in header_row]
            # Cari kolom penting secara otomatis
            col_nopin_arm = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
            col_plat_arm = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
            col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI', 'KEC'])
            col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
            col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
        else:
            # Mode manual: kolom sudah dipilih user
            df_ref = df_arm_raw.copy()
            col_nopin_arm = config['col_nopin_arm']
            col_plat_arm = config['col_plat_arm']
            col_kec = config.get('col_kec_arm')
            col_merk = config.get('col_merk_arm')
            col_type = config.get('col_type_arm')

        # Jika kolom NOPIN dan Plat ditemukan
        if col_nopin_arm and col_plat_arm:
            # Normalisasi NOPIN
            df_ref['NOPIN_RAW'] = df_ref[col_nopin_arm].astype(str).str.strip().str.upper()
            df_ref['NOPIN_NORM'] = df_ref['NOPIN_RAW'].apply(normalisasi_nopin)
            df_ref['NO.PLAT'] = df_ref[col_plat_arm].astype(str).str.strip().str.upper()
            # Buat dictionary dengan key NOPIN_NORM
            for _, row in df_ref.iterrows():
                key = row['NOPIN_NORM']
                ref_dict[key] = {'NO.PLAT': row['NO.PLAT']}
                if col_kec: ref_dict[key]['Kecamatan'] = str(row[col_kec]).strip() if pd.notna(row[col_kec]) else ''
                if col_merk: ref_dict[key]['MERK'] = str(row[col_merk]).strip() if pd.notna(row[col_merk]) else ''
                if col_type: ref_dict[key]['TYPE'] = str(row[col_type]).strip() if pd.notna(row[col_type]) else ''
            # DataFrame referensi untuk merge (menggunakan NOPIN_NORM sebagai kunci)
            ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN_NORM'})

    # =================== 2. PROSES SHEET HARIAN ===================
    cleaned = {}
    skipped = []
    for sheet in daily_sheets:
        if sheet not in sheets_dict:
            skipped.append(sheet)
            continue

        df_raw = sheets_dict[sheet].copy()

        # --- Deteksi header harian ---
        if 'col_nopin_day' not in config:   # Mode otomatis
            header_harian = None
            for idx, row in df_raw.iterrows():
                row_str = " ".join(row.astype(str).dropna().str.upper().values)
                if 'PINTU' in row_str or 'PLAT MOBIL' in row_str or 'NOPIN' in row_str:
                    header_harian = idx
                    break
            if header_harian is None:
                skipped.append(sheet)
                continue
            try:
                # Data setelah header
                df_hari = df_raw.iloc[header_harian+1:].reset_index(drop=True)
                header_row = df_raw.iloc[header_harian].astype(str).str.strip().str.upper()
                df_hari.columns = [str(c).strip().upper() for c in header_row]
            except:
                skipped.append(sheet)
                continue
            col_nopin_day = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
            col_plat_day = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
        else:
            # Mode manual
            df_hari = df_raw.copy()
            col_nopin_day = config['col_nopin_day']
            col_plat_day = config['col_plat_day']

        if not col_nopin_day or not col_plat_day:
            skipped.append(sheet)
            continue

        df_hari = df_hari.rename(columns={col_nopin_day: 'NOPIN', col_plat_day: 'NO_PLAT'})

        # --- Pembersihan NOPIN ---
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        # Hapus baris yang merupakan summary
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        # Normalisasi untuk pencocokan
        df_hari['NOPIN_NORM'] = df_hari['NOPIN'].apply(normalisasi_nopin)
        # Hapus NOPIN yang setelah normalisasi kosong
        df_hari = df_hari[df_hari['NOPIN_NORM'] != '']

        # Simpan NO_PLAT asli untuk fallback
        no_plat_asli = df_hari['NO_PLAT'].copy() if 'NO_PLAT' in df_hari.columns else pd.Series('', index=df_hari.index)

        # --- Sinkronisasi dengan master ---
        if ref_df is not None and not ref_df.empty:
            # Hapus dulu kolom yang akan diisi ulang dari master
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns:
                    df_hari.drop(columns=[col], inplace=True)
            # Merge berdasarkan NOPIN_NORM
            df_hari = df_hari.merge(ref_df, on='NOPIN_NORM', how='left')
            # Jika NO_PLAT hasil merge kosong, gunakan asli
            if 'NO.PLAT' in df_hari.columns:
                df_hari.rename(columns={'NO.PLAT': 'NO_PLAT'}, inplace=True)
            if 'NO_PLAT' in df_hari.columns:
                df_hari['NO_PLAT'] = df_hari['NO_PLAT'].fillna(no_plat_asli)
            else:
                df_hari['NO_PLAT'] = no_plat_asli
        else:
            if 'NO_PLAT' not in df_hari.columns:
                df_hari['NO_PLAT'] = no_plat_asli
            for col in ['Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns:
                    df_hari[col] = ''

        # Isi Kecamatan yang kosong dengan 'Tidak Diketahui'
        if 'Kecamatan' not in df_hari.columns:
            df_hari['Kecamatan'] = 'Tidak Diketahui'
        else:
            df_hari['Kecamatan'] = df_hari['Kecamatan'].fillna('Tidak Diketahui').replace('', 'Tidak Diketahui')

        # --- Tambah Tanggal ---
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        # Hapus duplikasi kolom
        df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]
        cleaned[sheet] = df_hari

    if not cleaned:
        return None

    # Gabungkan semua sheet harian
    df_master = pd.concat(cleaned.values(), ignore_index=True, sort=False)

    # Konversi numerik untuk tonase
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # Analisis waktu (jika kolom jam tersedia)
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['KELUAR_DT'] - df_master['MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        try:
            df_master['JAM_INPUT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce').dt.hour
        except:
            pass

    # --- Agregasi ---
    df_armada = pd.DataFrame()
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        df_armada = df_master.groupby(group_cols, dropna=False).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)

    teraktif = df_armada.iloc[0] if not df_armada.empty else None
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty and (df_armada['Total_Trip'] > 0).any() else None

    df_waktu = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']

    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)

    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('TANGGAL')

    return {
        'df_master': df_master,
        'df_armada': df_armada,
        'teraktif': teraktif,
        'tidak_efisien': tidak_efisien,
        'df_waktu_jenis': df_waktu,
        'df_kec': df_kec,
        'df_tren': df_tren,
        'col_netto': col_netto,
        'skipped': skipped,
        'cleaned_count': len(cleaned)
    }

# -------------------------- SESSION STATE --------------------------
if "hasil" not in st.session_state:
    st.session_state.hasil = None
if "sheets" not in st.session_state:
    st.session_state.sheets = None
if "config" not in st.session_state:
    st.session_state.config = None

# -------------------------- ANTARMUKA --------------------------
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
st.markdown("Unggah file Excel, lalu pilih mode **Otomatis** atau **Manual**.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        st.session_state.sheets = baca_semua_sheet(uploaded_file)
        if not st.session_state.sheets:
            st.error("File tidak memiliki sheet yang valid.")
        else:
            st.success(f"Terbaca {len(st.session_state.sheets)} sheet.")

    if st.session_state.sheets:
        st.markdown("---")
        st.header("⚙️ Mode Pemrosesan")
        mode = st.radio("Pilih mode", ["Otomatis", "Manual"])
        if mode == "Manual":
            with st.expander("Pengaturan Manual"):
                sheet_names = list(st.session_state.sheets.keys())
                armada_sheet = st.selectbox("Sheet List Armada", sheet_names)
                daily_candidates = [s for s in sheet_names if s.isdigit()]
                if not daily_candidates:
                    daily_candidates = sheet_names
                daily_sheets = st.multiselect("Sheet Harian", daily_candidates, default=daily_candidates)
                # Pemetaan kolom armada
                if armada_sheet:
                    # Ambil contoh tampilan kolom dari sheet armada (asumsi header di baris pertama)
                    df_arm = st.session_state.sheets[armada_sheet].iloc[0].values.tolist()
                    cols_arm = [str(x) for x in df_arm]
                    col_nopin_arm = st.selectbox("Kolom NOPIN di List Armada", cols_arm)
                    col_plat_arm = st.selectbox("Kolom Plat di List Armada", cols_arm)
                    col_kec_arm = st.selectbox("Kolom Kecamatan (opsional)", ["(tidak ada)"] + cols_arm)
                    col_merk_arm = st.selectbox("Kolom Merk (opsional)", ["(tidak ada)"] + cols_arm)
                    col_type_arm = st.selectbox("Kolom Type (opsional)", ["(tidak ada)"] + cols_arm)
                else:
                    col_nopin_arm = col_plat_arm = col_kec_arm = col_merk_arm = col_type_arm = None
                # Pemetaan kolom harian (dari sheet harian pertama)
                if daily_sheets:
                    df_day = st.session_state.sheets[daily_sheets[0]].iloc[0].values.tolist()
                    cols_day = [str(x) for x in df_day]
                    col_nopin_day = st.selectbox("Kolom NOPIN di Harian", cols_day)
                    col_plat_day = st.selectbox("Kolom Plat di Harian", cols_day)
                else:
                    col_nopin_day = col_plat_day = None
                config = {
                    'armada_sheet': armada_sheet,
                    'daily_sheets': daily_sheets,
                    'col_nopin_arm': col_nopin_arm,
                    'col_plat_arm': col_plat_arm,
                    'col_kec_arm': col_kec_arm if col_kec_arm != "(tidak ada)" else None,
                    'col_merk_arm': col_merk_arm if col_merk_arm != "(tidak ada)" else None,
                    'col_type_arm': col_type_arm if col_type_arm != "(tidak ada)" else None,
                    'col_nopin_day': col_nopin_day,
                    'col_plat_day': col_plat_day
                }
                st.session_state.config = config
        else:
            # Mode otomatis: reset config agar deteksi berjalan
            st.session_state.config = {'armada_sheet': None, 'daily_sheets': []}

        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Memproses..."):
                # Tentukan sheet otomatis jika diperlukan
                if mode == "Otomatis" or not st.session_state.config.get('daily_sheets'):
                    armada = next((s for s in st.session_state.sheets if 'list armada' in s.lower()), None)
                    if armada is None:
                        armada = next((s for s in st.session_state.sheets if 'armada' in s.lower()), None)
                    daily = [s for s in st.session_state.sheets if s.isdigit()]
                    if not daily:
                        daily = [s for s in st.session_state.sheets if s != armada and s not in ['Tugas', 'Master Data']]
                    st.session_state.config['armada_sheet'] = armada
                    st.session_state.config['daily_sheets'] = daily
                # Proses
                hasil = proses_data(st.session_state.sheets, st.session_state.config)
                if hasil is None:
                    st.error("Gagal memproses data. Periksa sheet/kolom yang dipilih.")
                else:
                    st.session_state.hasil = hasil
                    st.success(f"✅ {hasil['cleaned_count']} sheet berhasil digabung. {len(hasil['skipped'])} sheet dilewati.")
                    st.balloons()

# Tampilkan hasil
if st.session_state.hasil is not None:
    data = st.session_state.hasil
    df_master = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']

    st.sidebar.markdown("---")
    st.sidebar.header("📍 Analisis per Kecamatan")
    if 'Kecamatan' in df_master.columns:
        daftar_kec = sorted(df_master['Kecamatan'].unique().tolist())
        kec_terpilih = st.sidebar.selectbox("Pilih Kecamatan", daftar_kec)
        df_kec_filter = df_master[df_master['Kecamatan'] == kec_terpilih]
    else:
        kec_terpilih = None
        df_kec_filter = df_master

    st.subheader(f"📌 Statistik Kecamatan: **{kec_terpilih}**")
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
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)
        st.dataframe(armada_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True)

        fig_top_kec = px.bar(armada_kec.head(10), x='NOPIN', y='Total_Trip',
                             color='Total_Trip', color_continuous_scale='Blues',
                             title=f'10 Armada Teraktif di {kec_terpilih}')
        st.plotly_chart(fig_top_kec, use_container_width=True)

    st.markdown("---")
    st.subheader("📊 Ringkasan Seluruh Kecamatan")
    if not df_kec.empty:
        col1, col2 = st.columns([2, 1])
        with col1:
            st.dataframe(df_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True)
        with col2:
            fig_pie = px.pie(df_kec, names='Kecamatan', values='Total_Ritase',
                             title='Distribusi Trip per Kecamatan')
            st.plotly_chart(fig_pie, use_container_width=True)
        fig_ton = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', color='Total_Tonase',
                         color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
        fig_ton.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_ton, use_container_width=True)

    if not data['df_tren'].empty:
        st.subheader("📈 Tren Harian (Semua Kecamatan)")
        fig_tren = px.line(data['df_tren'], x='TANGGAL', y='Total_Ritase',
                           title='Tren Frekuensi Ritase Harian', markers=True)
        fig_tren.update_traces(line_color='#0D9488')
        st.plotly_chart(fig_tren, use_container_width=True)

    st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien (Keseluruhan)")
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']
    if teraktif is not None:
        col_a, col_b = st.columns(2)
        with col_a:
            st.success(f"**Teraktif:** {teraktif.get('NOPIN', '-')} ({teraktif.get('NO_PLAT', '')}) – {int(teraktif.get('Total_Trip', 0))} trip")
        with col_b:
            if tidak_efisien is not None:
                st.error(f"**Tidak Efisien:** {tidak_efisien.get('NOPIN', '-')} ({tidak_efisien.get('NO_PLAT', '')}) – {int(tidak_efisien.get('Total_Trip', 0))} trip")

    if not data['df_waktu_jenis'].empty:
        st.subheader("⏱️ Rata‑rata Waktu Tempuh per Jenis Armada")
        st.dataframe(data['df_waktu_jenis'].style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}))

    st.subheader("📝 Laporan Ringkasan Otomatis")
    kec_tertinggi = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else '-'
    laporan = f"""
**Ringkasan Operasional (Juni 2026):**
- Total trip: {len(df_master)}
- Armada aktif: {df_master['NOPIN'].nunique()} unit
- Total tonase: {df_master[col_netto].sum()/1000:,.1f} Ton
- Kecamatan tersibuk: **{kec_tertinggi}**
- Armada teraktif: {teraktif.get('NOPIN','-')} ({teraktif.get('NO_PLAT','')}) – {int(teraktif.get('Total_Trip',0))} trip
- Armada tidak efisien: {tidak_efisien.get('NOPIN','-')} ({tidak_efisien.get('NO_PLAT','')}) – {int(tidak_efisien.get('Total_Trip',0))} trip
"""
    st.markdown(laporan)
    st.download_button("📄 Unduh Ringkasan (TXT)", laporan.encode('utf-8'), "ringkasan.txt")

    st.subheader("📥 Unduh Data Hasil Analisis")
    @st.cache_data
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as w:
            dataframe.to_excel(w, index=False)
        return output.getvalue()

    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        st.download_button("📊 Master Data (Excel)", to_excel(df_master), "Master_Data.xlsx")
        st.download_button("📊 Statistik Armada (Excel)", to_excel(data['df_armada']), "Statistik_Armada.xlsx")
    with col_u2:
        st.download_button("📊 Laporan per Kecamatan (Excel)", to_excel(df_kec), "Kecamatan.xlsx")
        st.download_button("📈 Tren Harian (Excel)", to_excel(data['df_tren']), "Tren_Harian.xlsx")
    with col_u3:
        if not data['df_waktu_jenis'].empty:
            st.download_button("⏱️ Waktu per Jenis (Excel)", to_excel(data['df_waktu_jenis']), "Waktu_per_Jenis.xlsx")

    with st.expander("🔎 Lihat Data Mentah (200 baris pertama)"):
        st.dataframe(df_master.head(200))

else:
    st.info("👆 Unggah file Excel, pilih mode, lalu klik **Proses Data** untuk memulai.")
    st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)
