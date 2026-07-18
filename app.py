import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# -------------------------- Fungsi Bantu --------------------------
def cari_kolom(kolom_list, kata_kunci):
    """Mencari kolom yang mengandung salah satu kata kunci (case‑insensitive)."""
    for col in kolom_list:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file Excel...")
def baca_semua_sheet(uploaded_file):
    """Baca semua sheet satu kali. Return dict {nama_sheet: dataframe}."""
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('.xlsx') else None)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except Exception:
            pass
    return sheets

# -------------------------- Proses Data Utama (Tugas 1–5) --------------------------
def proses_utama(sheets_dict):
    """
    Menjalankan Tugas 1–5:
    - Gabung 30 sheet
    - Sinkronisasi dengan List Armada
    - Hitung trip & tonase
    - Cari armada teraktif & tidak efisien
    - Rata‑rata waktu tempuh per jenis armada
    Return dictionary berisi seluruh hasil.
    """
    # 1. Cari sheet List Armada
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)

    ref_df = None
    if armada_sheet:
        df_arm_raw = sheets_dict[armada_sheet].copy()
        # Deteksi header: cari baris yang mengandung NOPIN / NO.PLAT
        header_arm = 0
        for idx, row in df_arm_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'NOPIN' in row_str or 'NO.PLAT' in row_str:
                header_arm = idx
                break
        # Baca dengan header yang ditemukan
        if header_arm > 0:
            df_ref = df_arm_raw.iloc[header_arm:].reset_index(drop=True)
            df_ref.columns = df_arm_raw.iloc[header_arm].astype(str).str.strip().str.upper()
        else:
            # Fallback: gunakan baris pertama sebagai header
            df_ref = df_arm_raw.copy()
            df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

        # Cari kolom NOPIN dan NO.PLAT
        col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
        col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
        if col_nopin and col_plat:
            df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
            df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
            # Kolom tambahan jika ada
            col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
            col_merk = cari_kolom(df_ref.columns, ['MERK'])
            col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
            # Buat dictionary referensi
            ref_dict = {}
            for _, row in df_ref.iterrows():
                nopin = row['NOPIN']
                ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
                if col_kec: ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip()
                if col_merk: ref_dict[nopin]['MERK'] = str(row[col_merk]).strip()
                if col_type: ref_dict[nopin]['TYPE'] = str(row[col_type]).strip()
            # DataFrame referensi untuk merge
            ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN'})
        else:
            ref_dict = {}
            ref_df = None
    else:
        ref_dict = {}
        ref_df = None

    # 2. Pilih sheet harian (digit 1-30)
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]

    cleaned = {}
    skipped = []
    for sheet in daily_sheets:
        try:
            df_raw = sheets_dict[sheet].copy()
        except Exception:
            skipped.append(sheet)
            continue

        # Deteksi header harian
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str or 'NOPIN' in row_str:
                header_idx = idx
                break
        if header_idx is None:
            skipped.append(sheet)
            continue

        try:
            df_hari = df_raw.iloc[header_idx+1:].reset_index(drop=True)
            header_row = df_raw.iloc[header_idx].astype(str).str.strip().str.upper()
            df_hari.columns = [str(c).strip().upper() for c in header_row]
        except Exception:
            skipped.append(sheet)
            continue

        col_nopin_h = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
        if not col_nopin_h or not col_plat_h:
            skipped.append(sheet)
            continue

        df_hari = df_hari.rename(columns={col_nopin_h: 'NOPIN', col_plat_h: 'NO_PLAT'})

        # Pembersihan
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

        # Reset indeks
        df_hari = df_hari.reset_index(drop=True)

        # Sinkronisasi dengan master (jika tersedia)
        if ref_df is not None and 'NO_PLAT' in df_hari.columns:
            # Hapus kolom yang akan diupdate dari master (kecuali NOPIN sebagai kunci)
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns:
                    df_hari.drop(columns=[col], inplace=True)
            # Merge dengan data referensi
            df_hari = df_hari.merge(ref_df[['NOPIN', 'NO.PLAT'] + 
                                           ([c for c in ['Kecamatan', 'MERK', 'TYPE'] if c in ref_df.columns])], 
                                    on='NOPIN', how='left')
        else:
            # Jika tidak ada referensi, pastikan kolom NO_PLAT tetap ada
            if 'NO_PLAT' not in df_hari.columns:
                df_hari['NO_PLAT'] = ''
            for col in ['Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns:
                    df_hari[col] = ''

        # Tambah kolom TANGGAL
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        cleaned[sheet] = df_hari

    if not cleaned:
        return None  # Tidak ada data valid

    df_master = pd.concat(cleaned.values(), ignore_index=True)

    # Konversi numerik untuk kolom tonase
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # Hitung durasi jika kolom waktu tersedia
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

    # Agregasi trip & tonase per armada (Tugas 3 & 4)
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        df_armada = df_master.groupby(group_cols, dropna=False).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)
    else:
        df_armada = pd.DataFrame()

    # Armada teraktif & tidak efisien
    teraktif = df_armada.iloc[0] if not df_armada.empty else None
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty and any(df_armada['Total_Trip'] > 0) else None

    # Rata‑rata waktu per jenis armada (Tugas 5)
    df_waktu_jenis = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu_jenis = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu_jenis.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']

    # Agregasi per kecamatan untuk grafik
    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Ritase', ascending=False)

    # Tren harian
    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('TANGGAL')

    # Kembalikan semua dalam dictionary
    hasil = {
        'df_master': df_master,
        'df_armada': df_armada,
        'teraktif': teraktif,
        'tidak_efisien': tidak_efisien,
        'df_waktu_jenis': df_waktu_jenis,
        'df_kec': df_kec,
        'df_tren': df_tren,
        'col_netto': col_netto,
        'skipped': skipped,
        'cleaned_count': len(cleaned)
    }
    return hasil

# -------------------------- Inisialisasi Session State --------------------------
if "hasil" not in st.session_state:
    st.session_state.hasil = None
if "laporan_ai" not in st.session_state:
    st.session_state.laporan_ai = None

# -------------------------- Antarmuka Streamlit --------------------------
st.set_page_config(page_title="Dashboard Analitik Armada DLH", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
st.markdown("Unggah file Excel dengan 30 sheet harian dan 1 sheet **List Armada**. Dashboard ini akan menjalankan seluruh analisis secara otomatis.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Membaca file dan memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if not sheets:
                    st.error("File tidak memiliki sheet yang valid.")
                else:
                    hasil = proses_utama(sheets)
                    if hasil is None:
                        st.error("Tidak ada data harian yang valid. Proses gagal.")
                    else:
                        st.session_state.hasil = hasil
                        st.session_state.laporan_ai = None
                        st.success(f"✅ Proses selesai. {hasil['cleaned_count']} sheet berhasil digabung. {len(hasil['skipped'])} sheet dilewati.")
                        st.balloons()

# Tampilkan hasil jika sudah diproses
if st.session_state.hasil is not None:
    hasil = st.session_state.hasil
    df_master = hasil['df_master']
    df_armada = hasil['df_armada']
    teraktif = hasil['teraktif']
    tidak_efisien = hasil['tidak_efisien']
    df_waktu_jenis = hasil['df_waktu_jenis']
    df_kec = hasil['df_kec']
    df_tren = hasil['df_tren']
    col_netto = hasil['col_netto']

    # Sidebar filter (opsional)
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filter Data")
    if 'Kecamatan' in df_master.columns:
        kec_list = ['Semua'] + sorted(df_master['Kecamatan'].dropna().unique().tolist())
        kec_terpilih = st.sidebar.selectbox("Kecamatan", kec_list)
    else:
        kec_terpilih = 'Semua'
    if 'TANGGAL' in df_master.columns:
        tgl_list = sorted(df_master['TANGGAL'].unique())
        if len(tgl_list) > 1:
            rentang = st.sidebar.date_input("Rentang Tanggal", 
                [pd.to_datetime(tgl_list[0]), pd.to_datetime(tgl_list[-1])])
    # Aplikasi filter
    df = df_master.copy()
    if kec_terpilih != 'Semua':
        df = df[df['Kecamatan'] == kec_terpilih]
    if 'rentang' in locals() and len(rentang) == 2:
        df = df[(pd.to_datetime(df['TANGGAL']) >= pd.Timestamp(rentang[0])) & 
                (pd.to_datetime(df['TANGGAL']) <= pd.Timestamp(rentang[1]))]

    # --- Metrik Utama ---
    total_trip = len(df)
    total_armada = df['NOPIN'].nunique()
    total_tonase = df[col_netto].sum() / 1000 if col_netto else 0
    durasi_rata = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else None

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Trip", total_trip)
    col2.metric("Armada Aktif", total_armada)
    col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
    col4.metric("Rata² Durasi (menit)", f"{durasi_rata:.1f}" if durasi_rata else "-")

    st.markdown("---")

    # --- Tugas 4: Armada Teraktif & Tidak Efisien ---
    if teraktif is not None:
        st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien")
        colA, colB = st.columns(2)
        with colA:
            st.success(f"**Teraktif:** {teraktif['NOPIN']} ({teraktif['NO_PLAT']}) – {int(teraktif['Total_Trip'])} trip")
        with colB:
            if tidak_efisien is not None:
                st.error(f"**Tidak Efisien:** {tidak_efisien['NOPIN']} ({tidak_efisien['NO_PLAT']}) – {int(tidak_efisien['Total_Trip'])} trip")

    # --- Tugas 5: Waktu per Jenis Armada ---
    if not df_waktu_jenis.empty:
        st.subheader("⏱️ Rata‑rata Waktu Tempuh per Jenis Armada")
        st.dataframe(df_waktu_jenis.style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}))

    st.markdown("---")

    # --- Tugas 6: Grafik Interaktif ---
    st.subheader("📊 Visualisasi Data")
    if col_netto:
        # Tren Harian
        tren = df.groupby('TANGGAL', dropna=False).size().reset_index(name='Ritase')
        fig1 = px.line(tren, x='TANGGAL', y='Ritase', title='Tren Ritase Harian', markers=True)
        fig1.update_traces(line_color='#0D9488')
        st.plotly_chart(fig1, use_container_width=True)

        # Distribusi Tonase per Kecamatan
        if 'Kecamatan' in df.columns:
            kec = df.groupby('Kecamatan', dropna=False)[col_netto].sum().reset_index(name='Tonase')
            kec = kec.sort_values('Tonase', ascending=False)
            fig2 = px.bar(kec, x='Kecamatan', y='Tonase', color='Tonase',
                          color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
            fig2.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig2, use_container_width=True)

        # 10 Armada Teraktif
        if not df_armada.empty:
            top10 = df_armada.head(10)
            fig3 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip',
                          color_continuous_scale='OrRd', title='10 Armada Teraktif')
            st.plotly_chart(fig3, use_container_width=True)

        # Pola Jam Sibuk (jika tersedia)
        if 'JAM_INPUT' in df.columns:
            jam_counts = df.dropna(subset=['JAM_INPUT']).groupby('JAM_INPUT').size().reset_index(name='Jumlah')
            fig4 = px.area(jam_counts, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
            st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # --- Tugas 7: Laporan Ringkasan ---
    st.subheader("📝 Laporan Ringkasan Otomatis")
    laporan_teks = f"""
**Ringkasan Operasional:**
- Total trip (ritase): {total_trip}
- Armada aktif: {total_armada} unit
- Total volume sampah: {total_tonase:,.1f} Ton
- Kecamatan dengan aktivitas tertinggi: {df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else '-'}
- Armada teraktif: {teraktif['NOPIN'] if teraktif is not None else '-'}
- Armada paling tidak efisien: {tidak_efisien['NOPIN'] if tidak_efisien is not None else '-'}
"""
    st.markdown(laporan_teks)

    # Unduh ringkasan sebagai TXT
    st.download_button("📄 Unduh Ringkasan (TXT)", laporan_teks.encode('utf-8'), "ringkasan.txt")

    # Unduh data
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
        if not df_armada.empty:
            st.download_button("📊 Statistik Armada (Excel)", to_excel(df_armada), "Statistik_Armada.xlsx")
    with col_u2:
        if not df_kec.empty:
            st.download_button("📊 Laporan Kecamatan (Excel)", to_excel(df_kec), "Kecamatan.xlsx")
        if not df_tren.empty:
            st.download_button("📈 Tren Harian (Excel)", to_excel(df_tren), "Tren_Harian.xlsx")
    with col_u3:
        if not df_waktu_jenis.empty:
            st.download_button("⏱️ Waktu per Jenis (Excel)", to_excel(df_waktu_jenis), "Waktu_per_Jenis.xlsx")

    # Tampilkan data mentah (opsional)
    with st.expander("🔎 Lihat Data Mentah (200 baris pertama)"):
        st.dataframe(df_master.head(200))

else:
    st.info("👆 Silakan unggah file Excel dan klik **Proses Data** untuk memulai.")
    st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)
