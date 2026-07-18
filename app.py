import streamlit as st
import pandas as pd
import numpy as np
import re
import plotly.express as px
from io import BytesIO

# ------------------------------------------------------------
# KONFIGURASI HALAMAN
# ------------------------------------------------------------
st.set_page_config(page_title="Dashboard Armada DLH", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada Pengangkut Sampah")
st.markdown("Unggah file Excel bulanan, dan dashboard akan otomatis memproses seluruh sheet.")

# ------------------------------------------------------------
# FUNGSI BANTU
# ------------------------------------------------------------
def normalisasi_nopin(val):
    """Hapus semua karakter selain huruf dan angka, lalu uppercase."""
    if pd.isna(val):
        return ''
    return re.sub(r'[^A-Z0-9]', '', str(val).upper())

def cari_kolom(daftar_kolom, kata_kunci):
    """Cari kolom yang mengandung salah satu kata kunci (case‑insensitive)."""
    for col in daftar_kolom:
        if any(kw in str(col).upper() for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca dan memproses file Excel...")
def pipeline_data(uploaded_file):
    """
    Fungsi utama ETL & perhitungan.
    Mengembalikan dictionary berisi semua data dan metrik.
    """
    # Baca semua sheet
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        try:
            # Baca mentah (header=None) agar header tidak terlewat
            df = pd.read_excel(xls, sheet_name=name, header=None)
            if not df.empty:
                sheets[name] = df
        except Exception:
            pass

    if not sheets:
        return None

    # --------------------------------------------------------
    # 1. IDENTIFIKASI SHEET LIST ARMADA
    # --------------------------------------------------------
    armada_sheet = next((s for s in sheets if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets if 'armada' in s.lower()), None)

    ref_dict = {}
    if armada_sheet:
        df_arm_raw = sheets[armada_sheet].copy()
        # Deteksi header – cari baris yang mengandung NOPIN / NO.PLAT / PINTU
        header_arm = 0
        for idx, row in df_arm_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'NOPIN' in row_str or 'NO.PLAT' in row_str or 'PINTU' in row_str:
                header_arm = idx
                break
        if header_arm > 0:
            df_ref = df_arm_raw.iloc[header_arm+1:].reset_index(drop=True)
            header_row = df_arm_raw.iloc[header_arm].astype(str).str.strip().str.upper()
            df_ref.columns = [str(c).strip().upper() for c in header_row]
        else:
            # Fallback: gunakan baris pertama sebagai header
            df_ref = df_arm_raw.iloc[1:].reset_index(drop=True) if len(df_arm_raw) > 1 else df_arm_raw
            if len(df_arm_raw) > 0:
                header_row = df_arm_raw.iloc[0].astype(str).str.strip().str.upper()
                df_ref.columns = [str(c).strip().upper() for c in header_row]

        col_nopin_arm = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
        col_plat_arm = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
        col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI', 'KEC'])
        col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
        col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])

        if col_nopin_arm and col_plat_arm:
            df_ref['NOPIN_NORM'] = df_ref[col_nopin_arm].apply(normalisasi_nopin)
            df_ref['NO.PLAT'] = df_ref[col_plat_arm].astype(str).str.strip().str.upper()
            for _, row in df_ref.iterrows():
                key = row['NOPIN_NORM']
                ref_dict[key] = {'NO.PLAT': row['NO.PLAT']}
                if col_kec:
                    ref_dict[key]['Kecamatan'] = str(row[col_kec]).strip()
                if col_merk:
                    ref_dict[key]['MERK'] = str(row[col_merk]).strip()
                if col_type:
                    ref_dict[key]['TYPE'] = str(row[col_type]).strip()
            # Buat DataFrame referensi untuk left join
            ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN_NORM'})
        else:
            st.warning("Kolom NOPIN/Plat tidak lengkap di List Armada. Sinkronisasi dilewati.")
            ref_df = None
    else:
        st.info("Sheet 'List Armada' tidak ditemukan. Data harian akan diproses tanpa master.")
        ref_df = None

    # --------------------------------------------------------
    # 2. PROSES SHEET HARIAN (NAMA SHEET BERUPA ANGKA)
    # --------------------------------------------------------
    daily_sheets = [s for s in sheets if s.isdigit()]
    if not daily_sheets:
        # Fallback: semua sheet selain armada & bukan 'Tugas', 'Master Data'
        daily_sheets = [s for s in sheets if s != armada_sheet and s not in ['Tugas', 'Master Data']]
    if not daily_sheets:
        st.error("Tidak ada sheet harian yang dapat diproses.")
        return None

    cleaned = {}
    skipped = []
    for sheet in daily_sheets:
        df_raw = sheets[sheet].copy()

        # Deteksi header harian (cari baris dengan PINTU / PLAT MOBIL / NOPIN)
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
            df_hari = df_raw.iloc[header_harian+1:].reset_index(drop=True)
            header_row = df_raw.iloc[header_harian].astype(str).str.strip().str.upper()
            df_hari.columns = [str(c).strip().upper() for c in header_row]
        except Exception:
            skipped.append(sheet)
            continue

        col_nopin_day = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
        col_plat_day = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
        if not col_nopin_day or not col_plat_day:
            skipped.append(sheet)
            continue
        df_hari.rename(columns={col_nopin_day: 'NOPIN', col_plat_day: 'NO_PLAT'}, inplace=True)

        # Pembersihan NOPIN
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        # Hapus akhiran .0 jika ada (float -> int string)
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
        df_hari['NOPIN_NORM'] = df_hari['NOPIN'].apply(normalisasi_nopin)
        df_hari = df_hari[df_hari['NOPIN_NORM'] != '']

        # Simpan NO_PLAT asli untuk fallback
        no_plat_asli = df_hari['NO_PLAT'].copy() if 'NO_PLAT' in df_hari.columns else pd.Series('', index=df_hari.index)

        # Left join dengan referensi, lalu timpa data harian yang tidak sesuai
        if ref_df is not None:
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns:
                    df_hari.drop(columns=[col], inplace=True)
            df_hari = df_hari.merge(ref_df, on='NOPIN_NORM', how='left')
            # Gunakan data dari master jika ada, fallback ke asli
            if 'NO.PLAT' in df_hari.columns:
                df_hari['NO_PLAT'] = df_hari['NO.PLAT'].fillna(no_plat_asli)
            else:
                df_hari['NO_PLAT'] = no_plat_asli
        else:
            if 'NO_PLAT' not in df_hari.columns:
                df_hari['NO_PLAT'] = no_plat_asli
            for col in ['Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns:
                    df_hari[col] = ''

        if 'Kecamatan' not in df_hari.columns:
            df_hari['Kecamatan'] = 'Tidak Diketahui'
        else:
            df_hari['Kecamatan'] = df_hari['Kecamatan'].fillna('Tidak Diketahui')

        # Tambah kolom tanggal dari nama sheet
        try:
            tgl = f"2026-06-{int(sheet):02d}"  # asumsi Juni 2026, bisa diubah
        except ValueError:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        # Hapus duplikasi kolom (jaga-jaga)
        df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]
        cleaned[sheet] = df_hari

    if not cleaned:
        st.error("Semua sheet harian gagal diproses.")
        return None

    df_master = pd.concat(cleaned.values(), ignore_index=True, sort=False)

    # Hapus duplikat baris yang mungkin muncul
    key_cols = ['NOPIN', 'TANGGAL', 'NO_PLAT', 'Kecamatan']
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        key_cols.append(ton_col)
    # Hanya gunakan kolom yang ada
    available_keys = [c for c in key_cols if c in df_master.columns]
    if available_keys:
        df_master.drop_duplicates(subset=available_keys, keep='first', inplace=True)

    # Konversi kolom tonase ke numerik
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # --------------------------------------------------------
    # 3. HITUNG DURASI (JAM MASUK - JAM KELUAR)
    # --------------------------------------------------------
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1', 'IN'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2', 'OUT'])
    if col_masuk and col_keluar:
        # Parsing waktu fleksibel
        def parse_time(series):
            if pd.api.types.is_datetime64_any_dtype(series):
                return series
            if series.dtype == object:
                series = series.astype(str).str.strip().str.replace(r'\.', ':', regex=True)
            for fmt in ['%H:%M:%S', '%H:%M', '%H.%M.%S', '%I:%M:%S %p', '%I:%M %p', '%H.%M']:
                dt = pd.to_datetime(series, format=fmt, errors='coerce')
                if dt.notna().sum() > 0:
                    return dt
            if pd.api.types.is_numeric_dtype(series):
                try:
                    hours = series * 24
                    td = pd.to_timedelta(hours, unit='h')
                    base = pd.Timestamp('2026-01-01')
                    return base + td
                except Exception:
                    pass
            return pd.to_datetime(series, errors='coerce', dayfirst=True)

        df_master['MASUK_DT'] = parse_time(df_master[col_masuk])
        df_master['KELUAR_DT'] = parse_time(df_master[col_keluar])
        # Hanya yang valid
        mask_valid = df_master['MASUK_DT'].notna() & df_master['KELUAR_DT'].notna()
        df_master['DURASI_MENIT'] = np.nan
        df_master.loc[mask_valid, 'DURASI_MENIT'] = (
            (df_master.loc[mask_valid, 'KELUAR_DT'] - df_master.loc[mask_valid, 'MASUK_DT']).dt.total_seconds() / 60
        )
        # Filter outlier
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)].copy()
        # Simpan data asli untuk ditampilkan
        df_master['MASUK_ORI'] = df_master[col_masuk].astype(str)
        df_master['KELUAR_ORI'] = df_master[col_keluar].astype(str)
    else:
        st.info("Kolom waktu masuk/keluar tidak ditemukan. Analisis durasi dilewati.")

    # --------------------------------------------------------
    # 4. AGREGASI & METRIK
    # --------------------------------------------------------
    # Armada teraktif & paling tidak efisien
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
        # Armada tidak efisien (trip > 0 paling sedikit)
        tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty else None
    else:
        df_armada = pd.DataFrame()
        teraktif = None
        tidak_efisien = None

    # Rata-rata durasi per Type
    df_waktu_type = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu_type = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE')['DURASI_MENIT'].mean().reset_index()
        df_waktu_type.columns = ['Type Armada', 'Rata-rata Durasi (menit)']

    # Agregasi per kecamatan (untuk visualisasi)
    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan').agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Ritase', ascending=False)

    # Tren harian
    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL').agg(Total_Ritase=('NOPIN', 'count')).reset_index()
        df_tren['TANGGAL'] = pd.to_datetime(df_tren['TANGGAL'])
        df_tren = df_tren.sort_values('TANGGAL')

    # Total metrik
    total_trip = len(df_master)
    total_armada = df_master['NOPIN'].nunique()
    total_tonase = df_master[col_netto].sum() / 1000 if col_netto else 0  # dalam Ton

    return {
        'df_master': df_master,
        'df_armada': df_armada,
        'teraktif': teraktif,
        'tidak_efisien': tidak_efisien,
        'df_waktu_type': df_waktu_type,
        'df_kec': df_kec,
        'df_tren': df_tren,
        'col_netto': col_netto,
        'total_trip': total_trip,
        'total_armada': total_armada,
        'total_tonase': total_tonase,
        'skipped_sheets': skipped,
        'armada_sheet_found': armada_sheet is not None,
        'daily_sheets_count': len(daily_sheets)
    }

# ------------------------------------------------------------
# SIDEBAR - UNGGAH FILE
# ------------------------------------------------------------
st.sidebar.header("📂 Unggah File Excel")
uploaded_file = st.sidebar.file_uploader(
    "Pilih file .xlsx atau .xls",
    type=['xlsx', 'xls'],
    help="File harus memiliki sheet 'List Armada' dan 30 sheet harian bernama angka."
)

if uploaded_file is not None:
    # Jalankan pipeline (cache otomatis)
    with st.spinner("Memproses data..."):
        hasil = pipeline_data(uploaded_file)

    if hasil is None:
        st.error("Gagal memproses file. Pastikan struktur file sesuai.")
    else:
        st.sidebar.success(f"✅ {hasil['daily_sheets_count']} sheet harian berhasil diolah.")

        # Ekstrak variabel
        df_master = hasil['df_master']
        df_armada = hasil['df_armada']
        teraktif = hasil['teraktif']
        tidak_efisien = hasil['tidak_efisien']
        df_waktu_type = hasil['df_waktu_type']
        df_kec = hasil['df_kec']
        df_tren = hasil['df_tren']
        col_netto = hasil['col_netto']
        total_trip = hasil['total_trip']
        total_tonase = hasil['total_tonase']

        # --------------------------------------------------------
        # METRIK UTAMA (SCORECARDS)
        # --------------------------------------------------------
        st.header("📊 Ringkasan Performa Bulan Ini")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Trip", total_trip)
        col2.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
        col3.metric("Armada Teraktif",
                    teraktif['NOPIN'] if teraktif is not None else "-",
                    help=f"{teraktif['Total_Trip']} trip" if teraktif is not None else "")
        col4.metric("Armada Tidak Efisien",
                    tidak_efisien['NOPIN'] if tidak_efisien is not None else "-",
                    help=f"{tidak_efisien['Total_Trip']} trip" if tidak_efisien is not None else "")

        # --------------------------------------------------------
        # TABS VISUALISASI
        # --------------------------------------------------------
        tab1, tab2, tab3 = st.tabs(["📈 Tren & Distribusi", "⏱️ Efisiensi Armada", "📋 Data Master"])

        with tab1:
            st.subheader("Tren Ritase Harian")
            if not df_tren.empty:
                fig_tren = px.line(df_tren, x='TANGGAL', y='Total_Ritase', markers=True,
                                   title='Frekuensi Trip per Tanggal',
                                   labels={'TANGGAL': 'Tanggal', 'Total_Ritase': 'Jumlah Trip'})
                fig_tren.update_traces(line_color='#0D9488')
                fig_tren.update_layout(xaxis_tickformat='%d %b', xaxis_tickangle=-45)
                st.plotly_chart(fig_tren, use_container_width=True)
            else:
                st.info("Data tren tidak tersedia.")

            st.subheader("Distribusi Tonase per Kecamatan")
            if not df_kec.empty:
                fig_kec = px.bar(df_kec, x='Kecamatan', y='Total_Tonase',
                                 color='Total_Tonase', color_continuous_scale='Viridis',
                                 title='Total Tonase (Kg) per Wilayah')
                fig_kec.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_kec, use_container_width=True)
            else:
                st.info("Data kecamatan tidak tersedia.")

        with tab2:
            st.subheader("Rata‑rata Durasi Kerja per Tipe Armada")
            if not df_waktu_type.empty:
                fig_dur = px.bar(df_waktu_type, x='Type Armada', y='Rata-rata Durasi (menit)',
                                 color='Rata-rata Durasi (menit)', color_continuous_scale='Blues',
                                 title='Durasi Rata‑rata per Tipe Armada (menit)')
                st.plotly_chart(fig_dur, use_container_width=True)
            else:
                st.info("Data durasi tidak tersedia. Periksa kolom Jam Masuk/Keluar.")

            st.subheader("Top 10 Armada Berdasarkan Trip")
            if not df_armada.empty:
                top10 = df_armada.head(10)
                fig_top = px.bar(top10, x='NOPIN', y='Total_Trip',
                                 color='Total_Trip', color_continuous_scale='OrRd',
                                 title='10 Armada dengan Trip Terbanyak')
                st.plotly_chart(fig_top, use_container_width=True)
            else:
                st.info("Data armada tidak tersedia.")

        with tab3:
            st.subheader("Master Data (100 baris pertama)")
            cols_show = ['TANGGAL', 'NOPIN', 'NO_PLAT', 'Kecamatan', 'TYPE', 'MERK']
            if col_netto and col_netto not in cols_show:
                cols_show.append(col_netto)
            if 'DURASI_MENIT' in df_master.columns:
                cols_show.append('DURASI_MENIT')
            cols_available = [c for c in cols_show if c in df_master.columns]
            st.dataframe(df_master[cols_available].head(100), use_container_width=True)

        # --------------------------------------------------------
        # LAPORAN EKSEKUTIF OTOMATIS
        # --------------------------------------------------------
        st.markdown("---")
        st.header("📝 Laporan Eksekutif Otomatis")
        # Bangun teks dinamis
        teraktif_nopin = teraktif['NOPIN'] if teraktif is not None else "N/A"
        teraktif_plat = teraktif['NO_PLAT'] if teraktif is not None else ""
        teraktif_trip = int(teraktif['Total_Trip']) if teraktif is not None else 0
        tidak_nopin = tidak_efisien['NOPIN'] if tidak_efisien is not None else "N/A"
        tidak_plat = tidak_efisien['NO_PLAT'] if tidak_efisien is not None else ""
        tidak_trip = int(tidak_efisien['Total_Trip']) if tidak_efisien is not None else 0
        durasi_rata = df_master['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_master.columns else 0
        kec_tertinggi = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else "N/A"
        tonase_kec = df_kec.iloc[0]['Total_Tonase']/1000 if not df_kec.empty else 0

        executive_summary = f"""
        **Ringkasan Operasional Bulan Ini:**
        - Total perjalanan (trip): **{total_trip}**
        - Total volume sampah diangkut: **{total_tonase:,.1f} Ton**
        - Armada paling aktif: **{teraktif_nopin}** ({teraktif_plat}) dengan **{teraktif_trip} trip**
        - Armada paling rendah aktivitasnya (perlu evaluasi): **{tidak_nopin}** ({tidak_plat}) hanya **{tidak_trip} trip**
        - Rata‑rata durasi pelayanan (timbang): **{durasi_rata:.1f} menit**
        - Wilayah dengan beban tertinggi: **{kec_tertinggi}** ({tonase_kec:,.1f} Ton)

        **Rekomendasi:**
        - Lakukan pengecekan terhadap armada **{tidak_nopin}** karena jumlah trip sangat rendah.
        - Optimasi rute di Kecamatan **{kec_tertinggi}** untuk mengurangi waktu tempuh.
        - Pertimbangkan penjadwalan ulang agar distribusi trip lebih merata.
        """
        st.markdown(executive_summary)

        # Tambahan: tampilkan sheet yang terlewat
        if hasil['skipped_sheets']:
            with st.expander(f"ℹ️ {len(hasil['skipped_sheets'])} sheet dilewati karena format tidak sesuai"):
                st.write(hasil['skipped_sheets'])

else:
    st.info("👆 Silakan unggah file Excel melalui sidebar untuk memulai analisis.")
    st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)
