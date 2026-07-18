import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# -------------------------- Fungsi Bantu --------------------------
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
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except Exception:
            pass
    return sheets

def proses_data_armada(sheets_dict):
    # ---- 1. Cari sheet List Armada ----
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)

    ref_df = None
    ref_dict = {}
    if armada_sheet:
        df_arm_raw = sheets_dict[armada_sheet].copy()
        header_arm = 0
        for idx, row in df_arm_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'NOPIN' in row_str or 'NO.PLAT' in row_str:
                header_arm = idx
                break
        if header_arm > 0:
            df_ref = df_arm_raw.iloc[header_arm:].reset_index(drop=True)
            df_ref.columns = df_arm_raw.iloc[header_arm].astype(str).str.strip().str.upper()
        else:
            df_ref = df_arm_raw.copy()
            df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

        col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
        col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
        if col_nopin and col_plat:
            df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
            df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
            col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
            col_merk = cari_kolom(df_ref.columns, ['MERK'])
            col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
            for _, row in df_ref.iterrows():
                nopin = row['NOPIN']
                ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
                if col_kec: ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip()
                if col_merk: ref_dict[nopin]['MERK'] = str(row[col_merk]).strip()
                if col_type: ref_dict[nopin]['TYPE'] = str(row[col_type]).strip()
            ref_df = pd.DataFrame.from_dict(ref_dict, orient='index').reset_index().rename(columns={'index': 'NOPIN'})

    # ---- 2. Proses sheet harian ----
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]
    if not daily_sheets:
        return None

    cleaned = {}
    skipped = []

    for sheet in daily_sheets:
        try:
            df_raw = sheets_dict[sheet].copy()
        except Exception:
            skipped.append(sheet)
            continue

        # Cari header (baris yang mengandung PINTU / PLAT MOBIL / NOPIN)
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

        df_hari = df_hari.reset_index(drop=True)

        # Sinkronisasi dengan master (Tugas 2)
        if ref_df is not None:
            # Hapus dulu kolom yang akan diupdate dari master (kecuali NOPIN)
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns:
                    df_hari.drop(columns=[col], inplace=True)
            kolom_ref = ['NOPIN', 'NO_PLAT'] + [c for c in ['Kecamatan', 'MERK', 'TYPE'] if c in ref_df.columns]
            df_hari = df_hari.merge(ref_df[kolom_ref], on='NOPIN', how='left')
        else:
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns:
                    df_hari[col] = ''

        # Pastikan Kecamatan terisi – jika kosong isi "Tidak Diketahui"
        if 'Kecamatan' not in df_hari.columns:
            df_hari['Kecamatan'] = 'Tidak Diketahui'
        else:
            df_hari['Kecamatan'] = df_hari['Kecamatan'].fillna('Tidak Diketahui').replace('', 'Tidak Diketahui')

        # Tambah kolom TANGGAL
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        # Hapus duplikasi kolom (jika ada)
        df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]

        cleaned[sheet] = df_hari

    if not cleaned:
        return None

    # Gabungkan semua DataFrame yang sudah bersih
    df_master = pd.concat(cleaned.values(), ignore_index=True, sort=False)

    # Konversi numerik kolom tonase
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

    # Agregasi (Tugas 3 & 4)
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

    # Rata‑rata waktu per jenis armada (Tugas 5)
    df_waktu = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']

    # Agregasi per kecamatan (tampilkan SEMUA kecamatan)
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

# -------------------------- Session State --------------------------
if "hasil" not in st.session_state:
    st.session_state.hasil = None

# -------------------------- Antarmuka Streamlit --------------------------
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
st.markdown("Unggah file Excel (berisi sheet **List Armada** dan 30 sheet harian). Seluruh analisis dijalankan otomatis.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Membaca dan memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if not sheets:
                    st.error("File tidak memiliki sheet yang bisa dibaca.")
                else:
                    hasil = proses_data_armada(sheets)
                    if hasil is None:
                        st.error("Gagal memproses data. Pastikan ada sheet harian yang valid.")
                    else:
                        st.session_state.hasil = hasil
                        st.success(f"✅ {hasil['cleaned_count']} sheet harian berhasil digabung. {len(hasil['skipped'])} sheet dilewati.")
                        st.balloons()

# Tampilkan hasil jika sudah diproses
if st.session_state.hasil is not None:
    data = st.session_state.hasil
    df = data['df_master']
    col_netto = data['col_netto']

    # Filter
    st.sidebar.markdown("---")
    st.sidebar.header("🔍 Filter Data")
    if 'Kecamatan' in df.columns:
        # Semua kecamatan termasuk "Tidak Diketahui" akan muncul
        kec_list = ['Semua'] + sorted(df['Kecamatan'].unique().tolist())
        kec_terpilih = st.sidebar.selectbox("Kecamatan", kec_list)
    else:
        kec_terpilih = 'Semua'
    if 'TANGGAL' in df.columns:
        tgl_list = sorted(df['TANGGAL'].unique())
        if len(tgl_list) > 1:
            rentang = st.sidebar.date_input("Rentang Tanggal", 
                [pd.to_datetime(tgl_list[0]), pd.to_datetime(tgl_list[-1])])
    df_filter = df.copy()
    if kec_terpilih != 'Semua':
        df_filter = df_filter[df_filter['Kecamatan'] == kec_terpilih]
    if 'rentang' in locals() and len(rentang) == 2:
        df_filter = df_filter[(pd.to_datetime(df_filter['TANGGAL']) >= pd.Timestamp(rentang[0])) & 
                              (pd.to_datetime(df_filter['TANGGAL']) <= pd.Timestamp(rentang[1]))]

    # Metrik
    total_trip = len(df_filter)
    total_armada = df_filter['NOPIN'].nunique()
    total_tonase = df_filter[col_netto].sum() / 1000 if col_netto else 0
    durasi_rata = df_filter['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_filter.columns else None

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Total Trip", total_trip)
    col2.metric("Armada Aktif", total_armada)
    col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
    col4.metric("Rata² Durasi (menit)", f"{durasi_rata:.1f}" if durasi_rata else "-")

    st.markdown("---")

    # Armada Teraktif & Tidak Efisien
    st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien")
    teraktif = data['teraktif']
    tidak_efisien = data['tidak_efisien']
    if teraktif is not None:
        col_a, col_b = st.columns(2)
        with col_a:
            st.success(f"**Teraktif:** {teraktif['NOPIN']} ({teraktif['NO_PLAT']}) – {int(teraktif['Total_Trip'])} trip")
        with col_b:
            if tidak_efisien is not None:
                st.error(f"**Tidak Efisien:** {tidak_efisien['NOPIN']} ({tidak_efisien['NO_PLAT']}) – {int(tidak_efisien['Total_Trip'])} trip")

    # Waktu per Jenis
    if not data['df_waktu_jenis'].empty:
        st.subheader("⏱️ Rata‑rata Waktu Tempuh per Jenis Armada")
        st.dataframe(data['df_waktu_jenis'].style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}))

    st.markdown("---")

    # Grafik
    st.subheader("📊 Visualisasi Data")
    if col_netto:
        tren = df_filter.groupby('TANGGAL', dropna=False).size().reset_index(name='Ritase')
        fig1 = px.line(tren, x='TANGGAL', y='Ritase', title='Tren Ritase Harian', markers=True)
        fig1.update_traces(line_color='#0D9488')
        st.plotly_chart(fig1, use_container_width=True)

        # Distribusi per Kecamatan – menampilkan SEMUA kecamatan
        if 'Kecamatan' in df_filter.columns:
            kec = df_filter.groupby('Kecamatan', dropna=False)[col_netto].sum().reset_index(name='Tonase')
            kec = kec.sort_values('Tonase', ascending=False)
            fig2 = px.bar(kec, x='Kecamatan', y='Tonase', color='Tonase',
                          color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
            fig2.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig2, use_container_width=True)

        if not data['df_armada'].empty:
            top10 = data['df_armada'].head(10)
            fig3 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip',
                          color_continuous_scale='OrRd', title='10 Armada Teraktif')
            st.plotly_chart(fig3, use_container_width=True)

        if 'JAM_INPUT' in df_filter.columns:
            jam = df_filter.dropna(subset=['JAM_INPUT']).groupby('JAM_INPUT').size().reset_index(name='Jumlah')
            fig4 = px.area(jam, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
            st.plotly_chart(fig4, use_container_width=True)

    st.markdown("---")

    # Tabel ringkasan per kecamatan (seluruhnya)
    st.subheader("📋 Tabel Ringkasan per Kecamatan")
    if not data['df_kec'].empty:
        st.dataframe(data['df_kec'].style.format({'Total_Tonase': '{:,.0f}'}))
    else:
        st.info("Data kecamatan belum tersedia.")

    st.markdown("---")

    # Laporan Ringkasan
    st.subheader("📝 Laporan Ringkasan Otomatis")
    kec_tertinggi = data['df_kec'].iloc[0]['Kecamatan'] if not data['df_kec'].empty else '-'
    laporan = f"""
**Ringkasan Operasional:**
- Total trip: {total_trip}
- Armada aktif: {total_armada} unit
- Total volume sampah: {total_tonase:,.1f} Ton
- Kecamatan dengan aktivitas tertinggi: {kec_tertinggi}
- Armada teraktif: {teraktif['NOPIN'] if teraktif is not None else '-'}
- Armada paling tidak efisien: {tidak_efisien['NOPIN'] if tidak_efisien is not None else '-'}
"""
    st.markdown(laporan)
    st.download_button("📄 Unduh Ringkasan (TXT)", laporan.encode('utf-8'), "ringkasan.txt")

    # Unduhan
    st.subheader("📥 Unduh Data Hasil Analisis")
    @st.cache_data
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as w:
            dataframe.to_excel(w, index=False)
        return output.getvalue()

    col_u1, col_u2, col_u3 = st.columns(3)
    with col_u1:
        st.download_button("📊 Master Data (Excel)", to_excel(data['df_master']), "Master_Data.xlsx")
        if not data['df_armada'].empty:
            st.download_button("📊 Statistik Armada (Excel)", to_excel(data['df_armada']), "Statistik_Armada.xlsx")
    with col_u2:
        if not data['df_kec'].empty:
            st.download_button("📊 Laporan Kecamatan (Excel)", to_excel(data['df_kec']), "Kecamatan.xlsx")
        if not data['df_tren'].empty:
            st.download_button("📈 Tren Harian (Excel)", to_excel(data['df_tren']), "Tren_Harian.xlsx")
    with col_u3:
        if not data['df_waktu_jenis'].empty:
            st.download_button("⏱️ Waktu per Jenis (Excel)", to_excel(data['df_waktu_jenis']), "Waktu_per_Jenis.xlsx")

    with st.expander("🔎 Lihat Data Mentah (200 baris pertama)"):
        st.dataframe(data['df_master'].head(200))

else:
    st.info("👆 Silakan unggah file Excel dan klik **Proses Data** untuk memulai analisis.")
    st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)
