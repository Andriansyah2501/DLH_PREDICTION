import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO

# ======================== FUNGSI BANTU ========================
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
        except:
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
            # Cari kolom kecamatan – lebih fleksibel
            col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI KECAMATAN', 'LOKASI', 'KEC'])
            col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
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
        except:
            skipped.append(sheet)
            continue

        # Cari header
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
        except:
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

        # Sinkronisasi dengan master
        if ref_df is not None:
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col in df_hari.columns:
                    df_hari.drop(columns=[col], inplace=True)
            kolom_ref = ['NOPIN', 'NO_PLAT'] + [c for c in ['Kecamatan', 'MERK', 'TYPE'] if c in ref_df.columns]
            df_hari = df_hari.merge(ref_df[kolom_ref], on='NOPIN', how='left')
        else:
            for col in ['NO_PLAT', 'Kecamatan', 'MERK', 'TYPE']:
                if col not in df_hari.columns:
                    df_hari[col] = ''

        # Isi kecamatan kosong
        if 'Kecamatan' not in df_hari.columns:
            df_hari['Kecamatan'] = 'Tidak Diketahui'
        else:
            df_hari['Kecamatan'] = df_hari['Kecamatan'].fillna('Tidak Diketahui').replace('', 'Tidak Diketahui')

        # Tambah TANGGAL
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        df_hari = df_hari.loc[:, ~df_hari.columns.duplicated()]
        cleaned[sheet] = df_hari

    if not cleaned:
        return None

    df_master = pd.concat(cleaned.values(), ignore_index=True, sort=False)

    # Konversi numerik tonase
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # Waktu
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

    # Agregasi per armada
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

    # Ekstrem
    teraktif = df_armada.iloc[0] if not df_armada.empty else None
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if not df_armada.empty and (df_armada['Total_Trip'] > 0).any() else None

    # Waktu per jenis
    df_waktu = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE', dropna=False)['DURASI_MENIT'].mean().reset_index()
        df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']

    # Per Kecamatan (untuk ringkasan & grafik)
    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan', dropna=False).agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum'),
            Jumlah_Armada=('NOPIN', 'nunique')
        ).reset_index().sort_values('Total_Ritase', ascending=False)

    # Tren
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

# ======================== SESSION STATE ========================
if "hasil" not in st.session_state:
    st.session_state.hasil = None

# ======================== ANTARMUKA ========================
st.set_page_config(page_title="Dashboard DLH per Kecamatan", page_icon="🚛", layout="wide")
st.title("🚛 Dashboard Armada DLH – Analisis per Kecamatan")
st.markdown("Unggah file Excel (List Armada + 30 sheet harian). Data akan diolah per **Kecamatan** berdasarkan **NOPIN** dan **Plat Nomor**.")

with st.sidebar:
    uploaded_file = st.file_uploader("📂 Unggah file Excel", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if not sheets:
                    st.error("File tidak memiliki sheet yang valid.")
                else:
                    hasil = proses_data_armada(sheets)
                    if hasil is None:
                        st.error("Gagal memproses data.")
                    else:
                        st.session_state.hasil = hasil
                        st.success(f"✅ {hasil['cleaned_count']} sheet berhasil digabung.")
                        st.balloons()

if st.session_state.hasil is not None:
    data = st.session_state.hasil
    df = data['df_master']
    col_netto = data['col_netto']
    df_kec = data['df_kec']

    # ============ PILIHAN KECAMATAN ============
    st.sidebar.markdown("---")
    st.sidebar.header("📍 Pilih Kecamatan")
    if 'Kecamatan' in df.columns:
        daftar_kec = sorted(df['Kecamatan'].unique().tolist())
        kec_terpilih = st.sidebar.selectbox("Kecamatan", daftar_kec)
    else:
        kec_terpilih = None

    # Filter data untuk kecamatan terpilih
    if kec_terpilih:
        df_kec_filter = df[df['Kecamatan'] == kec_terpilih]
    else:
        df_kec_filter = df

    # ============ METRIK KECAMATAN ============
    st.subheader(f"📌 Data Kecamatan: **{kec_terpilih}**")
    total_trip_kec = len(df_kec_filter)
    total_armada_kec = df_kec_filter['NOPIN'].nunique()
    total_tonase_kec = df_kec_filter[col_netto].sum() / 1000 if col_netto else 0
    durasi_rata_kec = df_kec_filter['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_kec_filter.columns else 0

    colA, colB, colC, colD = st.columns(4)
    colA.metric("Total Trip", total_trip_kec)
    colB.metric("Armada Aktif", total_armada_kec)
    colC.metric("Total Tonase (Ton)", f"{total_tonase_kec:,.1f}")
    colD.metric("Rata² Durasi (menit)", f"{durasi_rata_kec:.1f}" if durasi_rata_kec else "-")

    # ============ DAFTAR ARMADA PER KECAMATAN ============
    st.subheader("🚚 Daftar Armada yang Bertugas")
    # Agregasi per armada dalam kecamatan ini
    if col_netto:
        armada_kec = df_kec_filter.groupby(['NOPIN', 'NO_PLAT'], dropna=False).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)
        st.dataframe(armada_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True)

        # Grafik top armada di kecamatan ini
        fig_top_kec = px.bar(armada_kec.head(15), x='NOPIN', y='Total_Trip',
                             color='Total_Trip', color_continuous_scale='Blues',
                             title=f'15 Armada Teraktif di {kec_terpilih}')
        st.plotly_chart(fig_top_kec, use_container_width=True)

    # ============ RINGKASAN SEMUA KECAMATAN ============
    st.markdown("---")
    st.subheader("📊 Ringkasan Seluruh Kecamatan")
    if not df_kec.empty:
        st.dataframe(df_kec.style.format({'Total_Tonase': '{:,.0f}'}), use_container_width=True)
        fig_kec = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', color='Total_Tonase',
                         color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
        fig_kec.update_layout(xaxis_tickangle=-45)
        st.plotly_chart(fig_kec, use_container_width=True)

    # ============ TREN HARIAN (GLOBAL) ============
    if not data['df_tren'].empty:
        st.subheader("📈 Tren Harian (Semua Kecamatan)")
        fig_tren = px.line(data['df_tren'], x='TANGGAL', y='Total_Ritase',
                           title='Tren Ritase Harian', markers=True)
        st.plotly_chart(fig_tren, use_container_width=True)

    # ============ UNDUHAN ============
    st.subheader("📥 Unduh Data")
    @st.cache_data
    def to_excel(dataframe):
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as w:
            dataframe.to_excel(w, index=False)
        return output.getvalue()

    col_u1, col_u2 = st.columns(2)
    with col_u1:
        st.download_button("📊 Master Data (Excel)", to_excel(df), "Master_Data.xlsx")
        st.download_button("📊 Statistik Armada (Excel)", to_excel(data['df_armada']), "Statistik_Armada.xlsx")
    with col_u2:
        st.download_button("📊 Laporan per Kecamatan (Excel)", to_excel(df_kec), "Kecamatan.xlsx")
        if not data['df_waktu_jenis'].empty:
            st.download_button("⏱️ Waktu per Jenis (Excel)", to_excel(data['df_waktu_jenis']), "Waktu_per_Jenis.xlsx")

else:
    st.info("👆 Unggah file Excel dan klik **Proses Data**.")
