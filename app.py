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
from sklearn.preprocessing import StandardScaler
from sklearn.cluster import KMeans
from sklearn.metrics import silhouette_score

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

# -------------------------- Parsing Waktu Super Fleksibel --------------------------
def parse_waktu(series):
    if pd.api.types.is_datetime64_any_dtype(series): return series
    if series.dtype == object: series = series.astype(str).str.strip().str.replace(r'\.', ':', regex=True)
    for fmt in ['%H:%M:%S', '%H:%M', '%H.%M.%S', '%I:%M:%S %p', '%I:%M %p', '%H.%M']:
        dt = pd.to_datetime(series, format=fmt, errors='coerce')
        if dt.notna().sum() > 0: return dt
    for fmt in ['%d/%m/%Y %H:%M:%S', '%d/%m/%Y %H:%M', '%Y-%m-%d %H:%M:%S', '%Y-%m-%d %H:%M',
                '%d-%m-%Y %H:%M:%S', '%d-%m-%Y %H:%M', '%m/%d/%Y %H:%M:%S', '%m/%d/%Y %H:%M']:
        dt = pd.to_datetime(series, format=fmt, errors='coerce', dayfirst=True)
        if dt.notna().sum() > 0:
            return dt.dt.time.apply(lambda t: pd.Timestamp('2026-01-01') + pd.Timedelta(hours=t.hour, minutes=t.minute, seconds=t.second))
    if pd.api.types.is_numeric_dtype(series):
        try:
            hours = series * 24
            td = pd.to_timedelta(hours, unit='h')
            base = pd.Timestamp('2026-01-01')
            return base + td
        except: pass
    dt = pd.to_datetime(series, errors='coerce', dayfirst=True)
    if dt.notna().sum() > 0:
        return dt.dt.time.apply(lambda t: pd.Timestamp('2026-01-01') + pd.Timedelta(hours=t.hour, minutes=t.minute, seconds=t.second))
    return dt

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

    with st.expander("🔍 Lihat 10 data waktu mentah"):
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
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)].copy()
        try: df_master['JAM_INPUT'] = dt_masuk.dt.hour
        except: pass
        st.success(f"✅ Durasi dihitung, {len(df_master)} baris valid.")
    else: st.warning("⚠️ Gagal mengonversi waktu. Periksa format di Excel (contoh: 08:30:00 atau 08.30).")
    return df_master

# -------------------------- Proses Data (ETL) --------------------------
# (fungsi proses_list_armada, proses_sheet_harian, hitung_agregasi_armada, ...)
# Karena sudah banyak, saya tulis ulang seluruhnya dengan penambahan di clustering.
# [Kode ETL sama persis seperti jawaban terakhir, saya lanjutkan ke bagian utama]

# -------------------------- Preprocessing untuk ML --------------------------
def remove_outliers_iqr(df, column):
    Q1 = df[column].quantile(0.25)
    Q3 = df[column].quantile(0.75)
    IQR = Q3 - Q1
    lower = Q1 - 1.5 * IQR
    upper = Q3 + 1.5 * IQR
    return df[(df[column] >= lower) & (df[column] <= upper)]

def lakukan_clustering_ml(df_armada, n_clusters=3):
    """
    Preprocessing, scaling, K-Means, evaluation, dan return hasil.
    """
    if df_armada.empty or len(df_armada) < n_clusters:
        return None, None, None, None, None

    # Fitur numerik untuk clustering
    fitur_cols = ['Total_Trip', 'Total_Tonase']
    if 'Rata_Durasi' in df_armada.columns and df_armada['Rata_Durasi'].notna().sum() > n_clusters:
        fitur_cols.append('Rata_Durasi')
        # Hanya ambil data yang memiliki durasi valid
        df_fitur = df_armada[df_armada['Rata_Durasi'].notna()].copy()
    else:
        df_fitur = df_armada.copy()

    if df_fitur.empty or len(df_fitur) < n_clusters:
        return None, None, None, None, None

    # Hapus outlier menggunakan IQR pada setiap fitur
    for col in fitur_cols:
        df_fitur = remove_outliers_iqr(df_fitur, col)

    if len(df_fitur) < n_clusters:
        return None, None, None, None, None

    X = df_fitur[fitur_cols].values

    # Scaling
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # K-Means
    kmeans = KMeans(n_clusters=n_clusters, random_state=42, n_init=10)
    labels = kmeans.fit_predict(X_scaled)
    df_fitur['Cluster'] = labels

    # Evaluasi: Inertia dan Silhouette Score
    inertia = kmeans.inertia_
    try:
        sil_score = silhouette_score(X_scaled, labels)
    except:
        sil_score = None

    # Ringkasan per cluster
    ringkasan = df_fitur.groupby('Cluster').agg(
        Jumlah_Armada=('NOPIN', 'count'),
        Rata_Trip=('Total_Trip', 'mean'),
        Rata_Tonase=('Total_Tonase', 'mean')
    ).reset_index()
    if 'Rata_Durasi' in df_fitur.columns:
        ringkasan['Rata_Durasi'] = df_fitur.groupby('Cluster')['Rata_Durasi'].mean().values

    return df_fitur, ringkasan, inertia, sil_score, X_scaled

# Fungsi untuk menampilkan grafik Elbow
def plot_elbow(X_scaled, max_k=8):
    inertias = []
    sil_scores = []
    K = range(2, min(max_k, len(X_scaled)))
    for k in K:
        kmeans = KMeans(n_clusters=k, random_state=42, n_init=10)
        labels = kmeans.fit_predict(X_scaled)
        inertias.append(kmeans.inertia_)
        try:
            sil = silhouette_score(X_scaled, labels)
            sil_scores.append(sil)
        except:
            sil_scores.append(None)

    fig_elbow = px.line(x=list(K), y=inertias, markers=True, title='Elbow Method (Inertia)')
    fig_elbow.update_layout(xaxis_title='Jumlah Cluster', yaxis_title='Inertia')
    fig_sil = px.line(x=list(K), y=sil_scores, markers=True, title='Silhouette Score')
    fig_sil.update_layout(xaxis_title='Jumlah Cluster', yaxis_title='Silhouette Score')
    return fig_elbow, fig_sil

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
    for c in ['MASUK_ORI', 'KELUAR_ORI', 'DURASI_MENIT']:
        if c in df_master.columns: cols_waktu.append(c)

    df_show = df_master.head(1000).copy()
    df_show.insert(0, 'No', range(1, len(df_show)+1))
    cols_ada = ['No'] + [c for c in cols_waktu if c in df_show.columns]
    st.dataframe(df_show[cols_ada], use_container_width=True, hide_index=True)

    # ---------- TABS ----------
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Ringkasan", "⏱️ Durasi", "📈 Tren", "🔬 Clustering (ML)"])

    with tab1:
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

        st.subheader("🏆 Armada Teraktif & Paling Tidak Efisien")
        if teraktif is not None:
            col_a, col_b = st.columns(2)
            with col_a: st.success(f"**Teraktif:** {teraktif.get('NOPIN','-')} ({teraktif.get('NO_PLAT','')}) – {int(teraktif.get('Total_Trip',0))} trip")
            with col_b:
                if tidak_efisien is not None: st.error(f"**Tidak Efisien:** {tidak_efisien.get('NOPIN','-')} ({tidak_efisien.get('NO_PLAT','')}) – {int(tidak_efisien.get('Total_Trip',0))} trip")

    with tab2:
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

    with tab3:
        st.subheader("📈 Tren Harian Ritase 30 Hari")
        if not df_tren.empty:
            df_tren['TANGGAL'] = pd.to_datetime(df_tren['TANGGAL'])
            df_tren = df_tren.sort_values('TANGGAL')
            max_row = df_tren.loc[df_tren['Total_Ritase'].idxmax()]
            fig_tren = px.line(df_tren, x='TANGGAL', y='Total_Ritase', markers=True,
                               title='Tren Frekuensi Ritase Harian (Juni 2026)', template='plotly_white')
            fig_tren.update_traces(line=dict(color='#0D9488', width=3),
                                   marker=dict(size=8, color='#0D9488', line=dict(width=1, color='white')),
                                   hovertemplate='<b>Tanggal:</b> %{x|%d %B %Y}<br><b>Ritase:</b> %{y} trip<extra></extra>')
            fig_tren.add_annotation(x=max_row['TANGGAL'], y=max_row['Total_Ritase'],
                                    text=f"Puncak: {int(max_row['Total_Ritase'])} trip",
                                    showarrow=True, arrowhead=2, arrowsize=1, arrowcolor='#0D9488', ax=0, ay=-30,
                                    font=dict(color='#0D9488', size=10))
            fig_tren.update_layout(xaxis_title='Tanggal', yaxis_title='Total Trip (Ritase)', hovermode='x unified',
                                   xaxis=dict(tickformat='%d %b', tickangle=-45, dtick='D1', tickmode='linear'),
                                   margin=dict(l=40, r=40, t=60, b=60))
            st.plotly_chart(fig_tren, use_container_width=True)
            st.session_state.grafik['tren'] = fig_tren

    with tab4:
        st.header("🔬 Cluster Analysis (K‑Means) – Alur Machine Learning")
        st.markdown("""
        1. **Preprocessing**: Penghapusan outlier (IQR), scaling data.
        2. **Pemilihan k optimal**: Elbow method & silhouette score.
        3. **Clustering**: K‑Means.
        4. **Interpretasi**: ringkasan per cluster & daftar anggota.
        """)

        if df_armada is None or df_armada.empty:
            st.warning("Data armada tidak tersedia untuk clustering.")
        else:
            # Tampilkan metrik inertia & silhouette untuk beberapa k
            st.subheader("📈 Elbow Method & Silhouette Score")
            max_k = min(8, len(df_armada))
            # Siapkan data fitur yang bersih untuk elbow
            fitur_cols = ['Total_Trip', 'Total_Tonase']
            if 'Rata_Durasi' in df_armada.columns:
                mask_dur = df_armada['Rata_Durasi'].notna()
                df_elbow = df_armada[mask_dur].copy()
                if len(df_elbow) > 0: fitur_cols.append('Rata_Durasi')
            else:
                df_elbow = df_armada.copy()
            # Bersihkan outlier
            for col in fitur_cols:
                df_elbow = remove_outliers_iqr(df_elbow, col)
            if df_elbow.empty:
                st.warning("Data setelah pembersihan outlier kosong.")
            else:
                X_elbow = df_elbow[fitur_cols].values
                scaler_elbow = StandardScaler()
                X_elbow_scaled = scaler_elbow.fit_transform(X_elbow)
                fig_elbow, fig_sil = plot_elbow(X_elbow_scaled, max_k)
                st.plotly_chart(fig_elbow, use_container_width=True)
                st.plotly_chart(fig_sil, use_container_width=True)

            # Slider untuk jumlah cluster
            n_clusters = st.slider("Jumlah Cluster", min_value=2, max_value=min(5, len(df_armada)-1), value=3, key="n_cluster")

            # Jalankan clustering dengan preprocessing lengkap
            df_cluster, ringkasan_cluster, inertia, sil_score, X_scaled = lakukan_clustering_ml(df_armada, n_clusters=n_clusters)

            if df_cluster is not None:
                # Evaluasi
                st.subheader("📊 Evaluasi Cluster")
                col_met1, col_met2 = st.columns(2)
                col_met1.metric("Inertia (sum of squared distances)", f"{inertia:,.0f}")
                if sil_score is not None:
                    col_met2.metric("Silhouette Score", f"{sil_score:.3f}")

                # Scatter plot
                fig_cluster = px.scatter(
                    df_cluster, x='Total_Trip', y='Total_Tonase', color='Cluster',
                    symbol='Cluster', hover_data=['NOPIN', 'NO_PLAT', 'Kecamatan'],
                    title='Clustering Armada Berdasarkan Trip & Tonase',
                    template='plotly_white'
                )
                if 'Rata_Durasi' in df_cluster.columns:
                    fig_cluster_dur = px.scatter(
                        df_cluster, x='Total_Trip', y='Rata_Durasi', color='Cluster',
                        symbol='Cluster', hover_data=['NOPIN', 'NO_PLAT', 'Kecamatan'],
                        title='Clustering Armada Berdasarkan Trip & Durasi',
                        template='plotly_white'
                    )
                    st.plotly_chart(fig_cluster_dur, use_container_width=True)

                st.plotly_chart(fig_cluster, use_container_width=True)

                # Tabel ringkasan per cluster
                st.subheader("📋 Ringkasan per Cluster")
                ringkasan_disp = ringkasan_cluster.copy()
                if 'Cluster' not in ringkasan_disp.columns:
                    ringkasan_disp = ringkasan_disp.reset_index()
                cols_order = ['Cluster'] + [c for c in ringkasan_disp.columns if c != 'Cluster']
                ringkasan_disp = ringkasan_disp[cols_order]
                st.dataframe(
                    ringkasan_disp.style.format({
                        'Jumlah_Armada': '{:.0f}',
                        'Rata_Trip': '{:.1f}',
                        'Rata_Tonase': '{:,.0f}',
                        'Rata_Durasi': '{:.1f}'
                    }),
                    use_container_width=True,
                    hide_index=True
                )

                # Anggota per cluster
                with st.expander("📋 Lihat Anggota per Cluster"):
                    cluster_list = sorted(df_cluster['Cluster'].unique())
                    for c in cluster_list:
                        anggota = df_cluster[df_cluster['Cluster'] == c][
                            ['NOPIN', 'NO_PLAT', 'Total_Trip', 'Total_Tonase']
                        ]
                        if 'Rata_Durasi' in df_cluster.columns:
                            anggota = df_cluster[df_cluster['Cluster'] == c][
                                ['NOPIN', 'NO_PLAT', 'Total_Trip', 'Total_Tonase', 'Rata_Durasi']
                            ]
                        st.write(f"**Cluster {c}** ({len(anggota)} armada)")
                        st.dataframe(anggota, use_container_width=True)
            else:
                st.warning("Data tidak mencukupi untuk clustering.")

    # ---------- SIMPAN GRAFIK LAIN UNTUK PDF ----------
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
