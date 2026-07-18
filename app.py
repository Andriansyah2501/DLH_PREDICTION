import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import plotly.io as pio
import requests
import os
from io import BytesIO
from datetime import datetime

# ==================== KONFIGURASI HALAMAN ====================
st.set_page_config(
    page_title="Dashboard DLH Armada – Full Analysis",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Kustom CSS untuk tampilan profesional
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 16px; padding: 24px; color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15); text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover { transform: translateY(-5px); }
    .metric-value { font-size: 2.6rem; font-weight: 800; margin: 8px 0; }
    .metric-label { font-size: 1rem; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; }
    .stButton button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 12px 28px; border-radius: 8px; font-weight: 600; transition: 0.3s; }
    .stButton button:hover { transform: scale(1.02); box-shadow: 0 4px 15px rgba(102,126,234,0.4); }
    .download-section { background: #f9fafb; padding: 20px; border-radius: 12px; margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# ==================== API DEEPSEEK (opsional) ====================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def laporan_ai(statistik: str) -> str:
    """Menghasilkan laporan analisis menggunakan DeepSeek AI."""
    if not DEEPSEEK_API_KEY:
        return "⚠️ API Key DeepSeek belum diatur. Laporan AI tidak dapat dibuat."
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
    prompt = f"""
Anda adalah analis data armada DLH. Buat laporan singkat (3 paragraf) dalam bahasa Indonesia dari statistik berikut:
{statistik}
Sertakan:
1. Gambaran umum performa armada.
2. Armada teraktif dan paling tidak efisien beserta dugaan penyebab.
3. Rekomendasi perbaikan (penjadwalan, rute, perawatan).
    """
    try:
        resp = requests.post(DEEPSEEK_URL, headers=headers, json={
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 500
        }, timeout=30)
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ Gagal membuat laporan: {str(e)}"

# ==================== FUNGSI BANTU ====================
def cari_kolom(df_columns, keywords):
    """Mencari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in df_columns:
        col_upper = col.upper()
        if any(kw in col_upper for kw in keywords):
            return col
    return None

@st.cache_data(show_spinner="Membaca file...")
def load_all_sheets(uploaded_file):
    """Membaca semua sheet dari file Excel yang diunggah."""
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except Exception:
            pass
    return sheets

def process_dlh_data(sheets_dict, uploaded_file):
    """
    Alur lengkap sesuai notebook dlh.ipynb:
    1. Baca List Armada (header otomatis).
    2. Proses 30 sheet harian (nama digit).
    3. Bersihkan & sinkronkan dengan master.
    4. Gabungkan jadi Master Data.
    5. Hitung agregasi untuk tiap poin analisis.
    Mengembalikan (df_master, semua_dataframe_hasil)
    """
    # === 1. Cari sheet List Armada ===
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)
    if armada_sheet is None:
        st.error("❌ Sheet 'List Armada' tidak ditemukan dalam file.")
        return None, None

    # === 2. Baca referensi armada (header=1 seperti notebook) ===
    df_arm_raw = sheets_dict[armada_sheet]
    header_arm = 1  # notebook pakai header=1
    xls = pd.ExcelFile(uploaded_file)
    df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_arm)
    df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

    # Cari kolom penting di referensi
    col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'NO. PINTU', 'PINTU'])
    col_plat = cari_kolom(df_ref.columns, ['NO.PLAT', 'PLAT', 'NO PLAT'])
    col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI KECAMATAN', 'LOKASI'])
    col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
    col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])

    if not col_nopin or not col_plat:
        st.error(f"Kolom NOPIN / Plat tidak ditemukan di sheet '{armada_sheet}'.")
        return None, None

    # Buat dictionary referensi (NOPIN → info armada)
    df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
    df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
    ref_dict = {}
    for _, row in df_ref.iterrows():
        nopin = row['NOPIN']
        entry = {'NO.PLAT': row['NO.PLAT']}
        if col_kec:
            entry['Kecamatan'] = str(row[col_kec]).strip() if pd.notna(row[col_kec]) else ''
        if col_merk:
            entry['MERK'] = str(row[col_merk]).strip() if pd.notna(row[col_merk]) else ''
        if col_type:
            entry['TYPE'] = str(row[col_type]).strip() if pd.notna(row[col_type]) else ''
        ref_dict[nopin] = entry

    # === 3. Proses sheet harian (1-30) ===
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]

    cleaned_sheets = {}
    total_baris = 0
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, hari in enumerate(daily_sheets):
        status_text.text(f"Memproses sheet {hari} ({i+1}/{len(daily_sheets)})")
        progress_bar.progress((i+1)/len(daily_sheets))

        df_raw = sheets_dict[hari]
        # Deteksi header – cari baris dengan "PINTU", "PLAT MOBIL", atau "NOPIN"
        header_harian = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str or 'NOPIN' in row_str:
                header_harian = idx
                break
        if header_harian is None:
            continue

        df_hari = pd.read_excel(uploaded_file, sheet_name=hari, header=header_harian)
        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]

        col_nopin_h = cari_kolom(df_hari.columns, ['NOPIN', 'NO. PINTU', 'PINTU'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT', 'NO PLAT', 'NO.PLAT'])
        if not col_nopin_h or not col_plat_h:
            continue

        df_hari.rename(columns={col_nopin_h: 'NOPIN', col_plat_h: 'NO_PLAT'}, inplace=True)

        # Pembersihan seperti notebook
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

        # Sinkronisasi dengan master
        def sinkron(row):
            nopin = row['NOPIN']
            if nopin in ref_dict:
                row['NO_PLAT'] = ref_dict[nopin]['NO.PLAT']
                for key in ['Kecamatan', 'MERK', 'TYPE']:
                    if key in ref_dict[nopin]:
                        row[key] = ref_dict[nopin][key]
            return row

        df_hari = df_hari.apply(sinkron, axis=1)
        df_hari['TANGGAL'] = f"2026-06-{int(hari):02d}" if hari.isdigit() else hari

        cleaned_sheets[hari] = df_hari
        total_baris += len(df_hari)

    progress_bar.empty()
    status_text.empty()

    if not cleaned_sheets:
        st.error("Tidak ada sheet harian yang berhasil diproses.")
        return None, None

    # Gabungkan semua data harian
    df_master = pd.concat(cleaned_sheets.values(), ignore_index=True)

    # Konversi kolom numerik (NETTO/GROSS/TARE) ke float
    for col in df_master.columns:
        if 'NETTO' in col or 'GROSS' in col or 'TARE' in col or 'BERAT' in col:
            df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)

    # Tentukan kolom Netto utama (untuk perhitungan tonase)
    col_netto = cari_kolom(df_master.columns, ['NETTO'])
    if not col_netto:
        col_netto = cari_kolom(df_master.columns, ['TOTAL', 'JUMLAH'])

    # === 4. Analisis Waktu (jika ada kolom jam) ===
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['WAKTU_MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['WAKTU_KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['WAKTU_KELUAR_DT'] - df_master['WAKTU_MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        try:
            df_master['JAM_INPUT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce').dt.hour
        except Exception:
            pass

    # === 5. Agregasi sesuai Poin 2–6 ===
    # Poin 2: Distribusi per Kecamatan
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kecamatan = df_master.groupby('Kecamatan').agg(
            Jumlah_Armada_Aktif=('NOPIN', 'nunique'),
            Total_Ritase_Trip=('NOPIN', 'count'),
            Total_Tonase_Bersih=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Ritase_Trip', ascending=False)
    else:
        df_kecamatan = pd.DataFrame()

    # Poin 3: Performa per Armada
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        if 'Kecamatan' in df_master.columns:
            group_cols.append('Kecamatan')
        if 'TYPE' in df_master.columns:
            group_cols.append('TYPE')
        if 'MERK' in df_master.columns:
            group_cols.append('MERK')
        df_armada_performa = df_master.groupby(group_cols).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Netto_Kg=(col_netto, 'sum')
        ).reset_index()
        df_armada_performa['Total_Netto_Ton'] = (df_armada_performa['Total_Netto_Kg'] / 1000).round(2)
        df_armada_performa = df_armada_performa.sort_values('Total_Netto_Kg', ascending=False)
    else:
        df_armada_performa = pd.DataFrame()

    # Poin 4: Tren Harian
    if col_netto:
        df_tren_harian = df_master.groupby('TANGGAL').agg(
            Total_Ritase_Trip=('NOPIN', 'count'),
            Total_Tonase_Kg=(col_netto, 'sum'),
            Rata_Rata_Muatan_Per_Trip=(col_netto, 'mean'),
            Jumlah_Armada_Operasional=('NOPIN', 'nunique')
        ).reset_index()
        df_tren_harian['Total_Tonase_Ton'] = (df_tren_harian['Total_Tonase_Kg'] / 1000).round(2)
        df_tren_harian = df_tren_harian.sort_values('TANGGAL')
    else:
        df_tren_harian = pd.DataFrame()

    # Poin 5: Efisiensi Waktu per Kecamatan & Jam Sibuk
    df_durasi_kec = pd.DataFrame()
    df_jam_puncak = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'Kecamatan' in df_master.columns:
        df_waktu_valid = df_master.dropna(subset=['DURASI_MENIT'])
        df_durasi_kec = df_waktu_valid.groupby('Kecamatan').agg(
            Total_Aktivitas_Trip=('NOPIN', 'count'),
            Rata_Durasi_Pelayanan_Menit=('DURASI_MENIT', 'mean')
        ).reset_index().round(2).sort_values('Rata_Durasi_Pelayanan_Menit', ascending=False)
        if 'JAM_INPUT' in df_waktu_valid.columns:
            df_jam_puncak = df_waktu_valid.groupby('JAM_INPUT').size().reset_index(name='Jumlah_Truk_Masuk').sort_values('JAM_INPUT')

    st.success(f"✅ Data siap: {total_baris} baris dari {len(cleaned_sheets)} sheet harian.")
    return df_master, (df_kecamatan, df_armada_performa, df_tren_harian, df_durasi_kec, df_jam_puncak)

# ==================== SESSION STATE ====================
if "data_processed" not in st.session_state:
    st.session_state.data_processed = False
    st.session_state.df_master = None
    st.session_state.dataframes = None
    st.session_state.figures = {}
    st.session_state.laporan_teks = ""

# ==================== APLIKASI UTAMA ====================
def main():
    st.title("🚛 Dashboard Analitik DLH – Armada Sampah (Juni 2026)")
    st.markdown("Unggah file Excel dengan sheet **List Armada** dan 30 sheet harian. Dashboard ini menjalankan **seluruh analisis notebook DLH** secara otomatis.")

    with st.sidebar:
        st.markdown("## 📂 Unggah File Excel")
        uploaded_file = st.file_uploader("Pilih file .xlsx atau .xls", type=["xlsx", "xls"])
        if uploaded_file:
            st.success("File siap diproses")

    if uploaded_file is not None:
        if st.sidebar.button("🚀 Proses Data (Jalankan Semua)", use_container_width=True):
            with st.spinner("Menjalankan alur notebook DLH..."):
                sheets_dict = load_all_sheets(uploaded_file)
                if not sheets_dict:
                    st.error("File tidak memiliki sheet yang bisa dibaca.")
                    return
                result = process_dlh_data(sheets_dict, uploaded_file)
                if result[0] is not None:
                    st.session_state.df_master, st.session_state.dataframes = result[0], result[1]
                    st.session_state.data_processed = True
                    st.balloons()

        if st.session_state.data_processed:
            df_master = st.session_state.df_master.copy()
            (df_kecamatan, df_armada_performa, df_tren_harian,
             df_durasi_kec, df_jam_puncak) = st.session_state.dataframes

            # ========== SIDEBAR FILTER ==========
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
            # Apply filter
            df_filtered = df_master.copy()
            if kec_terpilih != 'Semua':
                df_filtered = df_filtered[df_filtered['Kecamatan'] == kec_terpilih]
            if 'rentang' in locals() and len(rentang) == 2:
                df_filtered = df_filtered[(pd.to_datetime(df_filtered['TANGGAL']) >= pd.Timestamp(rentang[0])) &
                                          (pd.to_datetime(df_filtered['TANGGAL']) <= pd.Timestamp(rentang[1]))]

            # ========== METRIK UTAMA ==========
            total_trip = len(df_filtered)
            total_armada = df_filtered['NOPIN'].nunique()
            col_netto = cari_kolom(df_filtered.columns, ['NETTO']) or cari_kolom(df_filtered.columns, ['TOTAL'])
            total_tonase = df_filtered[col_netto].sum() / 1000 if col_netto else 0
            rata_durasi = df_filtered['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_filtered.columns else 0

            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{total_armada}</div></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase (Ton)</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Rata² Durasi (menit)</div><div class="metric-value">{rata_durasi:.1f}</div></div>', unsafe_allow_html=True)

            st.markdown("---")

            # ========== VISUALISASI ==========
            # 1. Tren Ritase Harian
            if not df_tren_harian.empty:
                fig_tren = px.line(df_tren_harian, x='TANGGAL', y='Total_Ritase_Trip',
                                   title='📈 Tren Frekuensi Ritase Harian (Juni 2026)', markers=True)
                fig_tren.update_traces(line_color='#0D9488', line_width=3, marker_size=8)
                fig_tren.update_layout(template='plotly_white')
                st.plotly_chart(fig_tren, use_container_width=True)
                st.session_state.figures['tren'] = fig_tren

            # 2. Distribusi Tonase per Kecamatan
            if not df_kecamatan.empty:
                fig_kec = px.bar(df_kecamatan, x='Kecamatan', y='Total_Tonase_Bersih',
                                 color='Total_Tonase_Bersih', color_continuous_scale='Viridis',
                                 title='📦 Total Volume Sampah per Kecamatan (Kg)')
                fig_kec.update_layout(xaxis_tickangle=-45, template='plotly_white')
                st.plotly_chart(fig_kec, use_container_width=True)
                st.session_state.figures['kec'] = fig_kec

            # 3. Pola Kedatangan per Jam (jika tersedia)
            if not df_jam_puncak.empty:
                fig_jam = px.area(df_jam_puncak, x='JAM_INPUT', y='Jumlah_Truk_Masuk',
                                  title='⏰ Pola Kedatangan Truk per Jam (Jembatan Timbang)')
                fig_jam.update_xaxes(dtick=1)
                fig_jam.update_traces(fillcolor='rgba(30, 58, 138, 0.4)', line_color='#1E3A8A')
                fig_jam.update_layout(template='plotly_white')
                st.plotly_chart(fig_jam, use_container_width=True)
                st.session_state.figures['jam'] = fig_jam

            # 4. Top 10 Armada Teraktif
            if not df_armada_performa.empty:
                top10 = df_armada_performa.head(10)
                fig_top = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip',
                                 color_continuous_scale='OrRd',
                                 title='🏆 10 Armada dengan Trip Terbanyak')
                fig_top.update_layout(xaxis_tickangle=-45, template='plotly_white')
                st.plotly_chart(fig_top, use_container_width=True)
                st.session_state.figures['top'] = fig_top

            # ========== RINGKASAN EKSEKUTIF (POIN 7) ==========
            st.markdown("---")
            st.header("📝 Ringkasan Eksekutif & Rekomendasi Strategis")

            # Ambil data untuk ringkasan
            if not df_kecamatan.empty:
                kec_tertinggi = df_kecamatan.iloc[0]['Kecamatan']
                tonase_tertinggi = df_kecamatan.iloc[0]['Total_Tonase_Bersih'] / 1000
            else:
                kec_tertinggi = "N/A"
                tonase_tertinggi = 0
            if not df_tren_harian.empty:
                hari_puncak_ritase = df_tren_harian.loc[df_tren_harian['Total_Ritase_Trip'].idxmax(), 'TANGGAL']
                trip_puncak = int(df_tren_harian['Total_Ritase_Trip'].max())
            else:
                hari_puncak_ritase = "N/A"
                trip_puncak = 0

            col_left, col_right = st.columns(2)
            with col_left:
                st.markdown(f"""
                **Ringkasan Operasional Juni 2026:**
                - Total armada aktif: **{total_armada} unit**
                - Total trip (ritase): **{total_trip}**
                - Total volume sampah: **{total_tonase:,.1f} Ton**
                - Wilayah beban tertinggi: **{kec_tertinggi}** ({tonase_tertinggi:,.1f} Ton)
                - Hari tersibuk: **{hari_puncak_ritase}** ({trip_puncak} trip)
                """)
            with col_right:
                st.markdown(f"""
                **Rekomendasi Strategis:**
                - **Alokasi Armada:** Prioritaskan penambahan unit di Kecamatan **{kec_tertinggi}** untuk mencegah penumpukan sampah.
                - **Manajemen Waktu:** Atur shift kedatangan agar tidak menumpuk di jam sibuk.
                - **Pemeliharaan:** Unit dengan trip rendah perlu evaluasi mekanis dan penjadwalan ulang.
                """)

            # ========== LAPORAN AI (opsional) ==========
            st.markdown("---")
            st.subheader("🤖 Laporan Cerdas (DeepSeek AI)")
            if st.button("🔮 Buat Laporan AI"):
                statistik_teks = f"""
Total trip: {total_trip}
Armada aktif: {total_armada}
Total tonase: {total_tonase:,.1f} Ton
Kecamatan tertinggi: {kec_tertinggi} ({tonase_tertinggi:,.1f} Ton)
Hari tersibuk: {hari_puncak_ritase} ({trip_puncak} trip)
Rata-rata durasi: {rata_durasi:.1f} menit
"""
                with st.spinner("Menghubungi DeepSeek..."):
                    st.session_state.laporan_teks = laporan_ai(statistik_teks)
                st.markdown("### 📄 Laporan AI")
                st.write(st.session_state.laporan_teks)
            else:
                st.info("Klik tombol di atas untuk menghasilkan laporan otomatis dari DeepSeek AI (API Key diperlukan).")

            # ========== UNDUH SEMUA FILE (SEPERTI NOTEBOOK) ==========
            st.markdown("---")
            st.header("📥 Unduh Semua Hasil Analisis")
            st.markdown("Semua file di bawah ini sesuai dengan output notebook `dlh.ipynb`.")

            @st.cache_data
            def to_excel(df, sheet_name='Sheet1'):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name=sheet_name)
                return output.getvalue()

            # Baris pertama tombol unduh
            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                st.download_button(
                    label="📊 Master Data Gabungan (Poin 1)",
                    data=to_excel(df_master, 'Master Data'),
                    file_name="Master_Data_Gabungan_Juni_2026.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                if not df_kecamatan.empty:
                    st.download_button(
                        label="📊 Laporan per Kecamatan (Poin 2)",
                        data=to_excel(df_kecamatan, 'Per Kecamatan'),
                        file_name="Laporan_Analisis_Kecamatan_Poin2.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                if not df_armada_performa.empty:
                    st.download_button(
                        label="📊 Performa Armada (Poin 3)",
                        data=to_excel(df_armada_performa, 'Performa Armada'),
                        file_name="Laporan_Performa_Tonase_Armada_Poin3.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
            with col_d2:
                if not df_tren_harian.empty:
                    st.download_button(
                        label="📈 Tren Harian (Poin 4)",
                        data=to_excel(df_tren_harian, 'Tren Harian'),
                        file_name="Laporan_Tren_Harian_Poin4.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                if not df_durasi_kec.empty:
                    st.download_button(
                        label="⏱️ Efisiensi Waktu (Poin 5)",
                        data=to_excel(df_durasi_kec, 'Efisiensi Waktu'),
                        file_name="Laporan_Efisiensi_Waktu_Poin5.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                # Ringkasan eksekutif dalam bentuk TXT
                ringkasan_teks = f"""LAPORAN KESIMPULAN EKSEKUTIF - REKAP TONASE DLH BATAM (JUNI 2026)
Total Armada Aktif: {total_armada} Unit
Total Ritase: {total_trip} Trip
Total Tonase Bersih: {total_tonase:,.1f} Ton
Kecamatan Tertinggi: {kec_tertinggi} ({tonase_tertinggi:,.1f} Ton)
Hari Tersibuk: {hari_puncak_ritase} ({trip_puncak} Trip)
"""
                st.download_button(
                    label="📝 Ringkasan Eksekutif (TXT)",
                    data=ringkasan_teks.encode('utf-8'),
                    file_name="Ringkasan_Eksekutif_Poin7.txt",
                    mime="text/plain",
                    use_container_width=True
                )
            with col_d3:
                # Pilih grafik untuk diunduh sebagai PNG
                pilihan_grafik = st.selectbox(
                    "Pilih grafik untuk diunduh (PNG)",
                    ["Tren Ritase", "Tonase per Kecamatan", "Jam Sibuk", "Top Armada"],
                    key="pilih_grafik"
                )
                fig_map = {
                    'Tren Ritase': 'tren',
                    'Tonase per Kecamatan': 'kec',
                    'Jam Sibuk': 'jam',
                    'Top Armada': 'top'
                }
                fig_key = fig_map.get(pilihan_grafik)
                if fig_key and fig_key in st.session_state.figures:
                    try:
                        img_bytes = pio.to_image(st.session_state.figures[fig_key], format='png', scale=2)
                        st.download_button(
                            label=f"📸 Unduh Grafik {pilihan_grafik} (PNG)",
                            data=img_bytes,
                            file_name=f"Grafik_{pilihan_grafik.replace(' ', '_')}.png",
                            mime="image/png",
                            use_container_width=True
                        )
                    except Exception as e:
                        st.warning(f"Gagal mengunduh grafik: {e}. Pastikan kaleido terinstal (pip install kaleido).")
                # Laporan AI
                if st.session_state.laporan_teks:
                    st.download_button(
                        label="🤖 Laporan AI (TXT)",
                        data=st.session_state.laporan_teks.encode('utf-8'),
                        file_name="Laporan_AI_DeepSeek.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                else:
                    st.markdown("*(Buat laporan AI terlebih dahulu)*")

            # ========== TAMPILKAN DATA MENTAH (opsional) ==========
            with st.expander("🔎 Lihat Master Data Mentah"):
                st.dataframe(df_master.head(100))

    else:
        st.info("👆 Silakan unggah file Excel Anda untuk memulai analisis lengkap.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
