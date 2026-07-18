import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO
from datetime import datetime

# ==================== KONFIGURASI HALAMAN ====================
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 16px; padding: 24px; color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15); text-align: center;
    }
    .metric-value { font-size: 2.6rem; font-weight: 800; margin: 8px 0; }
    .metric-label { font-size: 1rem; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; }
    .stButton button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 12px 28px; border-radius: 8px; font-weight: 600; transition: 0.3s; }
    .stButton button:hover { transform: scale(1.02); box-shadow: 0 4px 15px rgba(102,126,234,0.4); }
</style>
""", unsafe_allow_html=True)

# ==================== API DEEPSEEK (opsional) ====================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def laporan_ai(statistik: str) -> str:
    if not DEEPSEEK_API_KEY:
        return "⚠️ API Key DeepSeek belum diatur."
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
    prompt = f"""
Anda analis data armada DLH. Buat laporan singkat (3 paragraf) dalam bahasa Indonesia dari statistik berikut:
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
        return f"❌ Gagal: {str(e)}"

# ==================== FUNGSI UTAMA ====================
def cari_kolom(df_columns, keywords):
    """Mencari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in df_columns:
        col_upper = col.upper()
        if any(kw in col_upper for kw in keywords):
            return col
    return None

@st.cache_data(show_spinner="Membaca file...")
def load_all_sheets(uploaded_file):
    xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except:
            pass
    return sheets

def process_dlh_data(sheets_dict, uploaded_file):
    """
    Alur lengkap notebook dlh.ipynb.
    Mengembalikan (df_master, semua_dataframe_hasil, ringkasan_teks)
    """
    # 1. Cari sheet List Armada
    armada_sheet = None
    for name in sheets_dict.keys():
        if 'list armada' in name.lower():
            armada_sheet = name
            break
    if armada_sheet is None:
        st.error("❌ Sheet 'List Armada' tidak ditemukan.")
        return None, None, None

    # 2. Baca referensi dengan deteksi header (biasanya header=1)
    df_arm_raw = sheets_dict[armada_sheet]
    header_arm = 0
    for idx, row in df_arm_raw.iterrows():
        row_str = " ".join(row.astype(str).dropna().str.upper().values)
        if 'NOPIN' in row_str or 'NO.PLAT' in row_str:
            header_arm = idx
            break
    xls = pd.ExcelFile(uploaded_file)
    df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_arm)
    df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

    # Cari kolom vital
    col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'NO. PINTU', 'PINTU'])
    col_plat = cari_kolom(df_ref.columns, ['NO.PLAT', 'PLAT', 'NO PLAT'])
    col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI KECAMATAN', 'LOKASI'])
    col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
    col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])

    if not col_nopin or not col_plat:
        st.error(f"Kolom NOPIN/Plat tidak ditemukan di sheet '{armada_sheet}'.")
        return None, None, None

    # Buat dictionary referensi
    df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
    df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
    ref_dict = {}
    for _, row in df_ref.iterrows():
        nopin = row['NOPIN']
        ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
        if col_kec:
            ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip() if pd.notna(row[col_kec]) else ''
        if col_merk:
            ref_dict[nopin]['MERK'] = str(row[col_merk]).strip() if pd.notna(row[col_merk]) else ''
        if col_type:
            ref_dict[nopin]['TYPE'] = str(row[col_type]).strip() if pd.notna(row[col_type]) else ''

    # 3. Proses sheet harian (hanya sheet bernama angka 1-30)
    daily_sheets = [s for s in sheets_dict.keys() if s.isdigit()]
    if not daily_sheets:
        # fallback: semua sheet selain armada, Tugas, Master Data
        daily_sheets = [s for s in sheets_dict.keys() if s not in [armada_sheet, 'Tugas', 'Master Data']]

    cleaned_sheets = {}
    total_baris = 0
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, hari in enumerate(daily_sheets):
        status_text.text(f"Memproses sheet {hari} ({i+1}/{len(daily_sheets)})")
        progress_bar.progress((i+1)/len(daily_sheets))

        df_raw = sheets_dict[hari]
        # Deteksi header: cari baris yang mengandung "PINTU", "PLAT MOBIL", atau "NOPIN"
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

        # Pembersihan
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
        return None, None, None

    df_master = pd.concat(cleaned_sheets.values(), ignore_index=True)

    # Konversi kolom numerik (NETTO/GROSS/TARE)
    for col in df_master.columns:
        if 'NETTO' in col or 'GROSS' in col or 'TARE' in col or 'BERAT' in col:
            df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)

    # Identifikasi kolom tonase utama (Netto)
    col_netto = cari_kolom(df_master.columns, ['NETTO'])
    if not col_netto:
        col_netto = cari_kolom(df_master.columns, ['TOTAL', 'JUMLAH'])  # fallback

    # Hitung durasi jika kolom jam tersedia
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['WAKTU_MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['WAKTU_KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['WAKTU_KELUAR_DT'] - df_master['WAKTU_MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        try:
            df_master['JAM_INPUT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce').dt.hour
        except:
            pass

    st.success(f"✅ Data siap: {total_baris} baris dari {len(cleaned_sheets)} sheet.")

    # --- Agregasi analisis (sesuai poin notebook) ---
    # Poin 2: Per Kecamatan
    if 'Kecamatan' in df_master.columns:
        df_kecamatan = df_master.groupby('Kecamatan').agg(
            Jumlah_Armada_Aktif=('NOPIN', 'nunique'),
            Total_Ritase_Trip=('NOPIN', 'count'),
            Total_Tonase_Bersih=(col_netto, 'sum') if col_netto else ('NOPIN', 'count')
        ).reset_index().sort_values('Total_Ritase_Trip', ascending=False)
    else:
        df_kecamatan = pd.DataFrame()

    # Poin 3: Performa per Armada
    if col_netto:
        df_armada_performa = df_master.groupby(['NOPIN', 'NO_PLAT', 'Kecamatan', 'TYPE', 'MERK']).agg(
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

    # Poin 5: Efisiensi Waktu (jika ada)
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
    st.markdown("Unggah file Excel (.xls/.xlsx) dengan sheet **List Armada** dan sheet harian (1–30). Sistem otomatis menjalankan **seluruh proses notebook DLH**.")

    with st.sidebar:
        uploaded_file = st.file_uploader("📂 Pilih file Excel", type=["xlsx", "xls"])
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
            df = st.session_state.df_master.copy()
            df_kecamatan, df_armada_performa, df_tren_harian, df_durasi_kec, df_jam_puncak = st.session_state.dataframes

            # Filter interaktif
            st.sidebar.header("🔍 Filter")
            if 'Kecamatan' in df.columns:
                kec_list = ['Semua'] + sorted(df['Kecamatan'].dropna().unique().tolist())
                kec_terpilih = st.sidebar.selectbox("Kecamatan", kec_list)
                if kec_terpilih != 'Semua':
                    df = df[df['Kecamatan'] == kec_terpilih]
            if 'TANGGAL' in df.columns:
                tgl_list = sorted(df['TANGGAL'].unique())
                if len(tgl_list) > 1:
                    rentang = st.sidebar.date_input("Rentang Tanggal",
                        [pd.to_datetime(tgl_list[0]), pd.to_datetime(tgl_list[-1])])
                    if len(rentang) == 2:
                        df = df[(pd.to_datetime(df['TANGGAL']) >= pd.Timestamp(rentang[0])) &
                                (pd.to_datetime(df['TANGGAL']) <= pd.Timestamp(rentang[1]))]

            # Metrik
            total_trip = len(df)
            total_armada = df['NOPIN'].nunique()
            col_netto = cari_kolom(df.columns, ['NETTO']) or cari_kolom(df.columns, ['TOTAL'])
            total_tonase = df[col_netto].sum() / 1000 if col_netto else 0
            rata_durasi = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else 0

            col1, col2, col3, col4 = st.columns(4)
            col1.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{total_armada}</div></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase (Ton)</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
            col4.markdown(f'<div class="metric-card"><div class="metric-label">Rata² Durasi (menit)</div><div class="metric-value">{rata_durasi:.1f}</div></div>', unsafe_allow_html=True)

            st.markdown("---")

            # Visualisasi
            fig_tren = None
            if not df_tren_harian.empty:
                fig_tren = px.line(df_tren_harian, x='TANGGAL', y='Total_Ritase_Trip', title='Tren Ritase Harian', markers=True)
                fig_tren.update_traces(line_color='#0D9488', line_width=3)
                st.plotly_chart(fig_tren, use_container_width=True)
                st.session_state.figures['tren'] = fig_tren

            fig_kec = None
            if not df_kecamatan.empty:
                fig_kec = px.bar(df_kecamatan, x='Kecamatan', y='Total_Tonase_Bersih', color='Total_Tonase_Bersih',
                                 color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
                fig_kec.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_kec, use_container_width=True)
                st.session_state.figures['kec'] = fig_kec

            fig_jam = None
            if not df_jam_puncak.empty:
                fig_jam = px.area(df_jam_puncak, x='JAM_INPUT', y='Jumlah_Truk_Masuk', title='Pola Kedatangan per Jam')
                fig_jam.update_xaxes(dtick=1)
                st.plotly_chart(fig_jam, use_container_width=True)
                st.session_state.figures['jam'] = fig_jam

            fig_top = None
            if not df_armada_performa.empty:
                top_10 = df_armada_performa.head(10)
                fig_top = px.bar(top_10, x='NOPIN', y='Total_Trip', color='Total_Trip', title='10 Armada Teraktif')
                st.plotly_chart(fig_top, use_container_width=True)
                st.session_state.figures['top'] = fig_top

            # Ringkasan Eksekutif (Poin 7)
            st.subheader("📝 Ringkasan Eksekutif (Otomatis)")
            if not df_kecamatan.empty:
                kec_tertinggi = df_kecamatan.iloc[0]['Kecamatan']
                tonase_tertinggi = df_kecamatan.iloc[0]['Total_Tonase_Bersih'] / 1000
            else:
                kec_tertinggi = "N/A"
                tonase_tertinggi = 0

            eksekutif = f"""
**1. Ringkasan Operasional:**
- Total armada aktif: **{total_armada} unit**
- Total ritase: **{total_trip} trip**
- Total volume sampah: **{total_tonase:,.1f} Ton**
- Wilayah beban tertinggi: **{kec_tertinggi}** ({tonase_tertinggi:,.1f} Ton)

**2. Rekomendasi Strategis:**
- Optimasi rute & alokasi armada di **{kec_tertinggi}**
- Pengaturan shift untuk mengurangi kepadatan jam sibuk
- Evaluasi teknis unit dengan ritase rendah
"""
            st.markdown(eksekutif)

            # Laporan AI (opsional)
            if st.button("🔮 Buat Laporan AI (DeepSeek)"):
                with st.spinner("Menghubungi DeepSeek..."):
                    st.session_state.laporan_teks = laporan_ai(eksekutif)
                st.markdown("### 📄 Laporan AI")
                st.write(st.session_state.laporan_teks)

            # ========== UNDUH SEMUA HASIL (SEPERTI NOTEBOOK) ==========
            st.markdown("## 📥 Unduh Semua Hasil (seperti di notebook)")

            # Helper unduh
            @st.cache_data
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Sheet1')
                return output.getvalue()

            # Tombol unduh dalam beberapa kolom
            col_d1, col_d2, col_d3 = st.columns(3)
            with col_d1:
                # Master Data
                st.download_button("📊 Master Data Gabungan (Poin 1)", data=to_excel(df), file_name="Master_Data_Gabungan.xlsx")
                if not df_kecamatan.empty:
                    st.download_button("📊 Laporan Per Kecamatan (Poin 2)", data=to_excel(df_kecamatan), file_name="Laporan_Kecamatan.xlsx")
                if not df_armada_performa.empty:
                    st.download_button("📊 Performa Armada (Poin 3)", data=to_excel(df_armada_performa), file_name="Performa_Armada.xlsx")
            with col_d2:
                if not df_tren_harian.empty:
                    st.download_button("📈 Tren Harian (Poin 4)", data=to_excel(df_tren_harian), file_name="Tren_Harian.xlsx")
                if not df_durasi_kec.empty:
                    st.download_button("⏱️ Efisiensi Waktu (Poin 5)", data=to_excel(df_durasi_kec), file_name="Efisiensi_Waktu.xlsx")
                # Ringkasan teks
                st.download_button("📝 Ringkasan Eksekutif (TXT)", data=eksekutif.encode('utf-8'), file_name="Ringkasan_Eksekutif.txt")
            with col_d3:
                # Grafik PNG
                pilihan = st.selectbox("Pilih grafik untuk diunduh", ["Tren Ritase", "Tonase per Kecamatan", "Jam Sibuk", "Top Armada"])
                fig_dict = {'Tren Ritase': 'tren', 'Tonase per Kecamatan': 'kec', 'Jam Sibuk': 'jam', 'Top Armada': 'top'}
                fig_key = fig_dict.get(pilihan)
                if fig_key and fig_key in st.session_state.figures:
                    try:
                        png = pio.to_image(st.session_state.figures[fig_key], format='png', scale=2)
                        st.download_button("📸 Unduh Grafik (PNG)", data=png, file_name=f"{pilihan}.png", mime="image/png")
                    except:
                        st.warning("kaleido tidak terinstal, tidak bisa unduh PNG.")
                if st.session_state.laporan_teks:
                    st.download_button("📝 Laporan AI (TXT)", data=st.session_state.laporan_teks, file_name="laporan_ai.txt")

    else:
        st.info("👆 Silakan unggah file Excel Anda untuk memulai.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
