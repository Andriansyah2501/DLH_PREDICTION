import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO
from datetime import datetime

# -------------------------- KONFIGURASI HALAMAN --------------------------
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

# -------------------------- API DEEPSEEK (Opsional) --------------------------
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

# -------------------------- FUNGSI UTAMA PROSES DATA --------------------------
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
    Implementasi penuh alur notebook dlh.ipynb:
    - Baca List Armada (header=1 atau deteksi otomatis)
    - Baca 30 sheet harian dengan deteksi header
    - Bersihkan & sinkronkan
    - Gabungkan menjadi Master Data
    """
    # 1. Cari sheet referensi
    armada_sheet = None
    for name in sheets_dict.keys():
        if 'list armada' in name.lower():
            armada_sheet = name
            break
    if armada_sheet is None:
        st.error("❌ Sheet 'List Armada' tidak ditemukan dalam file.")
        return None, None

    # Baca List Armada (notebook menggunakan header=1, kita deteksi baris header)
    df_arm_raw = sheets_dict[armada_sheet]
    header_idx = None
    for idx, row in df_arm_raw.iterrows():
        row_str = " ".join(row.astype(str).dropna().str.upper().values)
        if 'NOPIN' in row_str or 'NO.PLAT' in row_str or 'PINTU' in row_str:
            header_idx = idx
            break
    if header_idx is None:
        header_idx = 1  # fallback seperti notebook

    # Baca ulang dengan header yang tepat
    xls = pd.ExcelFile(uploaded_file)
    df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_idx)
    df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

    # Cari kolom penting
    def cari_kolom(kolom_list, kata_kunci):
        for col in kolom_list:
            col_upper = col.upper()
            for kw in kata_kunci:
                if kw in col_upper:
                    return col
        return None

    col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'NO. PINTU', 'PINTU'])
    col_plat = cari_kolom(df_ref.columns, ['NO.PLAT', 'PLAT', 'NO PLAT'])
    col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
    col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK'])
    col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])

    if not col_nopin or not col_plat:
        st.error(f"Kolom No. Pintu/Plat tidak ditemukan di sheet '{armada_sheet}'. Kolom: {df_ref.columns.tolist()}")
        return None, None

    # Standarisasi & dictionary
    df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
    df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
    ref_dict = df_ref.set_index('NOPIN')[['NO.PLAT']].to_dict('index')

    if col_kec:
        df_ref['Kecamatan'] = df_ref[col_kec].astype(str).str.strip()
        for k in ref_dict:
            ref_dict[k]['Kecamatan'] = df_ref.loc[df_ref['NOPIN']==k, 'Kecamatan'].values[0]
    if col_merk:
        df_ref['MERK'] = df_ref[col_merk].astype(str).str.strip()
        for k in ref_dict:
            ref_dict[k]['MERK'] = df_ref.loc[df_ref['NOPIN']==k, 'MERK'].values[0]
    if col_type:
        df_ref['TYPE'] = df_ref[col_type].astype(str).str.strip()
        for k in ref_dict:
            ref_dict[k]['TYPE'] = df_ref.loc[df_ref['NOPIN']==k, 'TYPE'].values[0]

    # 2. Proses sheet harian (nama sheet digit atau terpilih)
    daily_sheets = [s for s in sheets_dict.keys() if s.isdigit()]  # seperti notebook: hanya sheet bernama angka
    if not daily_sheets:
        # fallback: semua sheet selain armada dan bukan 'Tugas'/'Master Data'
        daily_sheets = [s for s in sheets_dict.keys() if s != armada_sheet and s not in ['Tugas', 'Master Data']]

    cleaned = {}
    total_baris = 0
    progress_bar = st.progress(0)
    status_text = st.empty()

    for i, hari in enumerate(daily_sheets):
        status_text.text(f"Memproses sheet {hari} ({i+1}/{len(daily_sheets)})")
        progress_bar.progress((i+1)/len(daily_sheets))

        df_raw = sheets_dict[hari]
        # Cari header yang mengandung 'PINTU' atau 'PLAT MOBIL'
        header_hari = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str:
                header_hari = idx
                break
        if header_hari is None:
            continue  # lewati jika tidak ditemukan

        df_hari = pd.read_excel(uploaded_file, sheet_name=hari, header=header_hari)
        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]

        col_nopin_h = cari_kolom(df_hari.columns, ['PINTU', 'NOPIN', 'NO. PINTU'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT', 'NO PLAT'])
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
                if 'Kecamatan' in ref_dict[nopin]:
                    row['Kecamatan'] = ref_dict[nopin]['Kecamatan']
                if 'MERK' in ref_dict[nopin]:
                    row['MERK'] = ref_dict[nopin]['MERK']
                if 'TYPE' in ref_dict[nopin]:
                    row['TYPE'] = ref_dict[nopin]['TYPE']
            return row

        df_hari = df_hari.apply(sinkron, axis=1)
        df_hari['TANGGAL'] = f"2026-06-{int(hari):02d}" if hari.isdigit() else hari
        cleaned[hari] = df_hari
        total_baris += len(df_hari)

    progress_bar.empty()
    status_text.empty()

    if not cleaned:
        st.error("Tidak ada sheet harian yang berhasil diproses.")
        return None, None

    df_master = pd.concat(cleaned.values(), ignore_index=True)

    # Konversi numerik kolom tonase (cari NETTO/GROSS/TARE)
    for col in df_master.columns:
        if any(kw in col for kw in ['NETTO', 'GROSS', 'TARE', 'BERAT']):
            df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)

    # Hitung waktu jika ada (JAM MASUK/KELUAR)
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

    st.success(f"✅ Data berhasil diproses: {total_baris} baris dari {len(cleaned)} sheet harian.")
    return df_master, ref_dict

# -------------------------- SESSION STATE --------------------------
if "data_processed" not in st.session_state:
    st.session_state.data_processed = False
    st.session_state.df_master = None
    st.session_state.figures = {}
    st.session_state.laporan_teks = ""

# -------------------------- APLIKASI UTAMA --------------------------
def main():
    st.title("🚛 Dashboard Analitik DLH – Armada Sampah (Juni 2026)")
    st.markdown("Unggah file Excel Anda (format .xls/.xlsx) dengan sheet **List Armada** dan sheet harian (1–30). Proses otomatis seperti notebook DLH.")

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
                df_master, ref_dict = process_dlh_data(sheets_dict, uploaded_file)
                if df_master is not None:
                    st.session_state.df_master = df_master
                    st.session_state.ref_dict = ref_dict
                    st.session_state.data_processed = True
                    st.balloons()

        if st.session_state.data_processed:
            df = st.session_state.df_master.copy()

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
            col_netto = [c for c in df.columns if 'NETTO' in c]
            total_tonase = df[col_netto[0]].sum() / 1000 if col_netto else 0
            rata_durasi = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else 0

            col1, col2, col3, col4 = st.columns(4)
            col1.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{total_armada}</div></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase (Ton)</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
            col4.markdown(f'<div class="metric-card"><div class="metric-label">Rata² Durasi (menit)</div><div class="metric-value">{rata_durasi:.1f}</div></div>', unsafe_allow_html=True)

            st.markdown("---")

            # Grafik (persis seperti notebook)
            # Tren harian
            tren = df.groupby('TANGGAL').size().reset_index(name='Total_Ritase')
            tren['TANGGAL'] = pd.to_datetime(tren['TANGGAL'])
            tren = tren.sort_values('TANGGAL')
            fig1 = px.line(tren, x='TANGGAL', y='Total_Ritase', title='Tren Ritase Harian', markers=True)
            fig1.update_traces(line_color='#0D9488', line_width=3)
            st.plotly_chart(fig1, use_container_width=True)
            st.session_state.figures['tren'] = fig1

            # Distribusi kecamatan
            if 'Kecamatan' in df.columns and col_netto:
                kec = df.groupby('Kecamatan').agg(Total_Tonase=(col_netto[0], 'sum')).reset_index()
                fig2 = px.bar(kec.sort_values('Total_Tonase', ascending=False),
                              x='Kecamatan', y='Total_Tonase', color='Total_Tonase',
                              color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
                fig2.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig2, use_container_width=True)
                st.session_state.figures['kec'] = fig2

            # Jam sibuk (jika ada)
            if 'JAM_INPUT' in df.columns:
                jam = df.groupby('JAM_INPUT').size().reset_index(name='Jumlah')
                jam = jam.sort_values('JAM_INPUT')
                fig3 = px.area(jam, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
                fig3.update_xaxes(dtick=1)
                st.plotly_chart(fig3, use_container_width=True)
                st.session_state.figures['jam'] = fig3

            # Top 10 armada
            top_arm = df.groupby('NOPIN').size().reset_index(name='Trip').nlargest(10, 'Trip')
            fig4 = px.bar(top_arm, x='NOPIN', y='Trip', color='Trip', title='10 Armada Teraktif')
            st.plotly_chart(fig4, use_container_width=True)
            st.session_state.figures['top'] = fig4

            # Laporan AI (opsional)
            st.subheader("📝 Laporan Cerdas (DeepSeek AI)")
            if st.button("🔮 Buat Laporan AI"):
                statistik_teks = f"""
Total Trip: {total_trip}
Armada Aktif: {total_armada}
Total Tonase: {total_tonase:,.1f} Ton
Rata-rata Durasi: {rata_durasi:.1f} menit
                """
                with st.spinner("Menghubungi DeepSeek..."):
                    st.session_state.laporan_teks = laporan_ai(statistik_teks)
                st.markdown("### 📄 Hasil Laporan")
                st.write(st.session_state.laporan_teks)
            else:
                st.info("Klik tombol di atas untuk menghasilkan laporan otomatis (API Key diperlukan).")

            # Unduh berbagai file (seperti notebook)
            st.markdown("## 📥 Unduh Hasil Analisis")
            col_dl1, col_dl2, col_dl3, col_dl4 = st.columns(4)
            with col_dl1:
                @st.cache_data
                def to_excel(dataframe):
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        dataframe.to_excel(writer, index=False, sheet_name='Master Data')
                    return output.getvalue()
                st.download_button("📊 Master Data (Excel)", data=to_excel(df), file_name="master_data_dlh.xlsx")
            with col_dl2:
                csv = df.describe().to_csv().encode('utf-8')
                st.download_button("📈 Statistik (CSV)", data=csv, file_name="statistik.csv")
            with col_dl3:
                pilihan = st.selectbox("Pilih grafik", ["Tren Ritase", "Tonase per Kecamatan", "Jam Sibuk", "Top Armada"])
                fig_download = st.session_state.figures.get(
                    {'Tren Ritase':'tren', 'Tonase per Kecamatan':'kec', 'Jam Sibuk':'jam', 'Top Armada':'top'}[pilihan]
                )
                if fig_download:
                    try:
                        png = pio.to_image(fig_download, format='png', scale=2)
                        st.download_button("📸 Unduh Grafik (PNG)", data=png, file_name=f"{pilihan}.png", mime="image/png")
                    except:
                        st.warning("Gagal mengunduh grafik (kaleido tidak terinstal).")
            with col_dl4:
                if st.session_state.laporan_teks:
                    st.download_button("📝 Laporan AI (TXT)", data=st.session_state.laporan_teks, file_name="laporan_ai.txt")
                else:
                    st.write("(Buat laporan AI dulu)")

    else:
        st.info("👆 Silakan unggah file Excel Anda untuk memulai analisis.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
