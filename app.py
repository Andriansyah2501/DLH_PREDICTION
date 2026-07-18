import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO

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

# -------------------------- API DEEPSEEK --------------------------
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

# -------------------------- FUNGSI BANTU --------------------------
def cari_kolom(kolom_list, kata_kunci):
    """Mencari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in kolom_list:
        col_upper = col.upper()
        for kw in kata_kunci:
            if kw in col_upper:
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
    Memproses data sesuai alur notebook DLH.
    Mengembalikan (df_master, ref_dict, armada_sheet) atau (None, None, None) jika gagal.
    """
    # 1. Cari sheet referensi (List Armada)
    armada_sheet = None
    for name in sheets_dict.keys():
        if 'list armada' in name.lower():
            armada_sheet = name
            break
    if armada_sheet is None:
        for name in sheets_dict.keys():
            if 'armada' in name.lower() or 'master' in name.lower():
                armada_sheet = name
                break
    if armada_sheet is None:
        st.error("❌ Sheet 'List Armada' tidak ditemukan. Pastikan ada sheet dengan nama tersebut.")
        return None, None, None

    # Baca ulang sheet armada dengan mencari header yang tepat (bisa di baris kedua seperti notebook)
    # Di notebook asli, header=1. Kita coba deteksi baris header dengan mencari kata "PINTU" atau "PLAT".
    df_arm_raw = sheets_dict[armada_sheet].copy()
    header_row = 0
    for idx, row in df_arm_raw.iterrows():
        row_str = ' '.join(row.astype(str).dropna().str.upper().values)
        if 'PINTU' in row_str or 'PLAT' in row_str:
            header_row = idx
            break
    # Baca ulang dengan header yang ditemukan
    xls = pd.ExcelFile(uploaded_file)
    df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_row)
    # Pastikan tidak ada kolom kosong (Unnamed) yang mengganggu
    df_ref = df_ref.loc[:, ~df_ref.columns.str.contains('^Unnamed')]
    df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]

    # Cari kolom penting dengan berbagai variasi
    col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'NO. PINTU', 'NO PINTU', 'PINTU', 'NOMOR PINTU', 'NOMOR.PINTU'])
    if not col_nopin:
        st.error(f"❌ Kolom No. Pintu tidak ditemukan di sheet '{armada_sheet}'. Kolom tersedia: {', '.join(df_ref.columns[:10])}")
        return None, None, None

    col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL', 'NOMOR POLISI', 'NO.PLAT', 'NO PLAT'])
    if not col_plat:
        st.error(f"❌ Kolom Plat Mobil tidak ditemukan di sheet '{armada_sheet}'.")
        return None, None, None

    col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI KECAMATAN', 'LOKASI', 'KEC'])
    col_merk = cari_kolom(df_ref.columns, ['MERK', 'MEREK', 'BRAND'])
    col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE', 'MODEL'])

    # Standarisasi dan buat dictionary
    df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
    df_ref['NO_PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
    
    # Mapping untuk sinkronisasi
    ref_columns = {'NO_PLAT': 'NO_PLAT'}
    if col_kec:
        df_ref['Kecamatan'] = df_ref[col_kec].astype(str).str.strip()
        ref_columns['Kecamatan'] = 'Kecamatan'
    if col_merk:
        df_ref['MERK'] = df_ref[col_merk].astype(str).str.strip()
        ref_columns['MERK'] = 'MERK'
    if col_type:
        df_ref['TYPE'] = df_ref[col_type].astype(str).str.strip()
        ref_columns['TYPE'] = 'TYPE'

    ref_dict = df_ref.set_index('NOPIN')[list(ref_columns.values())].to_dict('index')

    # 2. Proses sheet harian (nama sheet berupa angka, atau pilih otomatis selain armada)
    daily_sheets = [s for s in sheets_dict.keys() if s != armada_sheet]
    hari_sheets = []
    for s in daily_sheets:
        if s.isdigit():
            hari_sheets.append(s)
        elif 'harian' in s.lower() or 'daily' in s.lower():
            hari_sheets.append(s)
    if not hari_sheets:
        # Fallback: semua selain armada
        hari_sheets = daily_sheets

    cleaned_sheets = {}
    total_baris = 0

    progress_text = st.empty()
    progress_bar = st.progress(0)

    for i, hari in enumerate(hari_sheets):
        progress_text.text(f"Memproses sheet: {hari} ({i+1}/{len(hari_sheets)})")
        progress_bar.progress((i+1)/len(hari_sheets))

        df_raw = sheets_dict[hari].copy()
        # Cari header yang mengandung 'PINTU' atau 'PLAT MOBIL'
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str:
                header_idx = idx
                break
        if header_idx is None:
            continue  # lewati sheet tanpa header yang jelas

        # Baca ulang dengan header
        try:
            df_hari = pd.read_excel(uploaded_file, sheet_name=hari, header=header_idx)
        except:
            # Fallback manual
            df_hari = df_raw.iloc[header_idx+1:].reset_index(drop=True)
            df_hari.columns = df_raw.iloc[header_idx].tolist()

        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]

        # Cari kolom No. Pintu dan Plat
        col_nopin_h = cari_kolom(df_hari.columns, ['PINTU', 'NOPIN', 'NO. PINTU', 'NO PINTU'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL', 'NO.PLAT', 'NO PLAT'])
        if not col_nopin_h or not col_plat_h:
            continue

        df_hari.rename(columns={col_nopin_h: 'NOPIN', col_plat_h: 'NO_PLAT'}, inplace=True)

        # Pembersihan NOPIN
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

        # Sinkronisasi dengan referensi
        def sinkron(row):
            nopin = row['NOPIN']
            if nopin in ref_dict:
                row['NO_PLAT'] = ref_dict[nopin]['NO_PLAT']
                if 'Kecamatan' in ref_dict[nopin]:
                    row['Kecamatan'] = ref_dict[nopin]['Kecamatan']
                if 'MERK' in ref_dict[nopin]:
                    row['MERK'] = ref_dict[nopin]['MERK']
                if 'TYPE' in ref_dict[nopin]:
                    row['TYPE'] = ref_dict[nopin]['TYPE']
            return row

        df_hari = df_hari.apply(sinkron, axis=1)

        # Tambah kolom tanggal
        try:
            tgl = f"2026-06-{int(hari):02d}" if hari.isdigit() else hari
        except:
            tgl = hari
        df_hari['TANGGAL'] = tgl

        cleaned_sheets[hari] = df_hari
        total_baris += len(df_hari)

    progress_text.empty()
    progress_bar.empty()

    if not cleaned_sheets:
        st.error("❌ Tidak ada sheet harian yang berhasil diproses.")
        return None, None, None

    # Gabungkan semua data harian
    df_master = pd.concat(cleaned_sheets.values(), ignore_index=True)

    # Konversi kolom tonase numerik
    for col in df_master.columns:
        if 'NETTO' in col or 'GROSS' in col or 'TARE' in col or 'BERAT' in col:
            df_master[col] = pd.to_numeric(df_master[col], errors='coerce').fillna(0)

    # Deteksi kolom waktu
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['WAKTU_MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['WAKTU_KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['WAKTU_KELUAR_DT'] - df_master['WAKTU_MASUK_DT']).dt.total_seconds() / 60
        # Filter durasi tidak wajar
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]

    # Jika ada kolom jam masuk, buat JAM_INPUT
    if col_masuk:
        try:
            df_master['JAM_INPUT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce').dt.hour
        except:
            pass

    return df_master, ref_dict, armada_sheet

# -------------------------- SESSION STATE --------------------------
if "data_processed" not in st.session_state:
    st.session_state.data_processed = False
    st.session_state.df_master = None
    st.session_state.figures = {}
    st.session_state.laporan_teks = ""

# -------------------------- APLIKASI UTAMA --------------------------
def main():
    st.title("🚛 Dashboard Analitik DLH – Armada Sampah")
    st.markdown("Unggah file Excel (format .xls/.xlsx) dengan sheet **List Armada** dan sheet harian (1–30). Sistem akan otomatis memproses seperti notebook DLH dan menampilkan analisis lengkap.")

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
                df_master, ref_dict, armada_sheet = process_dlh_data(sheets_dict, uploaded_file)
                if df_master is not None:
                    st.session_state.df_master = df_master
                    st.session_state.ref_dict = ref_dict
                    st.session_state.data_processed = True
                    st.success("✅ Semua proses selesai! Data siap dianalisis.")
                    st.balloons()

        if st.session_state.data_processed:
            df = st.session_state.df_master.copy()

            # Sidebar filter
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

            # Metrik utama
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

            # Grafik
            if 'TANGGAL' in df.columns:
                tren_harian = df.groupby('TANGGAL').size().reset_index(name='Total_Ritase')
                tren_harian['TANGGAL'] = pd.to_datetime(tren_harian['TANGGAL'])
                tren_harian = tren_harian.sort_values('TANGGAL')
                fig_tren = px.line(tren_harian, x='TANGGAL', y='Total_Ritase', title='Tren Ritase Harian', markers=True)
                fig_tren.update_traces(line_color='#0D9488', line_width=3)
                st.plotly_chart(fig_tren, use_container_width=True)
                st.session_state.figures['tren'] = fig_tren

            if 'Kecamatan' in df.columns and col_netto:
                kec_agg = df.groupby('Kecamatan').agg(Total_Tonase=(col_netto[0], 'sum')).reset_index()
                kec_agg = kec_agg.sort_values('Total_Tonase', ascending=False)
                fig_kec = px.bar(kec_agg, x='Kecamatan', y='Total_Tonase', color='Total_Tonase',
                                 color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
                fig_kec.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig_kec, use_container_width=True)
                st.session_state.figures['kec'] = fig_kec

            if 'JAM_INPUT' in df.columns:
                jam_counts = df.groupby('JAM_INPUT').size().reset_index(name='Jumlah')
                jam_counts = jam_counts.sort_values('JAM_INPUT')
                fig_jam = px.area(jam_counts, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
                fig_jam.update_xaxes(dtick=1)
                st.plotly_chart(fig_jam, use_container_width=True)
                st.session_state.figures['jam'] = fig_jam

            top_armada = df.groupby('NOPIN').size().reset_index(name='Trip').nlargest(10, 'Trip')
            fig_top = px.bar(top_armada, x='NOPIN', y='Trip', color='Trip', title='10 Armada Teraktif')
            st.plotly_chart(fig_top, use_container_width=True)
            st.session_state.figures['top'] = fig_top

            # Laporan AI
            st.subheader("📝 Laporan Cerdas (DeepSeek AI)")
            statistik_teks = f"""
Total Trip: {total_trip}
Armada Aktif: {total_armada}
Total Tonase: {total_tonase:,.1f} Ton
Rata-rata Durasi: {rata_durasi:.1f} menit
            """
            if st.button("🔮 Buat Laporan AI"):
                with st.spinner("Menghubungi DeepSeek..."):
                    st.session_state.laporan_teks = laporan_ai(statistik_teks)
                st.markdown("### 📄 Hasil Laporan")
                st.write(st.session_state.laporan_teks)

            # Download
            st.markdown("## 📥 Unduh Hasil")
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
                fig = None
                if pilihan == "Tren Ritase" and 'tren' in st.session_state.figures:
                    fig = st.session_state.figures['tren']
                elif pilihan == "Tonase per Kecamatan" and 'kec' in st.session_state.figures:
                    fig = st.session_state.figures['kec']
                elif pilihan == "Jam Sibuk" and 'jam' in st.session_state.figures:
                    fig = st.session_state.figures['jam']
                elif pilihan == "Top Armada" and 'top' in st.session_state.figures:
                    fig = st.session_state.figures['top']
                if fig:
                    try:
                        png = pio.to_image(fig, format='png', scale=2)
                        st.download_button("📸 Unduh Grafik", data=png, file_name=f"{pilihan}.png", mime="image/png")
                    except:
                        st.warning("Gagal mengunduh grafik (kaleido tidak terinstal).")
            with col_dl4:
                if st.session_state.laporan_teks:
                    st.download_button("📝 Laporan AI", data=st.session_state.laporan_teks, file_name="laporan_ai.txt")
                else:
                    st.write("(Buat laporan AI dulu)")

    else:
        st.info("👆 Silakan unggah file Excel (format .xls atau .xlsx) untuk memulai analisis ala notebook DLH.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
