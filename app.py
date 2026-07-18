import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO
from datetime import datetime

# ---------- Konfigurasi Halaman ----------
st.set_page_config(page_title="Dashboard Armada DLH", page_icon="🚛", layout="wide")

st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 16px; padding: 24px; color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15); text-align: center;
    }
    .metric-value { font-size: 2.6rem; font-weight: 800; margin: 8px 0; }
    .metric-label { font-size: 1rem; opacity: 0.9; text-transform: uppercase; letter-spacing: 1px; }
    .stButton button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 12px 28px; border-radius: 8px; font-weight: 600; }
    .stButton button:hover { transform: scale(1.02); }
</style>
""", unsafe_allow_html=True)

# ---------- API DeepSeek (opsional) ----------
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def laporan_ai(ringkasan: str) -> str:
    if not DEEPSEEK_API_KEY:
        return None
    try:
        headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
        prompt = f"Buat laporan singkat 3 paragraf dari ringkasan data armada berikut:\n{ringkasan}\nSertakan rekomendasi."
        resp = requests.post(DEEPSEEK_URL, headers=headers, json={
            "model": "deepseek-chat",
            "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7,
            "max_tokens": 500
        }, timeout=30)
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ Gagal menghasilkan laporan: {e}"

# ---------- Fungsi Bantu ----------
def cari_kolom(kolom_list, kata_kunci):
    """Mencari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in kolom_list:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file Excel...")
def baca_semua_sheet(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl' if uploaded_file.name.endswith('.xlsx') else None)
    except:
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

def proses_semua_data(sheets_dict, uploaded_file):
    """
    Melakukan seluruh tugas:
    1. Gabungkan 30 sheet harian menjadi Master Data
    2. Validasi dan sinkronisasi berdasarkan List Armada
    3. Hitung jumlah trip & tonase per armada
    4. Cari armada teraktif & paling tidak efisien
    5. Hitung rata-rata waktu tempuh per jenis armada
    """
    # --- 1. Cari sheet List Armada ---
    armada_sheet = None
    for nama in sheets_dict:
        if 'list armada' in nama.lower():
            armada_sheet = nama
            break
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)

    # --- Baca referensi armada (header=1 sesuai notebook) ---
    ref_dict = {}
    if armada_sheet:
        try:
            xls = pd.ExcelFile(uploaded_file)
            df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=1)
            df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]
            col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
            col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
            if col_nopin and col_plat:
                df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
                df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
                # Kolom tambahan
                col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
                col_merk = cari_kolom(df_ref.columns, ['MERK'])
                col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
                for _, row in df_ref.iterrows():
                    nopin = row['NOPIN']
                    ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
                    if col_kec:
                        ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip()
                    if col_merk:
                        ref_dict[nopin]['MERK'] = str(row[col_merk]).strip()
                    if col_type:
                        ref_dict[nopin]['TYPE'] = str(row[col_type]).strip()
                st.success(f"✅ Referensi List Armada dimuat ({len(ref_dict)} armada).")
            else:
                st.warning("⚠ Kolom NOPIN/Plat tidak ditemukan di List Armada. Sinkronisasi dilewati.")
        except Exception as e:
            st.warning(f"⚠ Gagal membaca List Armada: {e}")
    else:
        st.info("ℹ Sheet List Armada tidak ditemukan. Data harian akan diproses tanpa sinkronisasi.")

    # --- 2. Pilih sheet harian (nama digit 1-30) ---
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]
    if not daily_sheets:
        st.error("❌ Tidak ada sheet harian (bernama angka) yang ditemukan.")
        return None

    st.info(f"Memproses {len(daily_sheets)} sheet harian...")
    progress = st.progress(0)
    status = st.empty()

    cleaned = {}
    skipped = []
    for i, sheet in enumerate(daily_sheets):
        status.text(f"Sheet {sheet} ({i+1}/{len(daily_sheets)})")
        progress.progress((i+1)/len(daily_sheets))
        try:
            df_raw = sheets_dict[sheet]
        except:
            skipped.append(sheet)
            continue

        # Deteksi header (cari baris mengandung PINTU, PLAT MOBIL, atau NOPIN)
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
            df_hari = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_idx)
        except:
            skipped.append(sheet)
            continue

        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]
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

        # Sinkronisasi (Tugas 2): samakan plat dan data lain berdasarkan List Armada
        if ref_dict:
            def sinkron(row):
                nopin = row['NOPIN']
                if nopin in ref_dict:
                    row['NO_PLAT'] = ref_dict[nopin]['NO.PLAT']
                    for key in ['Kecamatan', 'MERK', 'TYPE']:
                        if key in ref_dict[nopin]:
                            row[key] = ref_dict[nopin][key]
                return row
            df_hari = df_hari.apply(sinkron, axis=1)

        # Tambah kolom TANGGAL
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        cleaned[sheet] = df_hari

    progress.empty()
    status.empty()

    if not cleaned:
        st.error(f"❌ Tidak ada sheet harian yang valid. {len(skipped)} sheet dilewati.")
        return None

    # Gabung semua sheet harian → Master Data (Tugas 1 selesai)
    df_master = pd.concat(cleaned.values(), ignore_index=True)
    st.success(f"✅ Master Data terbentuk: {len(df_master)} baris dari {len(cleaned)} sheet.")

    # Konversi kolom tonase numerik
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # Deteksi kolom waktu untuk durasi (Tugas 5)
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

    # --- Agregasi untuk analisis ---
    # Per Kecamatan (untuk grafik, dll)
    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan').agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Ritase', ascending=False)

    # Performa per Armada (Tugas 3 & 4)
    df_armada = pd.DataFrame()
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        df_armada = df_master.groupby(group_cols).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)

        # Tugas 4: Armada teraktif & paling tidak efisien
        if not df_armada.empty:
            teraktif = df_armada.iloc[0]
            tidak_efisien = df_armada.iloc[-1]
        else:
            teraktif = tidak_efisien = None
    else:
        teraktif = tidak_efisien = None

    # Tren Harian (Tugas 6)
    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL').agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('TANGGAL')

    # Rata-rata waktu tempuh per jenis armada (Tugas 5)
    df_waktu_jenis = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'TYPE' in df_master.columns:
        df_waktu_jenis = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE')['DURASI_MENIT'].mean().reset_index()
        df_waktu_jenis.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']

    # Ringkasan untuk laporan (Tugas 7)
    ringkasan = {
        'total_trip': len(df_master),
        'total_armada': df_master['NOPIN'].nunique(),
        'total_tonase': df_master[col_netto].sum() if col_netto else 0,
        'kec_terpadat': df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else '-',
        'armada_teraktif': teraktif,
        'armada_tidak_efisien': tidak_efisien,
        'rata_waktu_jenis': df_waktu_jenis
    }

    return df_master, df_armada, df_kec, df_tren, df_waktu_jenis, ringkasan

# ---------- Session State ----------
if "hasil" not in st.session_state:
    st.session_state.hasil = None
if "figures" not in st.session_state:
    st.session_state.figures = {}
if "laporan_ai" not in st.session_state:
    st.session_state.laporan_ai = None

# ---------- Aplikasi Utama ----------
def main():
    st.title("🚛 Dashboard Analitik Armada – DLH Kota Batam")
    st.markdown("Unggah file Excel dengan 30 sheet harian dan 1 sheet **List Armada**. Semua analisis dilakukan otomatis.")

    with st.sidebar:
        uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
        if uploaded_file:
            st.success("File siap diproses")
        if st.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Membaca dan memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if not sheets:
                    st.error("File tidak memiliki sheet yang bisa dibaca.")
                else:
                    hasil = proses_semua_data(sheets, uploaded_file)
                    if hasil:
                        st.session_state.hasil = hasil
                        st.session_state.figures = {}  # reset grafik
                        st.session_state.laporan_ai = None
                        st.balloons()
                    else:
                        st.session_state.hasil = None

    if st.session_state.hasil is not None:
        df_master, df_armada, df_kec, df_tren, df_waktu_jenis, ringkasan = st.session_state.hasil

        # --- Sidebar filter ---
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
        df = df_master.copy()
        if kec_terpilih != 'Semua':
            df = df[df['Kecamatan'] == kec_terpilih]
        if 'rentang' in locals() and len(rentang) == 2:
            df = df[(pd.to_datetime(df['TANGGAL']) >= pd.Timestamp(rentang[0])) & 
                    (pd.to_datetime(df['TANGGAL']) <= pd.Timestamp(rentang[1]))]

        # --- Metrik Utama ---
        total_trip = len(df)
        total_armada = df['NOPIN'].nunique()
        ton_col = cari_kolom(df.columns, ['NETTO', 'GROSS', 'TOTAL'])
        total_tonase = df[ton_col].sum() / 1000 if ton_col else 0
        durasi_rata = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else None

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Trip", total_trip)
        col2.metric("Armada Aktif", total_armada)
        col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
        col4.metric("Rata² Durasi (menit)", f"{durasi_rata:.1f}" if durasi_rata else "-")

        st.markdown("---")

        # --- Tugas 4: Armada Teraktif & Tidak Efisien ---
        if not df_armada.empty:
            teraktif = df_armada.iloc[0]
            tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if any(df_armada['Total_Trip'] > 0) else df_armada.iloc[-1]
            col_a, col_b = st.columns(2)
            with col_a:
                st.success(f"🥇 **Armada Teraktif**: {teraktif['NOPIN']} ({teraktif['NO_PLAT']}) – {int(teraktif['Total_Trip'])} trip")
            with col_b:
                st.error(f"🐌 **Armada Paling Tidak Efisien**: {tidak_efisien['NOPIN']} ({tidak_efisien['NO_PLAT']}) – {int(tidak_efisien['Total_Trip'])} trip")

        # --- Tugas 5: Rata-rata Waktu Tempuh per Jenis Armada ---
        if not df_waktu_jenis.empty:
            st.subheader("⏱️ Rata‑rata Waktu Tempuh per Jenis Armada")
            st.dataframe(df_waktu_jenis.style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}), use_container_width=True)

        # --- Grafik Interaktif (Tugas 6) ---
        st.subheader("📊 Visualisasi Data")
        # 1. Tren Harian
        if not df_tren.empty:
            fig1 = px.line(df_tren, x='TANGGAL', y='Total_Ritase', title='Tren Frekuensi Ritase Harian', markers=True)
            fig1.update_traces(line_color='#0D9488')
            st.plotly_chart(fig1, use_container_width=True)
            st.session_state.figures['tren'] = fig1

        # 2. Distribusi Tonase per Kecamatan
        if not df_kec.empty:
            fig2 = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', 
                          color='Total_Tonase', color_continuous_scale='Viridis',
                          title='Total Tonase per Kecamatan')
            fig2.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig2, use_container_width=True)
            st.session_state.figures['kec'] = fig2

        # 3. Top 10 Armada
        if not df_armada.empty:
            top10 = df_armada.head(10)
            fig3 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip',
                          color_continuous_scale='OrRd', title='10 Armada Teraktif')
            st.plotly_chart(fig3, use_container_width=True)
            st.session_state.figures['top'] = fig3

        # 4. Pola Jam Sibuk (jika data waktu ada)
        if 'JAM_INPUT' in df.columns:
            jam_counts = df.dropna(subset=['JAM_INPUT']).groupby('JAM_INPUT').size().reset_index(name='Jumlah')
            fig4 = px.area(jam_counts, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
            st.plotly_chart(fig4, use_container_width=True)
            st.session_state.figures['jam'] = fig4

        st.markdown("---")

        # --- Tugas 7: Laporan Singkat ---
        st.subheader("📝 Laporan Ringkas Otomatis")
        # Gunakan data ringkasan yang sudah dihitung
        kec_tertinggi = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else "-"
        laporan_teks = f"""
**Ringkasan Operasional:**
- Total trip: {total_trip}
- Armada aktif: {total_armada} unit
- Total tonase: {total_tonase:,.1f} Ton
- Kecamatan dengan aktivitas tertinggi: {kec_tertinggi}

**Armada Teraktif:** {teraktif['NOPIN']} ({teraktif['NO_PLAT']}) – {int(teraktif['Total_Trip'])} trip
**Armada Paling Tidak Efisien:** {tidak_efisien['NOPIN']} ({tidak_efisien['NO_PLAT']}) – {int(tidak_efisien['Total_Trip'])} trip

**Rekomendasi:**
- Alokasi tambahan unit untuk Kecamatan {kec_tertinggi}
- Periksa kondisi armada dengan trip rendah
        """
        st.markdown(laporan_teks)

        # Tombol unduh ringkasan sebagai TXT
        st.download_button("📄 Unduh Ringkasan (TXT)", laporan_teks.encode('utf-8'), "ringkasan.txt")

        # --- Laporan AI (opsional) ---
        if st.button("🤖 Buat Laporan AI (DeepSeek)"):
            laporan = laporan_ai(laporan_teks)
            if laporan:
                st.session_state.laporan_ai = laporan
                st.markdown("### 📄 Laporan AI")
                st.write(laporan)
            else:
                st.warning("Laporan AI tidak tersedia.")

        # --- Unduh Data & Grafik ---
        st.subheader("📥 Unduh Hasil Analisis")
        @st.cache_data
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w:
                df.to_excel(w, index=False)
            return output.getvalue()

        col1, col2, col3 = st.columns(3)
        with col1:
            st.download_button("📊 Master Data", to_excel(df_master), "Master_Data.xlsx")
            if not df_armada.empty:
                st.download_button("📊 Statistik Armada", to_excel(df_armada), "Statistik_Armada.xlsx")
            if not df_kec.empty:
                st.download_button("📊 Laporan Kecamatan", to_excel(df_kec), "Kecamatan.xlsx")
        with col2:
            if not df_tren.empty:
                st.download_button("📈 Tren Harian", to_excel(df_tren), "Tren_Harian.xlsx")
            if not df_waktu_jenis.empty:
                st.download_button("⏱️ Waktu per Jenis Armada", to_excel(df_waktu_jenis), "Waktu_per_Jenis.xlsx")
            if st.session_state.laporan_ai:
                st.download_button("🤖 Laporan AI (TXT)", st.session_state.laporan_ai, "laporan_ai.txt")
        with col3:
            pilihan = st.selectbox("Pilih grafik untuk diunduh (PNG)", ["Tren", "Kecamatan", "Top Armada", "Jam"])
            fig_key = {'Tren':'tren', 'Kecamatan':'kec', 'Top Armada':'top', 'Jam':'jam'}.get(pilihan)
            fig = st.session_state.figures.get(fig_key)
            if fig:
                try:
                    img_bytes = pio.to_image(fig, format='png', scale=2)
                    st.download_button("📸 Unduh Grafik", img_bytes, f"grafik_{pilihan}.png", "image/png")
                except Exception as e:
                    st.warning(f"Gagal mengunduh grafik: {e}")

        # --- Tampilkan data (opsional) ---
        with st.expander("🔎 Lihat Data Mentah (200 baris pertama)"):
            st.dataframe(df_master.head(200))

    else:
        st.info("👆 Silakan unggah file Excel Anda dan klik **Proses Data**.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
