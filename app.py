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
    .stButton button { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; border: none; padding: 12px 28px; border-radius: 8px; font-weight: 600; }
    .stButton button:hover { transform: scale(1.02); }
    .warning-box { background-color: #fff3cd; padding: 15px; border-radius: 8px; border-left: 5px solid #ffc107; }
</style>
""", unsafe_allow_html=True)

# -------------------------- API DEEPSEEK (opsional) --------------------------
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def laporan_ai(statistik: str) -> str:
    if not DEEPSEEK_API_KEY:
        return None
    try:
        headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
        prompt = f"Buat laporan singkat (3 paragraf) dalam bahasa Indonesia dari statistik berikut:\n{statistik}\nSertakan rekomendasi."
        resp = requests.post(DEEPSEEK_URL, headers=headers, json={
            "model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7, "max_tokens": 500
        }, timeout=30)
        return resp.json()["choices"][0]["message"]["content"]
    except:
        return None

# -------------------------- FUNGSI BANTU --------------------------
def cari_kolom(kolom, kata_kunci):
    """Cari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in kolom:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file...")
def baca_semua_sheet(uploaded_file):
    """Baca semua sheet, return dict {nama_sheet: dataframe}."""
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"Gagal membaca file: {e}")
        return {}
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except:
            pass
    return sheets

# -------------------------- PROSES DATA UTAMA --------------------------
def proses_data_utama(sheets_dict, uploaded_file):
    """
    Mengembalikan (df_master, dict_agregasi) atau None jika gagal total.
    dict_agregasi berisi: 'kecamatan', 'armada_performa', 'tren_harian', 'durasi_kec', 'jam_puncak'
    """
    # 1. Cari sheet armada (List Armada)
    armada_sheet = None
    for name in sheets_dict:
        if 'list armada' in name.lower():
            armada_sheet = name
            break
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)
    if armada_sheet is None:
        st.warning("Sheet 'List Armada' tidak ditemukan. Data tidak dapat disinkronkan.")
        ref_dict = {}
    else:
        # Baca referensi
        df_arm = sheets_dict[armada_sheet]
        # Coba deteksi header (biasanya baris 1)
        header_arm = 1
        xls = pd.ExcelFile(uploaded_file)
        try:
            df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=header_arm)
        except:
            st.warning(f"Gagal membaca sheet '{armada_sheet}' dengan header={header_arm}. Data master tidak digunakan.")
            ref_dict = {}
        else:
            df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]
            col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
            col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
            if col_nopin and col_plat:
                df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
                df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
                # Kolom tambahan opsional
                col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
                col_merk = cari_kolom(df_ref.columns, ['MERK'])
                col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
                ref_dict = {}
                for _, row in df_ref.iterrows():
                    nopin = row['NOPIN']
                    ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
                    if col_kec:
                        ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip()
                    if col_merk:
                        ref_dict[nopin]['MERK'] = str(row[col_merk]).strip()
                    if col_type:
                        ref_dict[nopin]['TYPE'] = str(row[col_type]).strip()
            else:
                st.warning(f"Kolom NOPIN/Plat tidak ditemukan di '{armada_sheet}'. Sinkronisasi dibatalkan.")
                ref_dict = {}

    # 2. Tentukan sheet harian (digit 1-30)
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]
    if not daily_sheets:
        st.error("Tidak ada sheet harian yang bisa diproses.")
        return None

    # 3. Proses setiap sheet harian
    cleaned_sheets = {}
    skipped = []
    total_rows = 0
    prog = st.progress(0)
    status_text = st.empty()

    for i, sheet in enumerate(daily_sheets):
        status_text.text(f"Memproses sheet {sheet} ({i+1}/{len(daily_sheets)})")
        prog.progress((i+1)/len(daily_sheets))
        try:
            df_raw = sheets_dict[sheet]
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
            df_hari = pd.read_excel(uploaded_file, sheet_name=sheet, header=header_idx)
        except:
            skipped.append(sheet)
            continue
        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]

        # Cari kolom NOPIN & Plat di harian
        col_nopin_h = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
        col_plat_h = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
        if not col_nopin_h or not col_plat_h:
            skipped.append(sheet)
            continue
        df_hari.rename(columns={col_nopin_h: 'NOPIN', col_plat_h: 'NO_PLAT'}, inplace=True)

        # Pembersihan
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

        # Sinkronisasi (jika master tersedia)
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

        # Tanggal
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        cleaned_sheets[sheet] = df_hari
        total_rows += len(df_hari)

    prog.empty()
    status_text.empty()

    if not cleaned_sheets:
        st.error("Tidak ada satupun sheet harian yang berhasil diproses.")
        return None

    df_master = pd.concat(cleaned_sheets.values(), ignore_index=True)

    # Konversi numerik kolom tonase
    tonase_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if tonase_col:
        df_master[tonase_col] = pd.to_numeric(df_master[tonase_col], errors='coerce').fillna(0)
    # Cari kolom Netto utama untuk agregasi
    col_netto = cari_kolom(df_master.columns, ['NETTO'])
    if not col_netto:
        col_netto = tonase_col  # fallback ke kolom tonase apapun

    # Analisis waktu (jika ada)
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2'])
    if col_masuk and col_keluar:
        df_master['WAKTU_MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['WAKTU_KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['WAKTU_KELUAR_DT'] - df_master['WAKTU_MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        try:
            df_master['JAM_INPUT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce').dt.hour
        except:
            pass

    # Agregasi (hanya jika kolom tersedia)
    agregasi = {}
    if 'Kecamatan' in df_master.columns and col_netto:
        agregasi['kecamatan'] = df_master.groupby('Kecamatan').agg(
            Jumlah_Armada=('NOPIN', 'nunique'),
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index()
    else:
        agregasi['kecamatan'] = pd.DataFrame()

    if col_netto:
        # Performa armada
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        agregasi['armada_performa'] = df_master.groupby(group_cols).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Tonase', ascending=False)
        # Tren harian
        agregasi['tren_harian'] = df_master.groupby('TANGGAL').agg(
            Total_Ritase=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index()
    else:
        agregasi['armada_performa'] = pd.DataFrame()
        agregasi['tren_harian'] = pd.DataFrame()

    if 'DURASI_MENIT' in df_master.columns and 'Kecamatan' in df_master.columns:
        df_valid = df_master.dropna(subset=['DURASI_MENIT'])
        agregasi['durasi_kec'] = df_valid.groupby('Kecamatan')['DURASI_MENIT'].mean().reset_index()
        if 'JAM_INPUT' in df_valid.columns:
            agregasi['jam_puncak'] = df_valid.groupby('JAM_INPUT').size().reset_index(name='Jumlah')
        else:
            agregasi['jam_puncak'] = pd.DataFrame()
    else:
        agregasi['durasi_kec'] = pd.DataFrame()
        agregasi['jam_puncak'] = pd.DataFrame()

    st.success(f"✅ {total_rows} baris dari {len(cleaned_sheets)} sheet berhasil diproses. {len(skipped)} sheet dilewati.")
    return df_master, agregasi

# -------------------------- SESSION STATE --------------------------
if "data" not in st.session_state:
    st.session_state.data = None
    st.session_state.agregasi = None
    st.session_state.figs = {}

# -------------------------- APLIKASI UTAMA --------------------------
def main():
    st.title("🚛 Dashboard DLH Armada – Analisis Otomatis & Unduh Lengkap")
    st.markdown("Unggah file Excel (format .xls/.xlsx) dengan sheet **List Armada** dan sheet harian (1-30). Sistem akan **melewatkan data yang tidak valid** dan hanya menampilkan hasil yang benar.")

    with st.sidebar:
        uploaded_file = st.file_uploader("📂 Pilih file Excel", type=["xlsx", "xls"])
        if uploaded_file:
            st.success("File siap diproses")
        if st.button("🚀 Mulai Proses", use_container_width=True):
            with st.spinner("Memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if not sheets:
                    st.error("File kosong atau tidak terbaca.")
                else:
                    hasil = proses_data_utama(sheets, uploaded_file)
                    if hasil is not None:
                        st.session_state.data, st.session_state.agregasi = hasil
                        st.session_state.figs = {}
                        st.balloons()
                    else:
                        st.error("Proses gagal. Periksa kembali file Anda.")

    if st.session_state.data is not None:
        df = st.session_state.data
        agg = st.session_state.agregasi

        # Filter sidebar
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Filter")
        if 'Kecamatan' in df.columns:
            opts = ['Semua'] + sorted(df['Kecamatan'].dropna().unique().tolist())
            kec = st.sidebar.selectbox("Kecamatan", opts)
            if kec != 'Semua':
                df = df[df['Kecamatan'] == kec]
        if 'TANGGAL' in df.columns:
            tgl = sorted(df['TANGGAL'].unique())
            if len(tgl) > 1:
                rentang = st.sidebar.date_input("Rentang Tanggal", [pd.to_datetime(tgl[0]), pd.to_datetime(tgl[-1])])
                if len(rentang) == 2:
                    df = df[(pd.to_datetime(df['TANGGAL']) >= pd.Timestamp(rentang[0])) &
                            (pd.to_datetime(df['TANGGAL']) <= pd.Timestamp(rentang[1]))]

        # Metrik
        total_trip = len(df)
        total_armada = df['NOPIN'].nunique()
        col_ton = cari_kolom(df.columns, ['NETTO', 'TOTAL']) or 'NOPIN'
        total_tonase = df[col_ton].sum() / 1000 if col_ton in df.columns else 0
        durasi_rata = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else 0

        c1,c2,c3,c4 = st.columns(4)
        c1.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
        c2.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{total_armada}</div></div>', unsafe_allow_html=True)
        c3.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase (Ton)</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
        c4.markdown(f'<div class="metric-card"><div class="metric-label">Rata² Durasi (menit)</div><div class="metric-value">{durasi_rata:.1f}</div></div>', unsafe_allow_html=True)

        st.markdown("---")

        # Grafik
        if not agg['tren_harian'].empty:
            fig1 = px.line(agg['tren_harian'], x='TANGGAL', y='Total_Ritase', title='Tren Ritase Harian', markers=True)
            fig1.update_traces(line_color='#0D9488')
            st.plotly_chart(fig1, use_container_width=True)
            st.session_state.figs['tren'] = fig1

        if not agg['kecamatan'].empty:
            fig2 = px.bar(agg['kecamatan'], x='Kecamatan', y='Total_Tonase', color='Total_Tonase',
                          color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
            st.plotly_chart(fig2, use_container_width=True)
            st.session_state.figs['kec'] = fig2

        if not agg.get('jam_puncak', pd.DataFrame()).empty:
            fig3 = px.area(agg['jam_puncak'], x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
            st.plotly_chart(fig3, use_container_width=True)
            st.session_state.figs['jam'] = fig3

        if not agg['armada_performa'].empty:
            top10 = agg['armada_performa'].head(10)
            fig4 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip', title='10 Armada Teraktif')
            st.plotly_chart(fig4, use_container_width=True)
            st.session_state.figs['top'] = fig4

        # Ringkasan eksekutif
        st.markdown("---")
        st.header("📝 Ringkasan Eksekutif")
        kec_tertinggi = agg['kecamatan'].iloc[0]['Kecamatan'] if not agg['kecamatan'].empty else "N/A"
        ton_tertinggi = agg['kecamatan'].iloc[0]['Total_Tonase']/1000 if not agg['kecamatan'].empty else 0
        hari_puncak = agg['tren_harian'].loc[agg['tren_harian']['Total_Ritase'].idxmax()]['TANGGAL'] if not agg['tren_harian'].empty else "N/A"
        st.markdown(f"""
- **Total Trip:** {total_trip}
- **Armada Aktif:** {total_armada}
- **Total Tonase:** {total_tonase:,.1f} Ton
- **Wilayah Tertinggi:** {kec_tertinggi} ({ton_tertinggi:,.1f} Ton)
- **Hari Tersibuk:** {hari_puncak}
        """)

        # Laporan AI (opsional)
        if st.button("🤖 Buat Laporan AI (DeepSeek)"):
            stat = f"Trip:{total_trip}, Armada:{total_armada}, Tonase:{total_tonase} Ton, Kec:{kec_tertinggi}"
            laporan = laporan_ai(stat)
            if laporan:
                st.session_state.laporan = laporan
                st.write(laporan)
            else:
                st.warning("Gagal menghasilkan laporan AI.")
        else:
            st.session_state.laporan = None

        # Unduh file
        st.markdown("---")
        st.header("📥 Unduh Semua Hasil")
        @st.cache_data
        def to_excel(dataframe):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w:
                dataframe.to_excel(w, index=False)
            return output.getvalue()

        col_u1, col_u2, col_u3 = st.columns(3)
        with col_u1:
            st.download_button("📊 Master Data (Excel)", data=to_excel(df), file_name="Master_Data.xlsx")
            if not agg['kecamatan'].empty:
                st.download_button("📊 Laporan Kecamatan", data=to_excel(agg['kecamatan']), file_name="Kecamatan.xlsx")
            if not agg['armada_performa'].empty:
                st.download_button("📊 Performa Armada", data=to_excel(agg['armada_performa']), file_name="Performa_Armada.xlsx")
        with col_u2:
            if not agg['tren_harian'].empty:
                st.download_button("📈 Tren Harian", data=to_excel(agg['tren_harian']), file_name="Tren_Harian.xlsx")
            if not agg['durasi_kec'].empty:
                st.download_button("⏱️ Durasi per Kecamatan", data=to_excel(agg['durasi_kec']), file_name="Durasi_Kec.xlsx")
            st.download_button("📝 Ringkasan (TXT)", data=f"Ringkasan Eksekutif\nTrip:{total_trip}\nTonase:{total_tonase}\n".encode(), file_name="Ringkasan.txt")
        with col_u3:
            pilihan = st.selectbox("Grafik untuk diunduh", ["Tren","Kecamatan","Jam","Top"])
            fig_key = {'Tren':'tren','Kecamatan':'kec','Jam':'jam','Top':'top'}.get(pilihan)
            if fig_key and fig_key in st.session_state.figs:
                try:
                    img = pio.to_image(st.session_state.figs[fig_key], format='png', scale=2)
                    st.download_button("📸 Unduh Grafik", data=img, file_name=f"{pilihan}.png", mime="image/png")
                except:
                    st.warning("kaleido tidak terinstal")
            if st.session_state.get('laporan'):
                st.download_button("🤖 Laporan AI", data=st.session_state.laporan, file_name="laporan_ai.txt")

        # Data mentah (expandable)
        with st.expander("🔍 Lihat Data Mentah (200 baris pertama)"):
            st.dataframe(df.head(200))

    else:
        st.info("👆 Unggah file dan klik 'Mulai Proses'.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
