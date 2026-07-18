import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO

# -------------------------- KONFIGURASI --------------------------
st.set_page_config(page_title="Dashboard DLH Armada", page_icon="🚛", layout="wide")

# -------------------------- API DEEPSEEK (opsional) --------------------------
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def laporan_ai(statistik: str):
    if not DEEPSEEK_API_KEY:
        return None
    try:
        headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
        prompt = f"Buat laporan singkat 3 paragraf tentang data armada berikut:\n{statistik}"
        resp = requests.post(DEEPSEEK_URL, headers=headers, json={
            "model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}],
            "temperature": 0.7, "max_tokens": 500
        }, timeout=30)
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        st.warning(f"Gagal menghubungi DeepSeek: {e}")
        return None

# -------------------------- FUNGSI BANTU --------------------------
def cari_kolom(kolom_list, kata_kunci):
    """Cari kolom yang mengandung salah satu kata kunci (case-insensitive)."""
    for col in kolom_list:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file...")
def baca_semua_sheet(uploaded_file):
    """Membaca semua sheet dari file Excel, return dict {nama_sheet: dataframe}."""
    try:
        xls = pd.ExcelFile(uploaded_file)
    except Exception as e:
        st.error(f"❌ File tidak bisa dibaca. Pastikan format .xls/.xlsx valid. Error: {e}")
        return {}
    sheets = {}
    for name in xls.sheet_names:
        try:
            df = pd.read_excel(xls, sheet_name=name)
            if not df.empty:
                sheets[name] = df
        except Exception as e:
            st.warning(f"Sheet '{name}' dilewati karena error: {e}")
    return sheets

def proses_data_utama(sheets_dict, uploaded_file):
    """
    Menggabungkan dan membersihkan data.
    Mengembalikan (df_master, df_kecamatan, df_armada, df_tren, df_durasi, df_jam) 
    atau None jika tidak ada data harian yang valid.
    """
    if not sheets_dict:
        st.error("Tidak ada sheet yang bisa dibaca.")
        return None

    # ========== 1. Cari sheet List Armada (opsional) ==========
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)
    ref_dict = {}
    if armada_sheet:
        try:
            df_arm = sheets_dict[armada_sheet]
            # Coba header = 1 (seperti notebook)
            xls = pd.ExcelFile(uploaded_file)
            df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=1)
            df_ref.columns = [str(c).strip().upper() for c in df_ref.columns]
            col_nopin = cari_kolom(df_ref.columns, ['NOPIN', 'PINTU'])
            col_plat = cari_kolom(df_ref.columns, ['PLAT', 'NOPOL'])
            if col_nopin and col_plat:
                df_ref['NOPIN'] = df_ref[col_nopin].astype(str).str.strip().str.upper()
                df_ref['NO.PLAT'] = df_ref[col_plat].astype(str).str.strip().str.upper()
                for _, row in df_ref.iterrows():
                    nopin = row['NOPIN']
                    ref_dict[nopin] = {'NO.PLAT': row['NO.PLAT']}
                    col_kec = cari_kolom(df_ref.columns, ['KECAMATAN', 'LOKASI'])
                    col_merk = cari_kolom(df_ref.columns, ['MERK'])
                    col_type = cari_kolom(df_ref.columns, ['TYPE', 'TIPE'])
                    if col_kec:
                        ref_dict[nopin]['Kecamatan'] = str(row[col_kec]).strip()
                    if col_merk:
                        ref_dict[nopin]['MERK'] = str(row[col_merk]).strip()
                    if col_type:
                        ref_dict[nopin]['TYPE'] = str(row[col_type]).strip()
                st.success(f"✅ Referensi List Armada dimuat ({len(ref_dict)} armada).")
            else:
                st.warning("⚠️ Kolom NOPIN/Plat tidak ditemukan di List Armada. Sinkronisasi dilewati.")
        except Exception as e:
            st.warning(f"⚠️ Gagal membaca List Armada: {e}. Data harian tetap diproses tanpa sinkronisasi.")
    else:
        st.info("ℹ️ Sheet List Armada tidak ditemukan. Data harian diproses tanpa master.")

    # ========== 2. Pilih sheet harian ==========
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

        # Cari baris header (seperti notebook: PINTU, PLAT MOBIL, NOPIN)
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
        except Exception as e:
            skipped.append(sheet)
            continue

        df_hari.columns = [str(c).strip().upper() for c in df_hari.columns]
        col_nopin = cari_kolom(df_hari.columns, ['NOPIN', 'PINTU'])
        col_plat = cari_kolom(df_hari.columns, ['PLAT', 'NOPOL'])
        if not col_nopin or not col_plat:
            skipped.append(sheet)
            continue

        df_hari = df_hari.rename(columns={col_nopin: 'NOPIN', col_plat: 'NO_PLAT'})

        # Pembersihan
        df_hari = df_hari.dropna(subset=['NOPIN'])
        df_hari['NOPIN'] = df_hari['NOPIN'].astype(str).str.strip().str.upper()
        df_hari = df_hari[~df_hari['NOPIN'].str.contains('TOTAL|GORO|JUMLAH|KETERANGAN|NAN|COLUMN', na=False)]
        df_hari = df_hari[df_hari['NOPIN'] != '']
        df_hari['NOPIN'] = df_hari['NOPIN'].apply(lambda x: x[:-2] if x.endswith('.0') else x)

        # Sinkronisasi jika master tersedia
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

    df_master = pd.concat(cleaned.values(), ignore_index=True)
    st.success(f"✅ {len(df_master)} baris data berhasil digabung dari {len(cleaned)} sheet.")

    # Konversi kolom tonase numerik
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col

    # Durasi jika kolom waktu ada
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
    df_kec = pd.DataFrame()
    if 'Kecamatan' in df_master.columns and col_netto:
        df_kec = df_master.groupby('Kecamatan').agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index()

    df_armada = pd.DataFrame()
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        df_armada = df_master.groupby(group_cols).agg(
            Total_Trip=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Tonase', ascending=False)

    df_tren = pd.DataFrame()
    if col_netto:
        df_tren = df_master.groupby('TANGGAL').agg(
            Total_Ritase=('NOPIN', 'count'), Total_Tonase=(col_netto, 'sum')
        ).reset_index()

    df_durasi = pd.DataFrame()
    if 'DURASI_MENIT' in df_master.columns and 'Kecamatan' in df_master.columns:
        df_durasi = df_master.dropna(subset=['DURASI_MENIT']).groupby('Kecamatan')['DURASI_MENIT'].mean().reset_index()

    df_jam = pd.DataFrame()
    if 'JAM_INPUT' in df_master.columns:
        df_jam = df_master.dropna(subset=['JAM_INPUT']).groupby('JAM_INPUT').size().reset_index(name='Jumlah')

    return df_master, df_kec, df_armada, df_tren, df_durasi, df_jam

# -------------------------- SESSION STATE --------------------------
if "hasil" not in st.session_state:
    st.session_state.hasil = None

# -------------------------- APLIKASI UTAMA --------------------------
def main():
    st.title("🚛 Dashboard DLH Armada – Analisis & Unduh Lengkap")
    st.markdown("Upload file Excel Anda. Sistem akan melewati data yang tidak valid dan menampilkan hasil analisis.")

    with st.sidebar:
        uploaded_file = st.file_uploader("📂 Pilih file Excel", type=["xlsx", "xls"])
        if uploaded_file:
            st.success("File siap diproses")
        if st.button("🚀 Mulai Proses", use_container_width=True):
            with st.spinner("Membaca dan memproses..."):
                sheets = baca_semua_sheet(uploaded_file)
                if sheets:
                    hasil = proses_data_utama(sheets, uploaded_file)
                    if hasil is not None:
                        st.session_state.hasil = hasil
                        st.balloons()
                    else:
                        st.session_state.hasil = None
                else:
                    st.error("Gagal membaca sheet. Pastikan file Excel valid.")

    if st.session_state.hasil is not None:
        df_master, df_kec, df_armada, df_tren, df_durasi, df_jam = st.session_state.hasil

        # Filter
        st.sidebar.markdown("---")
        st.sidebar.header("🔍 Filter")
        if 'Kecamatan' in df_master.columns:
            opts = ['Semua'] + sorted(df_master['Kecamatan'].dropna().unique().tolist())
            kec = st.sidebar.selectbox("Kecamatan", opts)
            if kec != 'Semua':
                df_master = df_master[df_master['Kecamatan'] == kec]
        if 'TANGGAL' in df_master.columns:
            tgls = sorted(df_master['TANGGAL'].unique())
            if len(tgls) > 1:
                rentang = st.sidebar.date_input("Rentang Tanggal", [pd.to_datetime(tgls[0]), pd.to_datetime(tgls[-1])])
                if len(rentang) == 2:
                    df_master = df_master[(pd.to_datetime(df_master['TANGGAL']) >= pd.Timestamp(rentang[0])) &
                                          (pd.to_datetime(df_master['TANGGAL']) <= pd.Timestamp(rentang[1]))]

        # Metrik
        total_trip = len(df_master)
        total_armada = df_master['NOPIN'].nunique()
        col_ton = cari_kolom(df_master.columns, ['NETTO']) or cari_kolom(df_master.columns, ['GROSS', 'TOTAL'])
        total_tonase = df_master[col_ton].sum()/1000 if col_ton else 0
        durasi_rata = df_master['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df_master.columns else 0

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Total Trip", total_trip)
        col2.metric("Armada Aktif", total_armada)
        col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
        col4.metric("Rata² Durasi (menit)", f"{durasi_rata:.1f}" if durasi_rata else "-")

        st.markdown("---")

        # Grafik
        if not df_tren.empty:
            fig1 = px.line(df_tren, x='TANGGAL', y='Total_Ritase', title='Tren Ritase Harian', markers=True)
            st.plotly_chart(fig1, use_container_width=True)
            st.session_state.fig_tren = fig1
        if not df_kec.empty:
            fig2 = px.bar(df_kec, x='Kecamatan', y='Total_Tonase', title='Total Tonase per Kecamatan',
                          color='Total_Tonase', color_continuous_scale='Viridis')
            st.plotly_chart(fig2, use_container_width=True)
            st.session_state.fig_kec = fig2
        if not df_jam.empty:
            fig3 = px.area(df_jam, x='JAM_INPUT', y='Jumlah', title='Pola Kedatangan per Jam')
            st.plotly_chart(fig3, use_container_width=True)
            st.session_state.fig_jam = fig3
        if not df_armada.empty:
            top10 = df_armada.head(10)
            fig4 = px.bar(top10, x='NOPIN', y='Total_Trip', title='10 Armada Teraktif', color='Total_Trip')
            st.plotly_chart(fig4, use_container_width=True)
            st.session_state.fig_top = fig4

        # Ringkasan
        st.subheader("📝 Ringkasan Eksekutif")
        kec_max = df_kec.iloc[0]['Kecamatan'] if not df_kec.empty else "N/A"
        st.markdown(f"""
- **Total trip:** {total_trip}
- **Armada aktif:** {total_armada}
- **Total tonase:** {total_tonase:,.1f} Ton
- **Kecamatan terpadat:** {kec_max}
        """)

        # Laporan AI
        if st.button("🤖 Buat Laporan AI"):
            laporan = laporan_ai(f"Trip:{total_trip}, Tonase:{total_tonase}, Armada:{total_armada}")
            if laporan:
                st.session_state.laporan_ai = laporan
                st.write(laporan)
            else:
                st.warning("Laporan AI tidak tersedia (cek API key atau koneksi).")

        # Unduh
        st.subheader("📥 Unduh Hasil")
        @st.cache_data
        def to_excel(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as w:
                df.to_excel(w, index=False)
            return output.getvalue()

        col_a, col_b, col_c = st.columns(3)
        with col_a:
            st.download_button("📊 Master Data", data=to_excel(df_master), file_name="Master_Data.xlsx")
            if not df_kec.empty:
                st.download_button("📊 Laporan Kecamatan", data=to_excel(df_kec), file_name="Kecamatan.xlsx")
            if not df_armada.empty:
                st.download_button("📊 Performa Armada", data=to_excel(df_armada), file_name="Performa_Armada.xlsx")
        with col_b:
            if not df_tren.empty:
                st.download_button("📈 Tren Harian", data=to_excel(df_tren), file_name="Tren_Harian.xlsx")
            if not df_durasi.empty:
                st.download_button("⏱️ Durasi per Kecamatan", data=to_excel(df_durasi), file_name="Durasi_Kec.xlsx")
            st.download_button("📝 Ringkasan", data=f"Ringkasan\nTrip:{total_trip}\nTonase:{total_tonase}".encode(), file_name="Ringkasan.txt")
        with col_c:
            pilihan = st.selectbox("Pilih grafik", ["Tren","Kecamatan","Jam","Top"])
            fig_dict = {'Tren':'fig_tren','Kecamatan':'fig_kec','Jam':'fig_jam','Top':'fig_top'}
            fig = st.session_state.get(fig_dict[pilihan])
            if fig:
                try:
                    img = pio.to_image(fig, format='png', scale=2)
                    st.download_button("📸 Unduh Grafik", data=img, file_name=f"{pilihan}.png", mime="image/png")
                except Exception as e:
                    st.warning(f"Gagal unduh grafik: {e}")
            if 'laporan_ai' in st.session_state:
                st.download_button("🤖 Laporan AI", data=st.session_state.laporan_ai, file_name="laporan_ai.txt")

        with st.expander("🔍 Lihat Data Mentah"):
            st.dataframe(df_master.head(200))

    else:
        st.info("👆 Upload file Excel (Test.xls) dan klik 'Mulai Proses'. Pastikan terdapat sheet 'List Armada' dan sheet harian bernama 1-30.")

if __name__ == "__main__":
    main()
