import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
from io import BytesIO

# ---------- Fungsi Bantu ----------
def cari_kolom(kolom_list, kata_kunci):
    for col in kolom_list:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file Excel...")
def baca_semua_sheet(uploaded_file):
    """Baca semua sheet sekali saja. Return dict {nama_sheet: dataframe}."""
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

# ---------- Fungsi Proses Data (Tugas 1 & 2) ----------
def gabung_dan_sinkron(sheets_dict):
    """
    Menggabungkan sheet harian dan sinkronisasi dengan List Armada.
    Return (df_master, ref_dict, cleaned, skipped, pesan_error)
    """
    # 1. Cari sheet List Armada
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)

    ref_dict = {}
    if armada_sheet:
        # Gunakan dataframe dari sheets_dict, asumsikan header di baris 1 (index 1)
        df_ref_raw = sheets_dict[armada_sheet].copy()
        if len(df_ref_raw) > 1:
            # Baris pertama sebagai header
            new_header = df_ref_raw.iloc[0]  # ambil baris indeks 0
            df_ref = df_ref_raw[1:]          # data mulai baris 1
            df_ref.columns = [str(c).strip().upper() for c in new_header]
        else:
            # fallback jika hanya 1 baris
            df_ref = df_ref_raw.copy()
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

    # 2. Proses sheet harian (digit)
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]
    if not daily_sheets:
        return None, None, None, [], "Tidak ada sheet harian."

    cleaned = {}
    skipped = []
    for sheet in daily_sheets:
        try:
            df_raw = sheets_dict[sheet].copy()
        except:
            skipped.append(sheet)
            continue

        # Deteksi header: cari baris yang mengandung "PINTU" atau "PLAT MOBIL" atau "NOPIN"
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = " ".join(row.astype(str).dropna().str.upper().values)
            if 'PINTU' in row_str or 'PLAT MOBIL' in row_str or 'NOPIN' in row_str:
                header_idx = idx
                break
        if header_idx is None:
            skipped.append(sheet)
            continue

        # Potong dataframe: baris sebelum header dibuang, header jadi nama kolom
        try:
            # Data setelah header
            df_hari = df_raw.iloc[header_idx+1:].reset_index(drop=True)
            # Nama kolom dari baris header
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

        # Sinkronisasi dengan master
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

        # Tambah kolom tanggal
        try:
            tgl = f"2026-06-{int(sheet):02d}"
        except:
            tgl = sheet
        df_hari['TANGGAL'] = tgl

        cleaned[sheet] = df_hari

    if not cleaned:
        return None, ref_dict, None, skipped, "Tidak ada sheet yang valid."

    df_master = pd.concat(cleaned.values(), ignore_index=True)
    return df_master, ref_dict, cleaned, skipped, None

# ---------- Fungsi Analisis ----------
def hitung_trip_tonase(df_master):
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col
    if not col_netto:
        return pd.DataFrame(), None
    group_cols = ['NOPIN', 'NO_PLAT']
    for c in ['Kecamatan', 'TYPE', 'MERK']:
        if c in df_master.columns:
            group_cols.append(c)
    df_armada = df_master.groupby(group_cols).agg(
        Total_Trip=('NOPIN', 'count'),
        Total_Tonase=(col_netto, 'sum')
    ).reset_index().sort_values('Total_Trip', ascending=False)
    return df_armada, col_netto

def armada_ekstrem(df_armada):
    if df_armada.empty:
        return None, None
    teraktif = df_armada.iloc[0]
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if any(df_armada['Total_Trip'] > 0) else df_armada.iloc[-1]
    return teraktif, tidak_efisien

def waktu_per_jenis(df_master):
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['KELUAR_DT'] - df_master['MASUK_DT']).dt.total_seconds() / 60
        df_valid = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        if 'TYPE' in df_valid.columns:
            df_waktu = df_valid.groupby('TYPE')['DURASI_MENIT'].mean().reset_index()
            df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']
            return df_waktu
    return pd.DataFrame()

# ---------- Session State ----------
for key, default in [("step", 0), ("sheets_dict", None), ("df_master", None),
                     ("df_armada", None), ("col_netto", None), ("teraktif", None),
                     ("tidak_efisien", None), ("df_waktu", None), ("ref_dict", None),
                     ("cleaned", None), ("skipped", None)]:
    if key not in st.session_state:
        st.session_state[key] = default

# ---------- Aplikasi Utama ----------
def main():
    st.title("🚛 Dashboard Armada DLH – Proses Bertahap")
    st.markdown("Unggah file Excel, lalu klik **Mulai/Ulangi Proses**. Ikuti langkah-langkah untuk melihat setiap tahap analisis.")

    uploaded_file = st.file_uploader("📂 Pilih file Excel (.xls/.xlsx)", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🔄 Mulai / Ulangi Proses"):
            st.session_state.sheets_dict = baca_semua_sheet(uploaded_file)
            st.session_state.step = 1
            # reset hasil
            for key in ["df_master", "df_armada", "col_netto", "teraktif", "tidak_efisien", "df_waktu", "ref_dict", "cleaned", "skipped"]:
                st.session_state[key] = None
            st.rerun()

    if st.session_state.sheets_dict is not None:
        sheets = st.session_state.sheets_dict

        # ---- Langkah 1 ----
        with st.expander("📋 Langkah 1: Membaca File Excel", expanded=(st.session_state.step == 1)):
            st.write(f"**Jumlah sheet ditemukan:** {len(sheets)}")
            st.write("**Nama sheet:**", list(sheets.keys()))
            armada_sheet = next((s for s in sheets if 'list armada' in s.lower()), None)
            daily = [s for s in sheets if s.isdigit()]
            st.write(f"**Sheet List Armada:** {armada_sheet}")
            st.write(f"**Sheet harian (digit):** {daily} ({len(daily)} sheet)")
            if st.button("➡ Lanjut ke Langkah 2", key="to2"):
                st.session_state.step = 2
                st.rerun()

        # ---- Langkah 2 ----
        if st.session_state.step >= 2:
            with st.expander("🔧 Langkah 2: Gabung & Sinkronisasi Data (Tugas 1 & 2)", expanded=(st.session_state.step == 2)):
                if st.session_state.df_master is None:
                    with st.spinner("Memproses..."):
                        df_master, ref_dict, cleaned, skipped, error_msg = gabung_dan_sinkron(sheets)
                        if error_msg:
                            st.error(error_msg)
                        else:
                            st.session_state.df_master = df_master
                            st.session_state.ref_dict = ref_dict
                            st.session_state.cleaned = cleaned
                            st.session_state.skipped = skipped
                if st.session_state.df_master is not None:
                    st.success(f"✅ Master Data: {len(st.session_state.df_master)} baris dari {len(st.session_state.cleaned)} sheet.")
                    st.write(f"Sheet dilewati: {st.session_state.skipped}")
                    st.write("**5 Baris pertama:**")
                    st.dataframe(st.session_state.df_master.head())
                    if st.button("➡ Lanjut ke Langkah 3", key="to3"):
                        st.session_state.step = 3
                        st.rerun()
                else:
                    st.warning("Gagal membuat Master Data. Periksa format file.")

        # ---- Langkah 3 ----
        if st.session_state.step >= 3 and st.session_state.df_master is not None:
            with st.expander("🧮 Langkah 3: Hitung Trip & Tonase per Armada (Tugas 3)", expanded=(st.session_state.step == 3)):
                if st.session_state.df_armada is None:
                    df_armada, col_netto = hitung_trip_tonase(st.session_state.df_master)
                    st.session_state.df_armada = df_armada
                    st.session_state.col_netto = col_netto
                if not st.session_state.df_armada.empty:
                    st.dataframe(st.session_state.df_armada.head(10))
                    if st.button("➡ Lanjut ke Langkah 4", key="to4"):
                        st.session_state.step = 4
                        st.rerun()
                else:
                    st.warning("Kolom tonase tidak ditemukan.")

        # ---- Langkah 4 ----
        if st.session_state.step >= 4 and st.session_state.df_armada is not None and not st.session_state.df_armada.empty:
            with st.expander("🏆 Langkah 4: Armada Teraktif & Tidak Efisien (Tugas 4)", expanded=(st.session_state.step == 4)):
                if st.session_state.teraktif is None:
                    teraktif, tidak_efisien = armada_ekstrem(st.session_state.df_armada)
                    st.session_state.teraktif = teraktif
                    st.session_state.tidak_efisien = tidak_efisien
                if st.session_state.teraktif is not None:
                    col1, col2 = st.columns(2)
                    col1.success(f"🥇 Teraktif: {st.session_state.teraktif['NOPIN']} ({st.session_state.teraktif['NO_PLAT']}) - {int(st.session_state.teraktif['Total_Trip'])} trip")
                    col2.error(f"🐌 Tidak Efisien: {st.session_state.tidak_efisien['NOPIN']} ({st.session_state.tidak_efisien['NO_PLAT']}) - {int(st.session_state.tidak_efisien['Total_Trip'])} trip")
                    if st.button("➡ Lanjut ke Langkah 5", key="to5"):
                        st.session_state.step = 5
                        st.rerun()

        # ---- Langkah 5 ----
        if st.session_state.step >= 5 and st.session_state.df_master is not None:
            with st.expander("⏱️ Langkah 5: Rata‑rata Waktu Tempuh per Jenis Armada (Tugas 5)", expanded=(st.session_state.step == 5)):
                if st.session_state.df_waktu is None:
                    st.session_state.df_waktu = waktu_per_jenis(st.session_state.df_master)
                if not st.session_state.df_waktu.empty:
                    st.dataframe(st.session_state.df_waktu.style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}))
                else:
                    st.info("Data waktu tidak tersedia.")
                if st.button("➡ Lanjut ke Langkah 6", key="to6"):
                    st.session_state.step = 6
                    st.rerun()

        # ---- Langkah 6 ----
        if st.session_state.step >= 6 and st.session_state.df_master is not None:
            with st.expander("📊 Langkah 6: Grafik Interaktif (Tugas 6)", expanded=(st.session_state.step == 6)):
                df = st.session_state.df_master
                col_netto = st.session_state.col_netto
                if col_netto:
                    # Tren harian
                    tren = df.groupby('TANGGAL').size().reset_index(name='Ritase')
                    fig1 = px.line(tren, x='TANGGAL', y='Ritase', title='Tren Ritase Harian', markers=True)
                    st.plotly_chart(fig1, use_container_width=True)
                    # Kecamatan
                    if 'Kecamatan' in df.columns:
                        kec = df.groupby('Kecamatan')[col_netto].sum().reset_index(name='Tonase')
                        fig2 = px.bar(kec.sort_values('Tonase', ascending=False), x='Kecamatan', y='Tonase', color='Tonase', color_continuous_scale='Viridis')
                        st.plotly_chart(fig2, use_container_width=True)
                    # Top Armada
                    if st.session_state.df_armada is not None and not st.session_state.df_armada.empty:
                        top10 = st.session_state.df_armada.head(10)
                        fig3 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip', title='10 Armada Teraktif')
                        st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.warning("Tonase tidak tersedia, grafik terbatas.")
                if st.button("➡ Lanjut ke Langkah 7", key="to7"):
                    st.session_state.step = 7
                    st.rerun()

        # ---- Langkah 7 ----
        if st.session_state.step >= 7:
            st.subheader("📝 Langkah 7: Laporan & Unduh")
            df = st.session_state.df_master
            total_trip = len(df)
            total_armada = df['NOPIN'].nunique()
            total_tonase = df[st.session_state.col_netto].sum()/1000 if st.session_state.col_netto else 0
            durasi = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else None

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Trip", total_trip)
            col2.metric("Armada Aktif", total_armada)
            col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
            col4.metric("Rata² Durasi", f"{durasi:.1f} menit" if durasi else "-")

            laporan = f"""Ringkasan:
- Trip: {total_trip}
- Armada: {total_armada}
- Tonase: {total_tonase:,.1f} Ton
- Teraktif: {st.session_state.teraktif['NOPIN'] if st.session_state.teraktif is not None else '-'}
- Tidak Efisien: {st.session_state.tidak_efisien['NOPIN'] if st.session_state.tidak_efisien is not None else '-'}
"""
            st.markdown(laporan)
            st.download_button("📄 Unduh Ringkasan (TXT)", laporan, "ringkasan.txt")

            @st.cache_data
            def to_excel(dataframe):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as w:
                    dataframe.to_excel(w, index=False)
                return output.getvalue()

            if st.session_state.df_armada is not None and not st.session_state.df_armada.empty:
                st.download_button("📊 Unduh Statistik Armada", to_excel(st.session_state.df_armada), "statistik_armada.xlsx")
            if st.session_state.df_master is not None:
                st.download_button("📊 Unduh Master Data", to_excel(st.session_state.df_master), "master_data.xlsx")

    else:
        st.info("👆 Silakan unggah file dan klik 'Mulai / Ulangi Proses'.")

if __name__ == "__main__":
    main()
