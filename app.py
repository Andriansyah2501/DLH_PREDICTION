import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO

# ---------- Konfigurasi ----------
st.set_page_config(page_title="Dashboard Armada DLH", page_icon="🚛", layout="wide")

# ---------- Fungsi Bantu ----------
def cari_kolom(kolom_list, kata_kunci):
    for col in kolom_list:
        col_up = str(col).upper()
        if any(kw in col_up for kw in kata_kunci):
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

# ---------- Fungsi per Tugas ----------
def tugas1_dan_2(sheets_dict, uploaded_file):
    """Tugas 1 & 2: Gabung 30 sheet & validasi/sinkronisasi dengan List Armada."""
    # Cari sheet List Armada
    armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
    if armada_sheet is None:
        armada_sheet = next((s for s in sheets_dict if 'armada' in s.lower()), None)

    ref_dict = {}
    if armada_sheet:
        xls = pd.ExcelFile(uploaded_file)
        df_ref = pd.read_excel(xls, sheet_name=armada_sheet, header=1)
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
    # Sheet harian
    daily_sheets = [s for s in sheets_dict if s.isdigit()]
    if not daily_sheets:
        daily_sheets = [s for s in sheets_dict if s != armada_sheet and s not in ['Tugas', 'Master Data']]

    cleaned = {}
    skipped = []
    for sheet in daily_sheets:
        try:
            df_raw = sheets_dict[sheet]
        except:
            skipped.append(sheet)
            continue
        # Deteksi header
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
        # Sinkronisasi
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
        cleaned[sheet] = df_hari

    if not cleaned:
        return None, None, None, skipped
    df_master = pd.concat(cleaned.values(), ignore_index=True)
    return df_master, ref_dict, cleaned, skipped

def tugas3_hitung_trip_tonase(df_master):
    """Hitung jumlah trip dan tonase per armada."""
    ton_col = cari_kolom(df_master.columns, ['NETTO', 'GROSS', 'TARE', 'BERAT'])
    if ton_col:
        df_master[ton_col] = pd.to_numeric(df_master[ton_col], errors='coerce').fillna(0)
    col_netto = cari_kolom(df_master.columns, ['NETTO']) or ton_col
    if col_netto:
        group_cols = ['NOPIN', 'NO_PLAT']
        for c in ['Kecamatan', 'TYPE', 'MERK']:
            if c in df_master.columns:
                group_cols.append(c)
        df_armada = df_master.groupby(group_cols).agg(
            Total_Trip=('NOPIN', 'count'),
            Total_Tonase=(col_netto, 'sum')
        ).reset_index().sort_values('Total_Trip', ascending=False)
    else:
        df_armada = pd.DataFrame()
    return df_armada, col_netto

def tugas4_teraktif_tidakefisien(df_armada):
    if df_armada.empty:
        return None, None
    teraktif = df_armada.iloc[0]
    tidak_efisien = df_armada[df_armada['Total_Trip'] > 0].iloc[-1] if any(df_armada['Total_Trip'] > 0) else df_armada.iloc[-1]
    return teraktif, tidak_efisien

def tugas5_waktu_per_jenis(df_master):
    col_masuk = cari_kolom(df_master.columns, ['MASUK', 'JAM_1', 'TIMBANG1'])
    col_keluar = cari_kolom(df_master.columns, ['KELUAR', 'JAM_2', 'TIMBANG2'])
    if col_masuk and col_keluar:
        df_master['WAKTU_MASUK_DT'] = pd.to_datetime(df_master[col_masuk], format='%H:%M:%S', errors='coerce')
        df_master['WAKTU_KELUAR_DT'] = pd.to_datetime(df_master[col_keluar], format='%H:%M:%S', errors='coerce')
        df_master['DURASI_MENIT'] = (df_master['WAKTU_KELUAR_DT'] - df_master['WAKTU_MASUK_DT']).dt.total_seconds() / 60
        df_master = df_master[(df_master['DURASI_MENIT'] >= 0) & (df_master['DURASI_MENIT'] <= 300)]
        if 'TYPE' in df_master.columns:
            df_waktu = df_master.dropna(subset=['DURASI_MENIT']).groupby('TYPE')['DURASI_MENIT'].mean().reset_index()
            df_waktu.columns = ['Jenis Armada', 'Rata2 Waktu Tempuh (menit)']
            return df_waktu
    return pd.DataFrame()

# ---------- Session State ----------
if "step" not in st.session_state:
    st.session_state.step = 0
    st.session_state.df_master = None
    st.session_state.df_armada = None
    st.session_state.df_kec = None
    st.session_state.df_tren = None
    st.session_state.df_waktu = None
    st.session_state.ref_dict = None
    st.session_state.skipped = []
    st.session_state.teraktif = None
    st.session_state.tidak_efisien = None
    st.session_state.sheets_dict = None
    st.session_state.cleaned_sheets = None
    st.session_state.col_netto = None

# ---------- Aplikasi Utama ----------
def main():
    st.title("🚛 Dashboard Armada DLH – Proses Bertahap")
    st.markdown("### Ikuti langkah-langkah di bawah untuk melihat setiap proses analisis.")

    uploaded_file = st.file_uploader("📂 Unggah file Excel (.xls/.xlsx)", type=["xlsx", "xls"], key="file_upload")

    if uploaded_file:
        # Tombol mulai/reset
        if st.button("🔄 Mulai / Ulangi Proses"):
            st.session_state.step = 1
            st.session_state.sheets_dict = baca_semua_sheet(uploaded_file)
            st.rerun()

    # Progres bertahap berdasarkan step
    if st.session_state.step >= 1 and st.session_state.sheets_dict is not None:
        sheets_dict = st.session_state.sheets_dict

        # ---------- Langkah 1: Membaca file ----------
        with st.expander("📋 Langkah 1: Membaca File Excel", expanded=(st.session_state.step==1)):
            st.write(f"**Jumlah sheet ditemukan:** {len(sheets_dict)}")
            st.write("**Nama-nama sheet:**")
            st.write(list(sheets_dict.keys()))
            # Cari armada sheet
            armada_sheet = next((s for s in sheets_dict if 'list armada' in s.lower()), None)
            daily_candidates = [s for s in sheets_dict if s.isdigit()]
            st.write(f"**Sheet List Armada terdeteksi:** {armada_sheet}")
            st.write(f"**Sheet harian (digit) terdeteksi:** {daily_candidates} ({len(daily_candidates)} sheet)")
            if st.button("Lanjut ke Langkah 2", key="to_step2"):
                st.session_state.step = 2
                st.rerun()

        # ---------- Langkah 2: Proses Gabung & Sinkronisasi ----------
        if st.session_state.step >= 2:
            with st.expander("🔧 Langkah 2: Menggabungkan Sheet & Sinkronisasi Data", expanded=(st.session_state.step==2)):
                if st.session_state.df_master is None:
                    with st.spinner("Memproses..."):
                        df_master, ref_dict, cleaned, skipped = tugas1_dan_2(sheets_dict, uploaded_file)
                        st.session_state.df_master = df_master
                        st.session_state.ref_dict = ref_dict
                        st.session_state.cleaned_sheets = cleaned
                        st.session_state.skipped = skipped
                if st.session_state.df_master is not None:
                    st.success(f"✅ Master Data berhasil dibuat: {len(st.session_state.df_master)} baris dari {len(st.session_state.cleaned_sheets)} sheet.")
                    st.write(f"Sheet dilewati: {st.session_state.skipped}")
                    st.write("**5 Baris pertama Master Data:**")
                    st.dataframe(st.session_state.df_master.head())
                if st.button("Lanjut ke Langkah 3", key="to_step3"):
                    st.session_state.step = 3
                    st.rerun()

        # ---------- Langkah 3: Hitung Trip & Tonase ----------
        if st.session_state.step >= 3 and st.session_state.df_master is not None:
            with st.expander("🧮 Langkah 3: Menghitung Jumlah Trip & Tonase per Armada", expanded=(st.session_state.step==3)):
                if st.session_state.df_armada is None:
                    df_armada, col_netto = tugas3_hitung_trip_tonase(st.session_state.df_master)
                    st.session_state.df_armada = df_armada
                    st.session_state.col_netto = col_netto
                if not st.session_state.df_armada.empty:
                    st.write("**10 Armada dengan Trip Terbanyak:**")
                    st.dataframe(st.session_state.df_armada.head(10))
                else:
                    st.warning("Tidak ada kolom tonase yang valid.")
                if st.button("Lanjut ke Langkah 4", key="to_step4"):
                    st.session_state.step = 4
                    st.rerun()

        # ---------- Langkah 4: Armada Teraktif & Tidak Efisien ----------
        if st.session_state.step >= 4 and st.session_state.df_armada is not None:
            with st.expander("🏆 Langkah 4: Identifikasi Armada Teraktif & Paling Tidak Efisien", expanded=(st.session_state.step==4)):
                if st.session_state.teraktif is None:
                    teraktif, tidak_efisien = tugas4_teraktif_tidakefisien(st.session_state.df_armada)
                    st.session_state.teraktif = teraktif
                    st.session_state.tidak_efisien = tidak_efisien
                if st.session_state.teraktif is not None:
                    col1, col2 = st.columns(2)
                    with col1:
                        st.success(f"🥇 **Armada Teraktif:** {st.session_state.teraktif['NOPIN']} ({st.session_state.teraktif['NO_PLAT']}) - {int(st.session_state.teraktif['Total_Trip'])} trip")
                    with col2:
                        st.error(f"🐌 **Armada Tidak Efisien:** {st.session_state.tidak_efisien['NOPIN']} ({st.session_state.tidak_efisien['NO_PLAT']}) - {int(st.session_state.tidak_efisien['Total_Trip'])} trip")
                if st.button("Lanjut ke Langkah 5", key="to_step5"):
                    st.session_state.step = 5
                    st.rerun()

        # ---------- Langkah 5: Rata-rata Waktu Tempuh per Jenis Armada ----------
        if st.session_state.step >= 5 and st.session_state.df_master is not None:
            with st.expander("⏱️ Langkah 5: Rata‑rata Waktu Tempuh per Jenis Armada", expanded=(st.session_state.step==5)):
                if st.session_state.df_waktu is None:
                    df_waktu = tugas5_waktu_per_jenis(st.session_state.df_master)
                    st.session_state.df_waktu = df_waktu
                if not st.session_state.df_waktu.empty:
                    st.dataframe(st.session_state.df_waktu.style.format({'Rata2 Waktu Tempuh (menit)': '{:.1f}'}))
                else:
                    st.info("Data waktu masuk/keluar tidak tersedia, sehingga langkah ini tidak dapat dilakukan.")
                if st.button("Lanjut ke Langkah 6", key="to_step6"):
                    st.session_state.step = 6
                    st.rerun()

        # ---------- Langkah 6: Visualisasi ----------
        if st.session_state.step >= 6 and st.session_state.df_master is not None:
            with st.expander("📊 Langkah 6: Grafik Interaktif", expanded=(st.session_state.step==6)):
                df = st.session_state.df_master
                # Tren Harian
                if st.session_state.col_netto:
                    tren = df.groupby('TANGGAL').size().reset_index(name='Total_Ritase')
                    fig1 = px.line(tren, x='TANGGAL', y='Total_Ritase', title='Tren Ritase Harian', markers=True)
                    st.plotly_chart(fig1, use_container_width=True)
                    # Distribusi Kecamatan
                    if 'Kecamatan' in df.columns:
                        kec = df.groupby('Kecamatan')[st.session_state.col_netto].sum().reset_index(name='Total_Tonase')
                        fig2 = px.bar(kec.sort_values('Total_Tonase', ascending=False), x='Kecamatan', y='Total_Tonase', color='Total_Tonase', color_continuous_scale='Viridis', title='Total Tonase per Kecamatan')
                        st.plotly_chart(fig2, use_container_width=True)
                    # Top Armada
                    if st.session_state.df_armada is not None and not st.session_state.df_armada.empty:
                        top10 = st.session_state.df_armada.head(10)
                        fig3 = px.bar(top10, x='NOPIN', y='Total_Trip', color='Total_Trip', title='10 Armada Teraktif')
                        st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.warning("Data tonase tidak tersedia, grafik tidak dapat dibuat.")
                if st.button("Lanjut ke Langkah 7", key="to_step7"):
                    st.session_state.step = 7
                    st.rerun()

        # ---------- Langkah 7: Laporan Ringkasan ----------
        if st.session_state.step >= 7:
            st.subheader("📝 Langkah 7: Laporan Ringkasan & Unduh")
            df = st.session_state.df_master
            total_trip = len(df)
            total_armada = df['NOPIN'].nunique()
            total_tonase = df[st.session_state.col_netto].sum()/1000 if st.session_state.col_netto else 0
            durasi_rata = df['DURASI_MENIT'].mean() if 'DURASI_MENIT' in df.columns else None

            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Total Trip", total_trip)
            col2.metric("Armada Aktif", total_armada)
            col3.metric("Total Tonase (Ton)", f"{total_tonase:,.1f}")
            col4.metric("Rata² Durasi (menit)", f"{durasi_rata:.1f}" if durasi_rata else "-")

            laporan = f"""
**Ringkasan Operasional:**
- Total trip: {total_trip}
- Armada aktif: {total_armada} unit
- Total tonase: {total_tonase:,.1f} Ton
- Armada teraktif: {st.session_state.teraktif['NOPIN'] if st.session_state.teraktif is not None else '-'}
- Armada tidak efisien: {st.session_state.tidak_efisien['NOPIN'] if st.session_state.tidak_efisien is not None else '-'}
            """
            st.markdown(laporan)
            st.download_button("📄 Unduh Ringkasan (TXT)", laporan.encode('utf-8'), "ringkasan.txt")

            # Unduhan data
            @st.cache_data
            def to_excel(dataframe):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as w:
                    dataframe.to_excel(w, index=False)
                return output.getvalue()

            if st.session_state.df_armada is not None and not st.session_state.df_armada.empty:
                st.download_button("📊 Unduh Statistik Armada (Excel)", to_excel(st.session_state.df_armada), "statistik_armada.xlsx")

    elif uploaded_file is None:
        st.info("👆 Silakan unggah file Excel untuk memulai proses analisis.")

if __name__ == "__main__":
    main()
