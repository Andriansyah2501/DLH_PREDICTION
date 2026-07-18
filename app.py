import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
import os
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Dashboard Armada", layout="wide")
st.title("🚛 Sistem Rekomendasi & Dashboard Visualisasi Armada")
st.markdown("Unggah file Excel berisi **List Armada** (1 sheet) dan **Data Harian** (29 sheet) untuk memulai analisis.")

# ================== 1. KONFIGURASI API DEEPSEEK ==================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "sk-796de1893bc044ecaa57b02090683f12")  # ganti dengan key Anda
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def generate_report(stats_text):
    """Menghasilkan laporan singkat dari statistik menggunakan DeepSeek."""
    if not DEEPSEEK_API_KEY or DEEPSEEK_API_KEY.startswith("sk-xxx"):
        return "⚠️ API Key DeepSeek belum diatur. Laporan tidak dapat dibuat."

    headers = {
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}",
        "Content-Type": "application/json"
    }
    prompt = f"""
Anda adalah asisten analis data armada truk. Berikut ringkasan statistik armada:
{stats_text}

Buatlah laporan singkat dalam bahasa Indonesia yang berisi:
- Gambaran umum performa armada.
- Armada paling aktif dan paling tidak efisien beserta kemungkinan penyebabnya.
- Rekomendasi perbaikan berdasarkan data (misal: penjadwalan ulang, perawatan, rute).
Tulis dengan gaya profesional namun mudah dipahami. Panjang maksimal 3 paragraf.
    """
    payload = {
        "model": "deepseek-chat",
        "messages": [{"role": "user", "content": prompt}],
        "temperature": 0.7,
        "max_tokens": 500
    }
    try:
        resp = requests.post(DEEPSEEK_URL, headers=headers, json=payload, timeout=30)
        resp.raise_for_status()
        return resp.json()["choices"][0]["message"]["content"]
    except Exception as e:
        return f"❌ Gagal membuat laporan: {str(e)}"

# ================== 2. BACA & GABUNG DATA ==================
@st.cache_data(show_spinner="Memproses file Excel...")
def load_and_merge(uploaded_file):
    """
    Membaca semua sheet dari file Excel.
    Sheet pertama dianggap sebagai LIST ARMADA, sisanya sheet HARIAN.
    """
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names

    if len(sheet_names) < 2:
        st.error("File minimal harus memiliki 2 sheet: List Armada dan 1 sheet harian.")
        return None, None

    # Sheet pertama sebagai master armada
    df_armada = pd.read_excel(xls, sheet_name=sheet_names[0])
    # 29 sheet berikutnya (atau semua kecuali pertama) adalah data harian
    daily_sheets = sheet_names[1:]

    list_daily = []
    for sheet in daily_sheets:
        try:
            df = pd.read_excel(xls, sheet_name=sheet)
            # Tambah kolom sumber sheet (bisa berupa tanggal jika nama sheet adalah tanggal)
            df["Sumber_Sheet"] = sheet
            list_daily.append(df)
        except Exception as e:
            st.warning(f"Sheet '{sheet}' dilewati karena error: {e}")

    if not list_daily:
        st.error("Tidak ada sheet harian yang bisa dibaca.")
        return None, None

    df_daily_raw = pd.concat(list_daily, ignore_index=True)
    return df_armada, df_daily_raw

# ================== 3. VALIDASI & PEMBERSIHAN ==================
def clean_and_validate(df_armada, df_daily):
    """
    Menyamakan No. Pintu dan Plat Mobil berdasarkan List Armada.
    - Jika No. Pintu di daily cocok dengan armada, pastikan Plat sesuai.
    - Jika Plat cocok, pastikan No. Pintu sesuai.
    - Data yang tidak bisa dicocokkan akan diberi flag.
    """
    # Pastikan kolom kunci ada (case insensitive)
    col_armada = {c.lower(): c for c in df_armada.columns}
    col_daily = {c.lower(): c for c in df_daily.columns}

    required_armada = ["no. pintu", "plat mobil"]
    required_daily = ["no. pintu", "plat mobil", "tonase", "waktu berangkat", "waktu tiba"]

    for req in required_armada:
        if req not in col_armada:
            st.error(f"Kolom '{req}' tidak ditemukan di List Armada.")
            return None
    for req in required_daily:
        if req not in col_daily:
            st.error(f"Kolom '{req}' tidak ditemukan di data harian.")
            return None

    # Ambil nama kolom asli
    no_pintu_arm = col_armada["no. pintu"]
    plat_arm = col_armada["plat mobil"]
    no_pintu_daily = col_daily["no. pintu"]
    plat_daily = col_daily["plat mobil"]
    tonase_col = col_daily["tonase"]
    waktu_berangkat_col = col_daily["waktu berangkat"]
    waktu_tiba_col = col_daily["waktu tiba"]

    # Buat lookup dari armada: no.pintu -> plat, plat -> no.pintu
    lookup_pintu_to_plat = df_armada.set_index(no_pintu_arm)[plat_arm].to_dict()
    lookup_plat_to_pintu = df_armada.set_index(plat_arm)[no_pintu_arm].to_dict()

    # Fungsi normalisasi string (hapus spasi, uppercase)
    def normalize(s):
        if pd.isna(s):
            return ""
        return str(s).strip().upper()

    df_daily[no_pintu_daily] = df_daily[no_pintu_daily].apply(normalize)
    df_daily[plat_daily] = df_daily[plat_daily].apply(normalize)

    # Salin untuk diupdate
    df_clean = df_daily.copy()
    mismatch_count = 0

    for idx, row in df_clean.iterrows():
        pintu = row[no_pintu_daily]
        plat = row[plat_daily]

        # Cek apakah no.pintu ada di master
        if pintu in lookup_pintu_to_plat:
            plat_master = normalize(lookup_pintu_to_plat[pintu])
            if plat != plat_master:
                # perbaiki plat
                df_clean.at[idx, plat_daily] = plat_master
                mismatch_count += 1
        # Jika plat ada di master, perbaiki no.pintu
        elif plat in lookup_plat_to_pintu:
            pintu_master = normalize(lookup_plat_to_pintu[plat])
            if pintu != pintu_master:
                df_clean.at[idx, no_pintu_daily] = pintu_master
                mismatch_count += 1
        else:
            # Tidak ditemukan di master, tandai tapi tetap disertakan (bisa jadi data baru)
            # bisa juga dihapus, tapi biarkan saja dengan flag
            pass

    # Gabungkan dengan informasi jenis armada dari master
    if "jenis armada" in col_armada:
        jenis_col = col_armada["jenis armada"]
        # Mapping no.pintu -> jenis
        map_jenis = df_armada.set_index(no_pintu_arm)[jenis_col].to_dict()
        df_clean["Jenis Armada"] = df_clean[no_pintu_daily].map(map_jenis).fillna("Tidak Diketahui")
    else:
        df_clean["Jenis Armada"] = "Tidak Diketahui"

    # Parsing waktu & hitung waktu tempuh (dalam jam)
    try:
        df_clean[waktu_berangkat_col] = pd.to_datetime(df_clean[waktu_berangkat_col], errors='coerce')
        df_clean[waktu_tiba_col] = pd.to_datetime(df_clean[waktu_tiba_col], errors='coerce')
        df_clean["Waktu Tempuh (jam)"] = (df_clean[waktu_tiba_col] - df_clean[waktu_berangkat_col]).dt.total_seconds() / 3600
        # Hapus yang negatif atau aneh
        df_clean.loc[df_clean["Waktu Tempuh (jam)"] < 0, "Waktu Tempuh (jam)"] = np.nan
    except Exception as e:
        st.warning(f"Gagal menghitung waktu tempuh: {e}")
        df_clean["Waktu Tempuh (jam)"] = np.nan

    st.success(f"✅ Data berhasil dibersihkan. Jumlah ketidaksesuaian yang diperbaiki: {mismatch_count} baris.")
    return df_clean

# ================== 4. PERHITUNGAN METRIK ==================
def compute_metrics(df):
    # Jumlah trip per armada (berdasarkan No. Pintu)
    trip_per_armada = df.groupby("No. Pintu").size().reset_index(name="Jumlah Trip")
    # Total tonase per armada
    tonase_per_armada = df.groupby("No. Pintu")["Tonase"].sum().reset_index(name="Total Tonase")
    # Gabungkan
    stat_armada = pd.merge(trip_per_armada, tonase_per_armada, on="No. Pintu", how="outer").fillna(0)

    # Armada teraktif
    most_active = stat_armada.loc[stat_armada["Jumlah Trip"].idxmax()]
    # Armada paling sedikit trip (dari yang ada data, atau bisa nol jika ada armada master tanpa trip)
    # Kita ambil dari data yang ada
    least_active = stat_armada.loc[stat_armada["Jumlah Trip"].idxmin()]

    # Rata-rata waktu tempuh per jenis armada
    avg_time_per_jenis = df.groupby("Jenis Armada")["Waktu Tempuh (jam)"].mean().reset_index()
    avg_time_per_jenis.columns = ["Jenis Armada", "Rata-rata Waktu Tempuh (jam)"]

    return stat_armada, most_active, least_active, avg_time_per_jenis

# ================== 5. VISUALISASI ==================
def create_visualizations(df, stat_armada, avg_time_per_jenis):
    # 1. Bar chart jumlah trip per armada
    fig1 = px.bar(stat_armada.sort_values("Jumlah Trip", ascending=False).head(20),
                  x="No. Pintu", y="Jumlah Trip",
                  title="20 Armada dengan Trip Terbanyak",
                  color="Jumlah Trip", color_continuous_scale="Viridis")
    fig1.update_layout(xaxis_tickangle=-45)

    # 2. Total tonase per armada
    fig2 = px.bar(stat_armada.sort_values("Total Tonase", ascending=False).head(20),
                  x="No. Pintu", y="Total Tonase",
                  title="20 Armada dengan Total Tonase Terbesar",
                  color="Total Tonase", color_continuous_scale="OrRd")
    fig2.update_layout(xaxis_tickangle=-45)

    # 3. Rata-rata waktu tempuh per jenis armada
    fig3 = px.bar(avg_time_per_jenis, x="Jenis Armada", y="Rata-rata Waktu Tempuh (jam)",
                  title="Rata-rata Waktu Tempuh per Jenis Armada",
                  color="Rata-rata Waktu Tempuh (jam)", color_continuous_scale="Blues")

    # 4. Distribusi trip per jenis armada (pie)
    jenis_dist = df["Jenis Armada"].value_counts().reset_index()
    jenis_dist.columns = ["Jenis Armada", "Jumlah Trip"]
    fig4 = px.pie(jenis_dist, names="Jenis Armada", values="Jumlah Trip",
                  title="Distribusi Trip per Jenis Armada")

    return fig1, fig2, fig3, fig4

# ================== 6. APLIKASI UTAMA ==================
uploaded_file = st.file_uploader("📂 Upload file Excel (format .xlsx)", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("Membaca dan memproses..."):
        df_armada, df_daily_raw = load_and_merge(uploaded_file)

    if df_armada is not None and df_daily_raw is not None:
        # Tampilkan data mentah opsional
        with st.expander("🔎 Lihat Data Mentah"):
            st.subheader("List Armada")
            st.dataframe(df_armada)
            st.subheader("Data Harian (Gabungan)")
            st.dataframe(df_daily_raw.head(100))

        # Bersihkan & validasi
        df_clean = clean_and_validate(df_armada, df_daily_raw)

        if df_clean is not None:
            st.subheader("📋 Master Data (Setelah Dibersihkan)")
            st.dataframe(df_clean.head(200))

            # Metrik
            stat_armada, most_active, least_active, avg_time_per_jenis = compute_metrics(df_clean)

            col1, col2, col3 = st.columns(3)
            col1.metric("Total Trip", len(df_clean))
            col2.metric("Total Tonase", f"{df_clean['Tonase'].sum():,.2f}")
            col3.metric("Rata-rata Waktu Tempuh (jam)", f"{df_clean['Waktu Tempuh (jam)'].mean():.2f}")

            st.subheader("🏆 Armada Teraktif")
            st.write(f"**No. Pintu:** {most_active['No. Pintu']}  |  **Trip:** {int(most_active['Jumlah Trip'])}  |  **Total Tonase:** {most_active['Total Tonase']:,.2f}")

            st.subheader("⚠️ Armada Paling Tidak Efisien (paling sedikit trip)")
            st.write(f"**No. Pintu:** {least_active['No. Pintu']}  |  **Trip:** {int(least_active['Jumlah Trip'])}  |  **Total Tonase:** {least_active['Total Tonase']:,.2f}")

            # Grafik interaktif
            fig1, fig2, fig3, fig4 = create_visualizations(df_clean, stat_armada, avg_time_per_jenis)

            st.plotly_chart(fig1, use_container_width=True)
            st.plotly_chart(fig2, use_container_width=True)
            st.plotly_chart(fig3, use_container_width=True)
            st.plotly_chart(fig4, use_container_width=True)

            # Laporan otomatis dengan DeepSeek
            st.subheader("📝 Laporan Singkat (dibuat oleh AI DeepSeek)")
            # Siapkan ringkasan statistik sebagai teks
            stats_summary = f"""
Total trip: {len(df_clean)}
Total tonase: {df_clean['Tonase'].sum():,.2f}
Rata-rata waktu tempuh: {df_clean['Waktu Tempuh (jam)'].mean():.2f} jam
Armada teraktif: {most_active['No. Pintu']} ({int(most_active['Jumlah Trip'])} trip)
Armada paling sedikit trip: {least_active['No. Pintu']} ({int(least_active['Jumlah Trip'])} trip)
Rata-rata waktu tempuh per jenis armada:
{avg_time_per_jenis.to_string(index=False)}
            """
            if st.button("🔮 Buat Laporan dengan DeepSeek"):
                with st.spinner("Menghubungi DeepSeek API..."):
                    report = generate_report(stats_summary)
                st.markdown("### Laporan Hasil Analisis")
                st.write(report)
            else:
                st.info("Klik tombol di atas untuk menghasilkan laporan otomatis menggunakan AI DeepSeek (API key diperlukan).")

            # Tombol unduh Master Data yang sudah bersih
            @st.cache_data
            def convert_df_to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Master Data', index=False)
                return output.getvalue()

            excel_data = convert_df_to_excel(df_clean)
            st.download_button(label="📥 Unduh Master Data (Excel)",
                               data=excel_data,
                               file_name="master_data_armada.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("👆 Silakan unggah file Excel Anda untuk memulai. Pastikan sheet pertama adalah **List Armada** dan sisanya adalah **data harian** (maksimal 29 sheet).")
