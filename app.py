import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.io as pio
import requests
import os
from io import BytesIO

# -------------------------- KONFIGURASI --------------------------
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
Anda analis data armada. Buat laporan singkat (3 paragraf) dalam bahasa Indonesia dari statistik berikut:
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
def cari_kolom(kolom, kata_kunci):
    for col in kolom:
        if any(kw in col.lower() for kw in kata_kunci):
            return col
    return None

@st.cache_data(show_spinner="Membaca file...")
def baca_semua_sheet(uploaded_file):
    xls = pd.ExcelFile(uploaded_file, engine='openpyxl') if uploaded_file.name.endswith('.xlsx') else pd.ExcelFile(uploaded_file)
    return {name: pd.read_excel(xls, sheet_name=name) for name in xls.sheet_names if not pd.read_excel(xls, sheet_name=name).empty}

def deteksi_otomatis(sheets_dict):
    sheet_names = list(sheets_dict.keys())
    armada_sheet = next((s for s in sheet_names if "list armada" in s.lower() or "master" in s.lower() or "armada" in s.lower()), sheet_names[0])
    daily_sheets = [s for s in sheet_names if s != armada_sheet]
    if not daily_sheets:
        return None

    df_arm = sheets_dict[armada_sheet]
    df_day = sheets_dict[daily_sheets[0]]
    kolom_arm = df_arm.columns.tolist()
    kolom_day = df_day.columns.tolist()

    no_pintu_arm = cari_kolom(kolom_arm, ["no. pintu", "no pintu", "pintu", "nopintu"])
    plat_arm = cari_kolom(kolom_arm, ["plat", "nopol", "nomor polisi"])
    jenis_arm = cari_kolom(kolom_arm, ["jenis", "tipe", "type"])

    no_pintu_day = cari_kolom(kolom_day, ["no. pintu", "no pintu", "pintu", "nopintu"])
    plat_day = cari_kolom(kolom_day, ["plat", "nopol"])
    tonase = cari_kolom(kolom_day, ["tonase", "ton", "berat", "muatan", "tonnase"])
    tgl_berangkat = cari_kolom(kolom_day, ["berangkat", "start", "berang"])
    tgl_tiba = cari_kolom(kolom_day, ["tiba", "finish", "end"])

    if not all([no_pintu_arm, plat_arm, no_pintu_day, plat_day, tonase, tgl_berangkat, tgl_tiba]):
        return None

    return {
        "armada_sheet": armada_sheet,
        "daily_sheets": daily_sheets,
        "kolom": {
            "arm_no_pintu": no_pintu_arm,
            "arm_plat": plat_arm,
            "arm_jenis": jenis_arm,
            "day_no_pintu": no_pintu_day,
            "day_plat": plat_day,
            "day_tonase": tonase,
            "day_berangkat": tgl_berangkat,
            "day_tiba": tgl_tiba
        }
    }

def proses_data(sheets_dict, mapping):
    df_armada = sheets_dict[mapping["armada_sheet"]].copy()
    dfs_harian = [sheets_dict[s] for s in mapping["daily_sheets"]]
    df_raw = pd.concat(dfs_harian, ignore_index=True)
    c = mapping["kolom"]

    def norm(x): return str(x).strip().upper() if pd.notna(x) else ""

    df_armada["_no_pintu"] = df_armada[c["arm_no_pintu"]].apply(norm)
    df_armada["_plat"] = df_armada[c["arm_plat"]].apply(norm)
    df_armada["_jenis"] = df_armada[c["arm_jenis"]].apply(lambda x: str(x).strip()) if c["arm_jenis"] else "Tidak Diketahui"

    df_raw["_no_pintu"] = df_raw[c["day_no_pintu"]].apply(norm)
    df_raw["_plat"] = df_raw[c["day_plat"]].apply(norm)
    df_raw["_tonase"] = pd.to_numeric(df_raw[c["day_tonase"]], errors='coerce')
    df_raw["_berangkat"] = pd.to_datetime(df_raw[c["day_berangkat"]], errors='coerce', dayfirst=True)
    df_raw["_tiba"] = pd.to_datetime(df_raw[c["day_tiba"]], errors='coerce', dayfirst=True)

    pintu_ke_plat = df_armada.set_index("_no_pintu")["_plat"].to_dict()
    plat_ke_pintu = df_armada.set_index("_plat")["_no_pintu"].to_dict()
    pintu_ke_jenis = df_armada.set_index("_no_pintu")["_jenis"].to_dict() if c["arm_jenis"] else {}

    perbaikan = 0
    for idx, row in df_raw.iterrows():
        p = row["_no_pintu"]
        pl = row["_plat"]
        if p in pintu_ke_plat and pl != pintu_ke_plat[p]:
            df_raw.at[idx, "_plat"] = pintu_ke_plat[p]
            perbaikan += 1
        elif pl in plat_ke_pintu and p != plat_ke_pintu[pl]:
            df_raw.at[idx, "_no_pintu"] = plat_ke_pintu[pl]
            perbaikan += 1

    df_raw["_jenis"] = df_raw["_no_pintu"].map(pintu_ke_jenis).fillna("Tidak Diketahui")
    df_raw["_waktu_tempuh"] = (df_raw["_tiba"] - df_raw["_berangkat"]).dt.total_seconds() / 3600
    df_raw.loc[df_raw["_waktu_tempuh"] < 0, "_waktu_tempuh"] = np.nan

    df_clean = df_raw.rename(columns={
        "_no_pintu": "No. Pintu",
        "_plat": "Plat Mobil",
        "_tonase": "Tonase",
        "_berangkat": "Waktu Berangkat",
        "_tiba": "Waktu Tiba",
        "_jenis": "Jenis Armada",
        "_waktu_tempuh": "Waktu Tempuh (jam)"
    })
    return df_armada, df_clean, perbaikan

# -------------------------- INISIALISASI SESSION STATE --------------------------
if "data_processed" not in st.session_state:
    st.session_state.data_processed = False
    st.session_state.df_clean = None
    st.session_state.stat_armada = None
    st.session_state.fig1 = None
    st.session_state.fig2 = None
    st.session_state.fig3 = None
    st.session_state.fig4 = None
    st.session_state.laporan_teks = ""

# -------------------------- APLIKASI UTAMA --------------------------
def main():
    st.title("🚛 Dashboard Analitik & Rekomendasi Armada")
    st.markdown("Unggah file Excel dengan banyak sheet (1 List Armada + 29 Harian) atau format lainnya. Sistem akan otomatis mendeteksi dan memproses data, lalu Anda bisa **mengunduh semua hasil** di akhir.")

    with st.sidebar:
        uploaded_file = st.file_uploader("📂 Pilih file Excel", type=["xlsx", "xls"])
        if uploaded_file:
            st.success("File berhasil diunggah")

    if uploaded_file:
        sheets_dict = baca_semua_sheet(uploaded_file)
        if not sheets_dict:
            st.error("File tidak memiliki sheet yang valid.")
            return

        mapping = deteksi_otomatis(sheets_dict)

        if mapping is None:
            st.warning("Deteksi otomatis gagal. Silakan pilih sheet dan kolom secara manual.")
            with st.sidebar.expander("⚙️ Pemetaan Manual", expanded=True):
                sheet_names = list(sheets_dict.keys())
                armada_sheet = st.selectbox("Sheet List Armada", sheet_names)
                daily_sheets = st.multiselect("Sheet Harian", [s for s in sheet_names if s != armada_sheet])
                if daily_sheets:
                    df_arm = sheets_dict[armada_sheet]
                    df_day = sheets_dict[daily_sheets[0]]
                    cols_arm = df_arm.columns.tolist()
                    cols_day = df_day.columns.tolist()
                    c_no_pintu_arm = st.selectbox("No. Pintu (Armada)", cols_arm, index=0)
                    c_plat_arm = st.selectbox("Plat (Armada)", cols_arm, index=min(1, len(cols_arm)-1))
                    c_jenis_arm = st.selectbox("Jenis Armada (opsional)", ["(tidak ada)"] + cols_arm, index=0)
                    c_no_pintu_day = st.selectbox("No. Pintu (Harian)", cols_day, index=0)
                    c_plat_day = st.selectbox("Plat (Harian)", cols_day, index=min(1, len(cols_day)-1))
                    c_tonase = st.selectbox("Tonase", cols_day, index=min(2, len(cols_day)-1))
                    c_berangkat = st.selectbox("Waktu Berangkat", cols_day, index=min(3, len(cols_day)-1))
                    c_tiba = st.selectbox("Waktu Tiba", cols_day, index=min(4, len(cols_day)-1))
                    mapping = {
                        "armada_sheet": armada_sheet,
                        "daily_sheets": daily_sheets,
                        "kolom": {
                            "arm_no_pintu": c_no_pintu_arm,
                            "arm_plat": c_plat_arm,
                            "arm_jenis": c_jenis_arm if c_jenis_arm != "(tidak ada)" else None,
                            "day_no_pintu": c_no_pintu_day,
                            "day_plat": c_plat_day,
                            "day_tonase": c_tonase,
                            "day_berangkat": c_berangkat,
                            "day_tiba": c_tiba
                        }
                    }
        else:
            st.sidebar.success("✅ Deteksi otomatis berhasil!")
            st.sidebar.info(f"Sheet Armada: {mapping['armada_sheet']}\nJumlah sheet harian: {len(mapping['daily_sheets'])}")

        if st.sidebar.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Menggabungkan dan membersihkan data..."):
                df_armada, df_clean, perbaikan = proses_data(sheets_dict, mapping)

            st.session_state.data_processed = True
            st.session_state.df_clean = df_clean
            st.success(f"✅ Data berhasil diproses. {perbaikan} ketidaksesuaian diperbaiki berdasarkan List Armada.")
            st.balloons()

        # Jika data sudah diproses, tampilkan dashboard
        if st.session_state.data_processed:
            df_clean = st.session_state.df_clean

            # Filter
            st.sidebar.header("🔍 Filter")
            jenis_list = ["Semua"] + sorted(df_clean["Jenis Armada"].unique())
            jenis_terpilih = st.sidebar.selectbox("Jenis Armada", jenis_list)
            if "Waktu Berangkat" in df_clean.columns:
                min_tgl = df_clean["Waktu Berangkat"].min()
                max_tgl = df_clean["Waktu Berangkat"].max()
                if pd.notna(min_tgl) and pd.notna(max_tgl):
                    rentang = st.sidebar.date_input("Rentang Tanggal", [min_tgl, max_tgl])

            df_filtered = df_clean.copy()
            if jenis_terpilih != "Semua":
                df_filtered = df_filtered[df_filtered["Jenis Armada"] == jenis_terpilih]
            if 'rentang' in locals() and len(rentang) == 2:
                df_filtered = df_filtered[(df_filtered["Waktu Berangkat"] >= pd.Timestamp(rentang[0])) &
                                          (df_filtered["Waktu Berangkat"] <= pd.Timestamp(rentang[1]))]

            # Metrik
            total_trip = len(df_filtered)
            total_tonase = df_filtered["Tonase"].sum()
            rata_waktu = df_filtered["Waktu Tempuh (jam)"].mean()
            armada_aktif = df_filtered["No. Pintu"].nunique()

            stat_armada = df_filtered.groupby("No. Pintu").agg(Trip=("No. Pintu", "count"), Tonase=("Tonase", "sum")).reset_index()
            paling_aktif = stat_armada.loc[stat_armada["Trip"].idxmax()] if not stat_armada.empty else None
            paling_sedikit = stat_armada.loc[stat_armada["Trip"].idxmin()] if not stat_armada.empty else None
            avg_jenis = df_filtered.groupby("Jenis Armada")["Waktu Tempuh (jam)"].mean().reset_index()

            # Grafik
            top_trip = df_filtered.groupby("No. Pintu").size().reset_index(name="Trip").nlargest(10, "Trip")
            fig1 = px.bar(top_trip, x="No. Pintu", y="Trip", color="Trip", color_continuous_scale="viridis",
                          title="🏆 10 Armada dengan Trip Terbanyak")
            fig1.update_layout(xaxis_tickangle=-45)

            distribusi_jenis = df_filtered["Jenis Armada"].value_counts().reset_index()
            distribusi_jenis.columns = ["Jenis", "Jumlah"]
            fig2 = px.pie(distribusi_jenis, names="Jenis", values="Jumlah", hole=0.4,
                          title="🍩 Distribusi Trip per Jenis Armada")

            top_ton = df_filtered.groupby("No. Pintu")["Tonase"].sum().reset_index().nlargest(10, "Tonase")
            fig3 = px.bar(top_ton, x="No. Pintu", y="Tonase", color="Tonase", color_continuous_scale="orrd",
                          title="📦 10 Armada dengan Tonase Terbesar")
            fig3.update_layout(xaxis_tickangle=-45)

            fig4 = px.bar(avg_jenis, x="Jenis Armada", y="Waktu Tempuh (jam)", color="Waktu Tempuh (jam)",
                          color_continuous_scale="blues", title="⏱️ Rata-rata Waktu Tempuh per Jenis")

            # Simpan di session state
            st.session_state.stat_armada = stat_armada
            st.session_state.fig1 = fig1
            st.session_state.fig2 = fig2
            st.session_state.fig3 = fig3
            st.session_state.fig4 = fig4

            # Tampilkan metrik
            col1, col2, col3, col4 = st.columns(4)
            col1.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
            col2.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
            col3.markdown(f'<div class="metric-card"><div class="metric-label">Rata² Waktu Tempuh</div><div class="metric-value">{rata_waktu:.1f} jam</div></div>', unsafe_allow_html=True)
            col4.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{armada_aktif}</div></div>', unsafe_allow_html=True)

            st.markdown("---")
            col_kiri, col_kanan = st.columns(2)
            with col_kiri:
                st.plotly_chart(fig1, use_container_width=True)
            with col_kanan:
                st.plotly_chart(fig2, use_container_width=True)

            col_kiri2, col_kanan2 = st.columns(2)
            with col_kiri2:
                st.plotly_chart(fig3, use_container_width=True)
            with col_kanan2:
                st.plotly_chart(fig4, use_container_width=True)

            st.markdown("---")
            if paling_aktif is not None:
                col_a, col_b = st.columns(2)
                with col_a:
                    st.info(f"### 🥇 Armada Teraktif\n**{paling_aktif['No. Pintu']}**  \nTrip: {int(paling_aktif['Trip'])} | Tonase: {paling_aktif['Tonase']:,.1f}")
                with col_b:
                    st.warning(f"### 🐌 Armada Paling Tidak Efisien\n**{paling_sedikit['No. Pintu']}**  \nTrip: {int(paling_sedikit['Trip'])} | Tonase: {paling_sedikit['Tonase']:,.1f}")

            with st.expander("📊 Lihat Data Lengkap & Statistik"):
                st.dataframe(stat_armada.sort_values("Trip", ascending=False).style.format({"Tonase": "{:,.1f}"}))

            # Laporan AI
            st.subheader("📝 Laporan Cerdas dari DeepSeek AI")
            statistik_teks = f"""
Total trip: {total_trip}
Total tonase: {total_tonase:,.1f}
Rata-rata waktu tempuh: {rata_waktu:.1f} jam
Armada teraktif: {paling_aktif['No. Pintu']} ({int(paling_aktif['Trip'])} trip)
Armada paling sedikit trip: {paling_sedikit['No. Pintu']} ({int(paling_sedikit['Trip'])} trip)
Rata-rata waktu tempuh per jenis armada:
{avg_jenis.to_string(index=False)}
            """
            if st.button("🔮 Buat Laporan AI", key="laporan_btn"):
                with st.spinner("Menghubungi DeepSeek..."):
                    st.session_state.laporan_teks = laporan_ai(statistik_teks)
                st.markdown("### 📄 Hasil Laporan")
                st.write(st.session_state.laporan_teks)
            else:
                st.info("Klik tombol di atas untuk menghasilkan laporan otomatis (API Key diperlukan).")

            # Bagian Download
            st.markdown("## 📥 Unduh Hasil Analisis")
            st.markdown("Pilih file yang ingin Anda unduh:")

            @st.cache_data
            def ke_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Master Data')
                return output.getvalue()

            @st.cache_data
            def ke_csv(df):
                return df.to_csv(index=False).encode('utf-8')

            def grafik_ke_png(fig):
                try:
                    return pio.to_image(fig, format='png', scale=2)
                except Exception as e:
                    st.error(f"Gagal mengonversi grafik: {e}. Pastikan kaleido terinstal (pip install kaleido).")
                    return None

            col_d1, col_d2, col_d3, col_d4 = st.columns(4)
            with col_d1:
                st.download_button(
                    label="📊 Master Data (Excel)",
                    data=ke_excel(df_clean),
                    file_name="master_data_armada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            with col_d2:
                st.download_button(
                    label="📈 Statistik Armada (CSV)",
                    data=ke_csv(stat_armada),
                    file_name="statistik_armada.csv",
                    mime="text/csv",
                    use_container_width=True
                )
            with col_d3:
                pilihan_grafik = st.selectbox("Pilih grafik", ["Trip Terbanyak", "Distribusi Jenis", "Tonase Terbesar", "Waktu Tempuh per Jenis"])
                fig_download = None
                if pilihan_grafik == "Trip Terbanyak":
                    fig_download = st.session_state.fig1
                elif pilihan_grafik == "Distribusi Jenis":
                    fig_download = st.session_state.fig2
                elif pilihan_grafik == "Tonase Terbesar":
                    fig_download = st.session_state.fig3
                else:
                    fig_download = st.session_state.fig4

                if fig_download is not None:
                    png_data = grafik_ke_png(fig_download)
                    if png_data:
                        st.download_button(
                            label="📸 Unduh Grafik (PNG)",
                            data=png_data,
                            file_name=f"grafik_{pilihan_grafik.lower().replace(' ', '_')}.png",
                            mime="image/png",
                            use_container_width=True
                        )
                    else:
                        st.warning("Grafik tidak dapat diunduh (kaleido tidak terinstal).")
                else:
                    st.warning("Proses data terlebih dahulu.")

            with col_d4:
                if st.session_state.laporan_teks:
                    st.download_button(
                        label="📝 Laporan AI (TXT)",
                        data=st.session_state.laporan_teks.encode('utf-8'),
                        file_name="laporan_ai.txt",
                        mime="text/plain",
                        use_container_width=True
                    )
                else:
                    st.markdown("*(Buat laporan AI dulu untuk mengunduh)*")

    else:
        st.info("👆 Silakan unggah file Excel Anda untuk memulai.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
