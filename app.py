import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import requests
import os
from datetime import datetime
from io import BytesIO

# ================== KONFIGURASI HALAMAN ==================
st.set_page_config(
    page_title="Dashboard Armada | Analitik & Rekomendasi",
    page_icon="🚛",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS untuk tampilan dashboard profesional
st.markdown("""
<style>
    .metric-card {
        background: linear-gradient(135deg, #1e3c72 0%, #2a5298 100%);
        border-radius: 16px;
        padding: 20px;
        color: white;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        text-align: center;
        transition: transform 0.2s;
    }
    .metric-card:hover {
        transform: translateY(-5px);
    }
    .metric-value {
        font-size: 2.8rem;
        font-weight: 800;
        margin: 10px 0;
    }
    .metric-label {
        font-size: 1rem;
        opacity: 0.85;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .stButton button {
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 12px 24px;
        border-radius: 8px;
        font-weight: 600;
        transition: all 0.3s ease;
    }
    .stButton button:hover {
        transform: scale(1.02);
        box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
    }
</style>
""", unsafe_allow_html=True)

# ================== API DEEPSEEK ==================
DEEPSEEK_API_KEY = os.getenv("DEEPSEEK_API_KEY", "sk-xxxxxxxxxxxxxxxxxxxxxxxx")
DEEPSEEK_URL = "https://api.deepseek.com/v1/chat/completions"

def generate_deepseek_report(stats_text):
    if not DEEPSEEK_API_KEY or DEEPSEEK_API_KEY.startswith("sk-xxx"):
        return "⚠️ API Key DeepSeek belum diatur. Laporan tidak dapat dibuat."
    headers = {"Authorization": f"Bearer {DEEPSEEK_API_KEY}", "Content-Type": "application/json"}
    prompt = f"""
Sebagai analis data armada truk, buat laporan singkat (3 paragraf) dalam bahasa Indonesia berdasarkan statistik:
{stats_text}
Sertakan: (1) gambaran umum, (2) armada teraktif dan paling tidak efisien beserta dugaan penyebab, 
(3) rekomendasi perbaikan (penjadwalan, rute, perawatan).
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
        return f"❌ Gagal menghasilkan laporan: {str(e)}"

# ================== FUNGSI DETEKSI KOLOM OTOMATIS ==================
def find_column(cols, keywords):
    for col in cols:
        col_lower = col.lower()
        for kw in keywords:
            if kw in col_lower:
                return col
    return None

# ================== MEMBACA FILE EXCEL ==================
@st.cache_data(show_spinner="Membaca file Excel...")
def read_excel_sheets(uploaded_file):
    try:
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
    except:
        xls = pd.ExcelFile(uploaded_file)
    sheets = {}
    for name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=name)
        if not df.empty:
            sheets[name] = df
    return sheets

# ================== PEMETAAN SHEET & KOLOM ==================
def mapping_ui(all_sheets):
    sheet_names = list(all_sheets.keys())
    st.sidebar.header("📂 Konfigurasi Data")
    armada_sheet = st.sidebar.selectbox("Sheet List Armada", sheet_names,
                                        help="Sheet berisi data referensi armada (No. Pintu, Plat).")
    daily_candidates = [s for s in sheet_names if s != armada_sheet]
    default_daily = [s for s in daily_candidates if "harian" in s.lower() or "daily" in s.lower()] or daily_candidates[:5]
    daily_sheets = st.sidebar.multiselect("Sheet Data Harian (bisa >30)", daily_candidates, default=default_daily)
    
    if not daily_sheets:
        st.error("Pilih minimal satu sheet harian!")
        return None
    
    # Sampel untuk deteksi kolom
    df_arm = all_sheets[armada_sheet]
    df_day_sample = all_sheets[daily_sheets[0]]
    
    with st.sidebar.expander("⚙️ Pemetaan Kolom Lanjutan", expanded=False):
        st.markdown("**List Armada**")
        cols_arm = df_arm.columns.tolist()
        no_pintu_arm = st.selectbox("No. Pintu", cols_arm, 
                                    index=cols_arm.index(find_column(cols_arm, ["no. pintu", "pintu", "nopintu"])) if find_column(cols_arm, ["no. pintu", "pintu", "nopintu"]) else 0)
        plat_arm = st.selectbox("Plat Mobil", cols_arm,
                                index=cols_arm.index(find_column(cols_arm, ["plat", "nopol", "nomor polisi"])) if find_column(cols_arm, ["plat", "nopol", "nomor polisi"]) else 0)
        jenis_arm = st.selectbox("Jenis Armada (opsional)", ["(tidak ada)"] + cols_arm,
                                 index=(["("+")"] + cols_arm).index(find_column(cols_arm, ["jenis", "tipe", "type"])) if find_column(cols_arm, ["jenis", "tipe", "type"]) else 0)
        
        st.markdown("**Data Harian**")
        cols_day = df_day_sample.columns.tolist()
        no_pintu_day = st.selectbox("No. Pintu (Harian)", cols_day,
                                    index=cols_day.index(find_column(cols_day, ["no. pintu", "pintu", "nopintu"])) if find_column(cols_day, ["no. pintu", "pintu", "nopintu"]) else 0)
        plat_day = st.selectbox("Plat Mobil (Harian)", cols_day,
                                index=cols_day.index(find_column(cols_day, ["plat", "nopol"])) if find_column(cols_day, ["plat", "nopol"]) else 0)
        tonase = st.selectbox("Tonase", cols_day,
                              index=cols_day.index(find_column(cols_day, ["tonase", "ton", "berat", "muatan"])) if find_column(cols_day, ["tonase", "ton", "berat", "muatan"]) else 0)
        tgl_berangkat = st.selectbox("Waktu Berangkat", cols_day,
                                     index=cols_day.index(find_column(cols_day, ["berangkat", "start", "berang"])) if find_column(cols_day, ["berangkat", "start", "berang"]) else 0)
        tgl_tiba = st.selectbox("Waktu Tiba", cols_day,
                                index=cols_day.index(find_column(cols_day, ["tiba", "finish", "end"])) if find_column(cols_day, ["tiba", "finish", "end"]) else 0)
    
    return {
        "armada_sheet": armada_sheet,
        "daily_sheets": daily_sheets,
        "cols": {
            "arm_no_pintu": no_pintu_arm,
            "arm_plat": plat_arm,
            "arm_jenis": jenis_arm if jenis_arm != "(tidak ada)" else None,
            "day_no_pintu": no_pintu_day,
            "day_plat": plat_day,
            "day_tonase": tonase,
            "day_berangkat": tgl_berangkat,
            "day_tiba": tgl_tiba
        }
    }

# ================== PROSES DATA ==================
def process_and_clean(all_sheets, mapping):
    df_armada = all_sheets[mapping["armada_sheet"]].copy()
    dfs_daily = [all_sheets[s] for s in mapping["daily_sheets"]]
    df_raw = pd.concat(dfs_daily, ignore_index=True)
    
    c = mapping["cols"]
    
    # Normalisasi
    def norm(x): return str(x).strip().upper() if pd.notna(x) else ""
    
    df_armada["_no_pintu"] = df_armada[c["arm_no_pintu"]].apply(norm)
    df_armada["_plat"] = df_armada[c["arm_plat"]].apply(norm)
    df_armada["_jenis"] = df_armada[c["arm_jenis"]].apply(lambda x: str(x).strip() if pd.notna(x) else "Tidak Diketahui") if c["arm_jenis"] else "Tidak Diketahui"
    
    df_raw["_no_pintu"] = df_raw[c["day_no_pintu"]].apply(norm)
    df_raw["_plat"] = df_raw[c["day_plat"]].apply(norm)
    df_raw["_tonase"] = pd.to_numeric(df_raw[c["day_tonase"]], errors='coerce')
    df_raw["_berangkat"] = pd.to_datetime(df_raw[c["day_berangkat"]], errors='coerce', dayfirst=True)
    df_raw["_tiba"] = pd.to_datetime(df_raw[c["day_tiba"]], errors='coerce', dayfirst=True)
    
    # Lookup armada
    pintu_to_plat = df_armada.set_index("_no_pintu")["_plat"].to_dict()
    plat_to_pintu = df_armada.set_index("_plat")["_no_pintu"].to_dict()
    pintu_to_jenis = df_armada.set_index("_no_pintu")["_jenis"].to_dict() if c["arm_jenis"] else {}
    
    mismatch = 0
    for idx, row in df_raw.iterrows():
        p = row["_no_pintu"]
        pl = row["_plat"]
        if p in pintu_to_plat and pl != pintu_to_plat[p]:
            df_raw.at[idx, "_plat"] = pintu_to_plat[p]
            mismatch += 1
        elif pl in plat_to_pintu and p != plat_to_pintu[pl]:
            df_raw.at[idx, "_no_pintu"] = plat_to_pintu[pl]
            mismatch += 1
    
    df_raw["_jenis"] = df_raw["_no_pintu"].map(pintu_to_jenis).fillna("Tidak Diketahui") if pintu_to_jenis else "Tidak Diketahui"
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
    
    return df_armada, df_clean, mismatch

# ================== APLIKASI UTAMA ==================
def main():
    st.title("🚛 Dashboard Analitik & Rekomendasi Armada")
    st.markdown("Upload file Excel (multi-sheet) untuk melihat performa armada, trip, tonase, efisiensi, dan dapatkan laporan AI.")
    
    uploaded_file = st.sidebar.file_uploader("Unggah file Excel", type=["xlsx", "xls"])
    
    if uploaded_file:
        all_sheets = read_excel_sheets(uploaded_file)
        if not all_sheets:
            st.error("File kosong atau tidak bisa dibaca.")
            return
        
        mapping = mapping_ui(all_sheets)
        if mapping is None:
            return
        
        if st.sidebar.button("🚀 Proses Data", use_container_width=True):
            with st.spinner("Memproses..."):
                df_armada, df_clean, mismatch = process_and_clean(all_sheets, mapping)
            
            st.success(f"✅ Data siap! {mismatch} ketidaksesuaian diperbaiki berdasarkan List Armada.")
            
            # Filter interaktif
            st.sidebar.header("🔍 Filter Data")
            jenis_list = ["Semua"] + sorted(df_clean["Jenis Armada"].unique().tolist())
            jenis_filter = st.sidebar.selectbox("Jenis Armada", jenis_list)
            if "Waktu Berangkat" in df_clean.columns:
                min_date = df_clean["Waktu Berangkat"].min()
                max_date = df_clean["Waktu Berangkat"].max()
                if pd.notna(min_date) and pd.notna(max_date):
                    date_range = st.sidebar.date_input("Rentang Tanggal", [min_date, max_date])
            
            # Aplikasi filter
            df = df_clean.copy()
            if jenis_filter != "Semua":
                df = df[df["Jenis Armada"] == jenis_filter]
            if "date_range" in locals() and len(date_range) == 2:
                df = df[(df["Waktu Berangkat"] >= pd.Timestamp(date_range[0])) & 
                        (df["Waktu Berangkat"] <= pd.Timestamp(date_range[1]))]
            
            # --- METRIK UTAMA ---
            total_trip = len(df)
            total_tonase = df["Tonase"].sum()
            avg_time = df["Waktu Tempuh (jam)"].mean()
            aktif = df["No. Pintu"].nunique()
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Trip</div><div class="metric-value">{total_trip}</div></div>', unsafe_allow_html=True)
            with col2:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Total Tonase</div><div class="metric-value">{total_tonase:,.1f}</div></div>', unsafe_allow_html=True)
            with col3:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Rata2 Waktu Tempuh</div><div class="metric-value">{avg_time:.1f} jam</div></div>', unsafe_allow_html=True)
            with col4:
                st.markdown(f'<div class="metric-card"><div class="metric-label">Armada Aktif</div><div class="metric-value">{aktif}</div></div>', unsafe_allow_html=True)
            
            st.markdown("---")
            
            # --- GRAFIK INTERAKTIF ---
            col_left, col_right = st.columns(2)
            
            with col_left:
                # Top armada by trip
                top_trip = df.groupby("No. Pintu").size().reset_index(name="Jumlah Trip").sort_values("Jumlah Trip", ascending=False).head(10)
                fig1 = px.bar(top_trip, x="No. Pintu", y="Jumlah Trip", color="Jumlah Trip",
                              color_continuous_scale="viridis", title="🏆 Top 10 Armada Berdasarkan Trip")
                fig1.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig1, use_container_width=True)
            
            with col_right:
                # Distribusi jenis armada
                jenis_counts = df["Jenis Armada"].value_counts().reset_index()
                jenis_counts.columns = ["Jenis", "Jumlah"]
                fig2 = px.pie(jenis_counts, names="Jenis", values="Jumlah",
                              title="🍩 Distribusi Trip per Jenis Armada", hole=0.4)
                st.plotly_chart(fig2, use_container_width=True)
            
            col_left2, col_right2 = st.columns(2)
            
            with col_left2:
                # Top tonase
                top_ton = df.groupby("No. Pintu")["Tonase"].sum().reset_index().sort_values("Tonase", ascending=False).head(10)
                fig3 = px.bar(top_ton, x="No. Pintu", y="Tonase", color="Tonase",
                              color_continuous_scale="orrd", title="📦 Top 10 Armada Berdasarkan Tonase")
                fig3.update_layout(xaxis_tickangle=-45)
                st.plotly_chart(fig3, use_container_width=True)
            
            with col_right2:
                # Rata-rata waktu tempuh per jenis
                avg_jenis = df.groupby("Jenis Armada")["Waktu Tempuh (jam)"].mean().reset_index()
                fig4 = px.bar(avg_jenis, x="Jenis Armada", y="Waktu Tempuh (jam)", color="Waktu Tempuh (jam)",
                              color_continuous_scale="blues", title="⏱️ Rata-rata Waktu Tempuh per Jenis Armada")
                st.plotly_chart(fig4, use_container_width=True)
            
            st.markdown("---")
            
            # --- ARMADA TERAKTIF & TIDAK EFISIEN ---
            stat_armada = df.groupby("No. Pintu").agg(
                Trip=("No. Pintu", "count"),
                Tonase=("Tonase", "sum")
            ).reset_index()
            most_active = stat_armada.loc[stat_armada["Trip"].idxmax()]
            least_active = stat_armada.loc[stat_armada["Trip"].idxmin()]
            
            col_a, col_b = st.columns(2)
            with col_a:
                st.info(f"### 🥇 Armada Teraktif\n**{most_active['No. Pintu']}**  \nTrip: {int(most_active['Trip'])} | Tonase: {most_active['Tonase']:,.1f}")
            with col_b:
                st.warning(f"### 🐌 Armada Paling Tidak Efisien\n**{least_active['No. Pintu']}**  \nTrip: {int(least_active['Trip'])} | Tonase: {least_active['Tonase']:,.1f}")
            
            # --- TABEL DATA ---
            with st.expander("📊 Lihat Data Armada & Statistik Lengkap"):
                st.dataframe(stat_armada.sort_values("Trip", ascending=False).style.background_gradient(cmap="Blues", subset=["Trip"]).format({"Tonase": "{:,.1f}"}))
            
            # --- LAPORAN DEEPSEEK ---
            st.subheader("📝 Laporan Cerdas dari DeepSeek AI")
            summary = f"""
Total trip: {total_trip}
Total tonase: {total_tonase:,.1f}
Rata-rata waktu tempuh: {avg_time:.1f} jam
Armada teraktif: {most_active['No. Pintu']} ({int(most_active['Trip'])} trip)
Armada paling sedikit trip: {least_active['No. Pintu']} ({int(least_active['Trip'])} trip)
Rata-rata waktu tempuh per jenis armada:
{avg_jenis.to_string(index=False)}
            """
            if st.button("🔮 Buat Laporan AI", key="report_btn"):
                with st.spinner("Menghubungi DeepSeek..."):
                    report = generate_deepseek_report(summary)
                st.markdown("### 📄 Laporan Hasil Analisis")
                st.write(report)
            else:
                st.info("Klik tombol di atas untuk menghasilkan laporan otomatis dengan AI (memerlukan API Key DeepSeek).")
            
            # --- DOWNLOAD MASTER DATA ---
            @st.cache_data
            def to_excel(df):
                output = BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, index=False, sheet_name='Master Data')
                return output.getvalue()
            
            st.download_button(
                label="📥 Unduh Master Data (Excel)",
                data=to_excel(df_clean),
                file_name="master_data_armada.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
    else:
        st.info("👆 Silakan unggah file Excel Anda untuk memulai analisis.")
        st.image("https://cdn-icons-png.flaticon.com/512/3081/3081559.png", width=150)

if __name__ == "__main__":
    main()
