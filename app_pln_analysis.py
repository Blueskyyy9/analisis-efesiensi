import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import base64
import openpyxl

# -------------------------------
# KONFIGURASI AWAL
# -------------------------------
st.set_page_config(
    page_title="Simulasi RAB & Efisiensi PLN",
    layout="wide",
    page_icon="⚡"
)

# Konstanta
REQUIRED_SHEETS = ["RAB", "Gambar"]
DEFAULT_KABEL_DB = {
    "NFA2X-T 2 x 70 + N 70 mm²": 0.443,
    "NFA2X-T 3 x 70 + N 70 mm²": 0.443
}
COLORS = {
    "good": "#65e23b",
    "moderate": "#e6bf42",
    "poor": "#e62535",
    "very_good": "#68e240",
    "fair": "#e0b62a",
    "low": "#e01425",
    "excellent": "#77f04f",
    "adequate": "#e4b725",
    "unfeasible": "#f12738"
}

# -------------------------------
# FUNGSI UI DAN STYLING
# -------------------------------
def set_background(image_file):
    """Menambahkan latar belakang untuk aplikasi."""
    with open(image_file, "rb") as file:
        encoded_string = base64.b64encode(file.read()).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url(data:image/jpeg;base64,{encoded_string});
            background-size: cover;
            background-position: center;
            background-attachment: fixed;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

def set_sidebar_background(image_file):
    """Menambahkan latar belakang untuk sidebar."""
    with open(image_file, "rb") as file:
        encoded_string = base64.b64encode(file.read()).decode()
    st.markdown(
        f"""
        <style>
        [data-testid="stSidebar"] {{
            background-image: url(data:image/png;base64,{encoded_string});
            background-size: cover;
            background-position: center;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

def display_header():
    """Menampilkan header aplikasi dengan informasi utama."""
    st.markdown(
        """
        <h1 style='text-align:center;color:#FFFFFF;font-family:"Trebuchet MS", sans-serif;font-size:35px;font-weight:800;text-shadow:5px 5px 9px #000000;'>
        ⚡ Simulasi Efisiensi & Analisis RAB PLN ULP Batang ⚡
        </h1>
        <p style='text-align:center;'>Program ini membaca file <b>RAB PLN (2 Sheet: RAB & Gambar)</b> untuk menghitung:</p>
        <ul>
            <li>Total biaya per komponen</li>
            <li>Rugi daya (losses I²R) berdasarkan jenis kabel dan panjang</li>
            <li>Efisiensi teknis & trafo</li>
            <li>ROI proyek berdasarkan penghematan rugi daya</li>
        </ul>
        """,
        unsafe_allow_html=True
    )

def configure_sidebar():
    """Mengatur sidebar untuk input asumsi dan pemilihan kabel."""
    st.sidebar.header("⚙️ Pengaturan")
    tarif_kwh = st.sidebar.number_input("Tarif Listrik (Rp/kWh)", value=1500, min_value=0)
    faktor_daya = st.sidebar.number_input("Faktor Daya", value=0.8, min_value=0.0, max_value=1.0)
    tipe_phase = st.sidebar.selectbox("Tipe Phase", ["3 Phase", "1 Phase"], index=0)
    rugi_trafo = st.sidebar.number_input("Rugi Daya Trafo Default (kW)", value=0.5, min_value=0.0)
    baseline_losses = st.sidebar.number_input("Baseline Losses per Lokasi (kW)", value=5.0, min_value=0.0,
                                             help="Rugi daya sistem lama untuk menghitung penghematan")
    core_loss = st.sidebar.number_input("Core Losses Default (kW)", value=0.2, min_value=0.0,
                                        help="Rugi inti trafo jika tidak ada data")
    full_load_loss = st.sidebar.number_input("Full-Load Losses Default (kW)", value=1.0, min_value=0.0,
                                             help="Rugi tembaga trafo pada beban penuh")
    
    st.sidebar.subheader("🔌 Jenis Kabel")
    selected_kabel = st.sidebar.multiselect(
        "Pilih Jenis Kabel",
        list(DEFAULT_KABEL_DB.keys()),
        default=list(DEFAULT_KABEL_DB.keys()),
        help="Pilih kabel untuk perhitungan. Resistansi tetap 0,443 ohm/km."
    )
    resistansi_kabel = {k: DEFAULT_KABEL_DB[k] for k in selected_kabel}
    
    return {
        "tarif_kwh": tarif_kwh,
        "faktor_daya": faktor_daya,
        "tipe_phase": tipe_phase,
        "rugi_trafo": rugi_trafo,
        "baseline_losses": baseline_losses,
        "core_loss": core_loss,
        "full_load_loss": full_load_loss,
        "resistansi_kabel": resistansi_kabel,
        "selected_kabel": selected_kabel
    }

# -------------------------------
# FUNGSI VALIDASI DATA
# -------------------------------
def validate_excel_file(uploaded_file):
    """Memvalidasi file Excel dan mengembalikan DataFrame untuk RAB dan Gambar."""
    try:
        excel_file = pd.ExcelFile(uploaded_file)
        if not all(sheet in excel_file.sheet_names for sheet in REQUIRED_SHEETS):
            st.error(f"File Excel harus memiliki sheet: {REQUIRED_SHEETS}.")
            st.stop()
        return pd.read_excel(uploaded_file, sheet_name="RAB"), pd.read_excel(uploaded_file, sheet_name="Gambar")
    except Exception as e:
        st.error(f"Terjadi kesalahan saat membaca file: {e}")
        st.stop()

def validate_columns(df_rab, df_gambar):
    """Memvalidasi kolom wajib di sheet RAB dan Gambar."""
    required_columns_rab = ["Total (Rp)"]
    required_columns_gambar = ["Nama Lokasi", "Jenis Kabel", "Panjang Jaringan (m)", "Beban Total (kVA)", "Tegangan (V)"]
    
    if not set(required_columns_rab).issubset(df_rab.columns):
        st.error("Kolom wajib di sheet 'RAB' tidak lengkap.")
        st.stop()
    if not set(required_columns_gambar).issubset(df_gambar.columns):
        st.error("Kolom wajib di sheet 'Gambar' tidak lengkap.")
        st.stop()

def validate_numeric(df, columns):
    """Memvalidasi bahwa kolom berisi data numerik yang valid."""
    for col in columns:
        if not pd.api.types.is_numeric_dtype(df[col]):
            st.error(f"Kolom '{col}' harus berisi angka.")
            st.stop()
        if (df[col] < 0).any():
            st.error(f"Kolom '{col}' tidak boleh negatif.")
            st.stop()
        if col in ["Beban Total (kVA)", "Tegangan (V)"] and (df[col] <= 0).any():
            st.error(f"Kolom '{col}' tidak boleh nol atau negatif.")
            st.stop()

def validate_kabel(df_gambar, selected_kabel):
    """Memvalidasi jenis kabel di sheet Gambar sesuai dengan pilihan sidebar."""
    valid_kabel = selected_kabel + ["-"]
    kabel_excel = df_gambar["Jenis Kabel"].unique()
    invalid_kabel = [k for k in kabel_excel if k not in valid_kabel]
    
    if invalid_kabel:
        st.error(f"Jenis kabel tidak diizinkan: {invalid_kabel}. Hanya {valid_kabel} yang diperbolehkan.")
        st.stop()

def preprocess_data(df_gambar, tipe_phase, rugi_trafo, baseline_losses, core_loss, full_load_loss):
    """Memproses data Gambar: validasi tipe phase, rugi trafo, dan baseline losses."""
    df_gambar = df_gambar.copy()
    
    # Normalisasi Tipe Phase
    if "Tipe Phase" in df_gambar.columns:
        df_gambar["Tipe Phase"] = df_gambar["Tipe Phase"].str.upper().replace({"1 PHASE": "1 Phase", "3 PHASE": "3 Phase"})
    else:
        st.warning(f"Kolom 'Tipe Phase' tidak ada, diasumsikan '{tipe_phase}'.")
        df_gambar["Tipe Phase"] = tipe_phase
    
    # Validasi dan hitung rugi trafo
    if "Daya Trafo (kVA)" in df_gambar.columns:
        validate_numeric(df_gambar, ["Daya Trafo (kVA)"])
        if (df_gambar["Beban Total (kVA)"] > 0.9 * df_gambar["Daya Trafo (kVA)"]).any():
            st.warning("Beban Total (kVA) melebihi 90% Daya Trafo (kVA) di beberapa lokasi.")
        df_gambar["Rugi Trafo (kW)"] = df_gambar.apply(
            lambda row: calculate_transformer_loss(row["Beban Total (kVA)"], row["Daya Trafo (kVA)"], core_loss, full_load_loss),
            axis=1
        )
    else:
        if "Rugi Trafo (kW)" in df_gambar.columns:
            validate_numeric(df_gambar, ["Rugi Trafo (kW)"])
        else:
            df_gambar["Rugi Trafo (kW)"] = rugi_trafo
    
    # Validasi Baseline Losses
    if "Baseline Losses (kW)" not in df_gambar.columns:
        df_gambar["Baseline Losses (kW)"] = baseline_losses
    
    # Bersihkan data
    df_gambar.fillna({
        "Jenis Kabel": "-",
        "Panjang Jaringan (m)": 0,
        "Beban Total (kVAA)": 0,
        "Tegangan (V)": 380
    }, inplace=True)
    
    return df_gambar

# -------------------------------
# FUNGSI PERHITUNGAN
# -------------------------------
@st.cache_data
def calculate_transformer_loss(beban, daya_trafo, core_loss, full_load_loss):
    """Menghitung rugi trafo berdasarkan beban relatif."""
    if beban <= 0 or daya_trafo <= 0:
        return core_loss
    load_ratio = beban / daya_trafo
    copper_loss = (load_ratio ** 2) * full_load_loss
    return core_loss + copper_loss

@st.cache_data
def calculate_conductor_loss(row, resistansi_kabel):
    """Menghitung rugi daya konduktor (I²R) dalam kW."""
    jenis = row.get("Jenis Kabel", "-")
    panjang = row.get("Panjang Jaringan (m)", 0)
    beban = row.get("Beban Total (kVA)", 0)
    tegangan = row.get("Tegangan (V)", 380)
    tipe_phase = row.get("Tipe Phase", "3 Phase").upper()
    
    if beban <= 0 or tegangan <= 0:
        st.warning(f"Input tidak valid di lokasi '{row.get('Nama Lokasi', '-')}' (beban/tegangan <=0), losses dianggap 0.")
        return 0
    
    if jenis not in resistansi_kabel:
        default_kabel = list(resistansi_kabel.keys())[0]
        st.warning(f"Jenis kabel '{jenis}' tidak dikenali di lokasi '{row.get('Nama Lokasi', '-')}', menggunakan '{default_kabel}'.")
        jenis = default_kabel
    
    r = resistansi_kabel[jenis] * (panjang / 1000)
    i = (beban * 1000) / (np.sqrt(3) * tegangan if tipe_phase == "3 PHASE" else tegangan)
    return (i ** 2) * r / 1000  # kW

@st.cache_data
def calculate_efficiency(conductor_loss, transformer_loss, beban, faktor_daya):
    """Menghitung efisiensi sistem dalam persen."""
    if beban <= 0 or faktor_daya <= 0:
        return 0
    total_loss = conductor_loss + transformer_loss
    daya_masuk = beban * faktor_daya
    return 100 * (1 - total_loss / daya_masuk)

@st.cache_data
def recommend_cable(row, resistansi_kabel):
    """Merekomendasikan kabel dengan rugi daya terendah."""
    beban = row.get("Beban Total (kVA)", 0)
    if beban <= 0:
        return "-"
    losses = {}
    for jenis, resistansi in resistansi_kabel.items():
        r = resistansi * (row["Panjang Jaringan (m)"] / 1000)
        i = (beban * 1000) / (np.sqrt(3) * row["Tegangan (V)"] if row["Tipe Phase"] == "3 Phase" else row["Tegangan (V)"])
        losses[jenis] = (i ** 2) * r / 1000
    return min(losses, key=losses.get)

# -------------------------------
# FUNGSI VISUALISASI DAN OUTPUT
# -------------------------------
def display_results(df_rab, df_gambar, tarif_kwh, faktor_daya, resistansi_kabel):
    """Menampilkan hasil analisis, tabel, grafik, dan kesimpulan."""
    # Perhitungan
    df_gambar["Losses Konduktor (kW)"] = df_gambar.apply(lambda row: calculate_conductor_loss(row, resistansi_kabel), axis=1)
    df_gambar["Losses Total (kW)"] = df_gambar["Losses Konduktor (kW)"] + df_gambar["Rugi Trafo (kW)"]
    df_gambar["Efisiensi (%)"] = df_gambar.apply(
        lambda row: calculate_efficiency(row["Losses Konduktor (kW)"], row["Rugi Trafo (kW)"], row["Beban Total (kVA)"], faktor_daya),
        axis=1
    )
    df_gambar["Penghematan Losses (kW)"] = df_gambar["Baseline Losses (kW)"] - df_gambar["Losses Total (kW)"]
    df_gambar["Manfaat (Rp/tahun)"] = df_gambar["Penghematan Losses (kW)"].clip(lower=0) * 8760 * tarif_kwh
    df_gambar["Rekomendasi Kabel"] = df_gambar.apply(lambda row: recommend_cable(row, resistansi_kabel), axis=1)
    
    # Validasi efisiensi
    if (df_gambar["Efisiensi (%)"] < 0).any() or (df_gambar["Efisiensi (%)"] > 100).any():
        st.warning("Efisiensi di luar rentang realistis (0-100%). Periksa data input.")
    
    # Hitung total biaya dan ROI
    total_biaya = df_rab["Total (Rp)"].sum()
    total_manfaat = df_gambar["Manfaat (Rp/tahun)"].sum()
    roi = (total_manfaat / total_biaya) * 100 if total_biaya > 0 else 0
    
    # Tampilkan tabel
    st.subheader("Data RAB")
    st.dataframe(df_rab, use_container_width=True)
    st.subheader("Data Teknis Jaringan")
    st.dataframe(df_gambar, use_container_width=True)
    
    # Filter berdasarkan lokasi
    st.subheader("Hasil Analisis Teknis & Ekonomi")
    lokasi_options = ['Semua'] + list(df_gambar['Nama Lokasi'].unique())
    selected_lokasi = st.selectbox("Pilih Lokasi untuk Analisis", lokasi_options)
    df_filtered = df_gambar if selected_lokasi == 'Semua' else df_gambar[df_gambar['Nama Lokasi'] == selected_lokasi]
    
    st.dataframe(
        df_filtered[["Nama Lokasi", "Jenis Kabel", "Panjang Jaringan (m)", "Beban Total (kVA)", "Tegangan (V)",
                     "Daya Trafo (kVA)", "Losses Konduktor (kW)", "Losses Total (kW)", "Penghematan Losses (kW)",
                     "Efisiensi (%)", "Manfaat (Rp/tahun)", "Rekomendasi Kabel"]]
        .style.format({
            "Losses Konduktor (kW)": "{:.2f}",
            "Losses Total (kW)": "{:.2f}",
            "Penghematan Losses (kW)": "{:.2f}",
            "Efisiensi (%)": "{:.2f}",
            "Manfaat (Rp/tahun)": "Rp {:,.0f}"
        }),
        use_container_width=True
    )
    
    # Visualisasi Losses
    st.subheader("Visualisasi Losses per Lokasi")
    fig = px.bar(
        df_filtered,
        x="Nama Lokasi",
        y="Losses Total (kW)",
        color="Losses Total (kW)",
        color_continuous_scale="Plasma",
        title="Rugi Daya Total per Lokasi",
        height=500,
        text="Losses Total (kW)"
    )
    fig.update_traces(
        texttemplate="%{text:.2f} kW",
        textposition="outside",
        marker=dict(line=dict(color="#000000", width=1))
    )
    fig.update_layout(
        title=dict(x=0.5, font=dict(size=20, family="Arial", color="#F1E9E9")),
        xaxis_title="Lokasi",
        yaxis_title="Rugi Daya Total (kW)",
        font=dict(family="Arial", size=12, color="#EEE6E6"),
        plot_bgcolor="rgba(0,0,0,1)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, tickangle=45),
        yaxis=dict(showgrid=True, gridcolor="rgba(200,200,200,0.3)"),
        margin=dict(l=50, r=50, t=80, b=100),
        showlegend=False
    )
    st.plotly_chart(fig, use_container_width=True)
    
    # Visualisasi Efisiensi
    st.subheader("Visualisasi Efisiensi per Lokasi")
    fig_eff = px.bar(
        df_filtered,
        x="Nama Lokasi",
        y="Efisiensi (%)",
        color="Efisiensi (%)",
        color_continuous_scale="Viridis",
        title="Efisiensi Sistem per Lokasi",
        height=500,
        text="Efisiensi (%)"
    )
    fig_eff.update_traces(
        texttemplate="%{text:.2f}%",
        textposition="outside",
        marker=dict(line=dict(color="#FFFFFF", width=1))
    )
    fig_eff.update_layout(
        title=dict(x=0.5, font=dict(size=20, family="Arial", color="#F0E7E7")),
        xaxis_title="Lokasi",
        yaxis_title="Efisiensi (%)",
        font=dict(family="Arial", size=12, color="#F3EBEB"),
        plot_bgcolor="rgba(0,0,0,1)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, tickangle=45),
        yaxis=dict(showgrid=True, gridcolor="rgba(200,200,200,0.3)"),
        margin=dict(l=50, r=50, t=80, b=100),
        showlegend=False
    )
    st.plotly_chart(fig_eff, use_container_width=True)
    
    # Visualisasi Penghematan
    st.subheader("Visualisasi Penghematan Losses per Lokasi")
    fig_penghematan = px.bar(
        df_filtered,
        x="Nama Lokasi",
        y="Penghematan Losses (kW)",
        color="Penghematan Losses (kW)",
        color_continuous_scale="Inferno",
        title="Penghematan Rugi Daya per Lokasi",
        height=500,
        text="Penghematan Losses (kW)"
    )
    fig_penghematan.update_traces(
        texttemplate="%{text:.2f} kW",
        textposition="outside",
        marker=dict(line=dict(color="#FFFFFF", width=1))
    )
    fig_penghematan.update_layout(
        title=dict(x=0.5, font=dict(size=20, family="Arial", color="#F1E7E7")),
        xaxis_title="Lokasi",
        yaxis_title="Penghematan Losses (kW)",
        font=dict(family="Arial", size=12, color="#F0E7E7"),
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(showgrid=False, tickangle=45),
        yaxis=dict(showgrid=True, gridcolor="rgba(200,200,200,0.3)"),
        margin=dict(l=50, r=50, t=80, b=100),
        showlegend=False
    )
    st.plotly_chart(fig_penghematan, use_container_width=True)
    
    # Metrik
    col1, col2 = st.columns(2)
    col1.metric("Total Biaya RAB", f"Rp {total_biaya:,.0f}")
    col2.metric("ROI Tahunan", f"{roi:.2f}%")
    
    # Unduh hasil
    csv = df_filtered.to_csv(index=False)
    st.download_button(
        label="📥 Unduh Hasil Analisis (CSV)",
        data=csv,
        file_name="hasil_analisis_rab_pln.csv",
        mime="text/csv"
    )
    
    # Kesimpulan
    st.subheader("Kesimpulan")
    avg_efisiensi = df_filtered["Efisiensi (%)"].mean()
    avg_penghematan = df_filtered["Penghematan Losses (kW)"].mean()
    
    if avg_efisiensi > 95:
        efisiensi_teks = f"Efisiensi sistem sangat baik ({avg_efisiensi:.2f}%)."
        efisiensi_color = COLORS["good"]
    elif avg_efisiensi > 90:
        efisiensi_teks = f"Efisiensi sistem baik ({avg_efisiensi:.2f}%)."
        efisiensi_color = COLORS["moderate"]
    else:
        efisiensi_teks = f"Efisiensi sistem perlu ditingkatkan ({avg_efisiensi:.2f}%)."
        efisiensi_color = COLORS["poor"]
    
    if avg_penghematan > 3:
        penghematan_teks = f"Penghematan rugi daya sangat baik ({avg_penghematan:.2f} kW)."
        penghematan_color = COLORS["very_good"]
    elif avg_penghematan > 1:
        penghematan_teks = f"Penghematan rugi daya cukup baik ({avg_penghematan:.2f} kW)."
        penghematan_color = COLORS["fair"]
    else:
        penghematan_teks = f"Penghematan rugi daya rendah ({avg_penghematan:.2f} kW), perlu optimasi."
        penghematan_color = COLORS["low"]
    
    if roi > 20:
        roi_teks = f"Proyek ini layak secara ekonomi (ROI {roi:.2f}%)."
        roi_color = COLORS["excellent"]
    elif roi > 10:
        roi_teks = f"Proyek cukup layak (ROI {roi:.2f}%)."
        roi_color = COLORS["adequate"]
    else:
        roi_teks = f"Proyek kurang layak (ROI {roi:.2f}%)."
        roi_color = COLORS["unfeasible"]
    
    st.markdown(f"<div style='background-color:{efisiensi_color}; padding:10px; border-radius:10px; margin-bottom:8px;'>{efisiensi_teks}</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='background-color:{penghematan_color}; padding:10px; border-radius:10px; margin-bottom:8px;'>{penghematan_teks}</div>", unsafe_allow_html=True)
    st.markdown(f"<div style='background-color:{roi_color}; padding:10px; border-radius:10px;'>{roi_teks}</div>", unsafe_allow_html=True)
    st.success("✅ Analisis selesai! Semua perhitungan dan kesimpulan telah ditampilkan.")
    st.markdown("---")
    st.caption("Dibuat dengan 💡 Streamlit | PLN ULP Batang © 2025")

# -------------------------------
# LOGIKA UTAMA
# -------------------------------
def main():
    """Menjalankan aplikasi utama."""
    set_background("tema2.jpg")
    set_sidebar_background("temasd1.jpg")
    display_header()
    config = configure_sidebar()
    
    uploaded_file = st.file_uploader("📂 Unggah File Excel Template RAB PLN (.xlsx)", type=["xlsx"])
    if uploaded_file:
        df_rab, df_gambar = validate_excel_file(uploaded_file)
        validate_columns(df_rab, df_gambar)
        validate_numeric(df_gambar, ["Panjang Jaringan (m)", "Beban Total (kVA)", "Tegangan (V)"])
        validate_kabel(df_gambar, config["selected_kabel"])
        df_gambar = preprocess_data(
            df_gambar,
            config["tipe_phase"],
            config["rugi_trafo"],
            config["baseline_losses"],
            config["core_loss"],
            config["full_load_loss"]
        )
        display_results(df_rab, df_gambar, config["tarif_kwh"], config["faktor_daya"], config["resistansi_kabel"])
    else:
        st.info("📥 Silakan unggah file Excel RAB PLN terlebih dahulu untuk memulai analisis.")

if __name__ == "__main__":

    main()
