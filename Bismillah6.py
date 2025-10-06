import streamlit as st
import pandas as pd
import altair as alt
import io

st.set_page_config(page_title="📊 Gabung & Visualisasi Data ODS/Excel", layout="wide")
st.title("📊 Gabung Data Excel/ODS + Visualisasi Aman")

# ======== Upload ========
uploaded_files = st.file_uploader(
    "Upload beberapa file Excel/ODS", 
    type=["xlsx", "xls", "ods"], 
    accept_multiple_files=True
)

data_frames = []

if uploaded_files:
    for file in uploaded_files:
        try:
            sheets = pd.read_excel(file, sheet_name=None, engine="openpyxl")
        except:
            sheets = pd.read_excel(file, sheet_name=None, engine="odf")

        for sheet_name, df in sheets.items():
            # Bersihkan header kosong
            df = df.dropna(how="all").reset_index(drop=True)
            df.columns = [
                f"Kolom_{i+1}" if ("Unnamed" in str(c) or pd.isna(c)) else str(c).strip()
                for i, c in enumerate(df.columns)
            ]
            df["__SHEET__"] = sheet_name
            df["__FILE__"] = file.name
            data_frames.append(df)

    # ======== Gabungkan Semua Sheet ========
    data_gabungan = pd.concat(data_frames, ignore_index=True, sort=False)
    data_gabungan = data_gabungan.dropna(how="all")
    st.subheader("📄 Data Gabungan (Semua Sheet)")
    st.dataframe(data_gabungan, use_container_width=True)

    # ======== Filter dengan Penanda ========
    st.subheader("🔍 Filter (Tandai Data Sesuai Pilihan)")
    filter_col = st.selectbox("Pilih kolom untuk filter", data_gabungan.columns)
    if filter_col:
        nilai_unik = sorted(data_gabungan[filter_col].dropna().unique().tolist())
        pilihan = st.multiselect(f"Pilih nilai dari kolom '{filter_col}'", nilai_unik)

        if pilihan:
            data_gabungan["TERPILIH"] = data_gabungan[filter_col].isin(pilihan)
        else:
            data_gabungan["TERPILIH"] = False

        st.dataframe(data_gabungan, use_container_width=True)

    # ======== Unduh Data ========
    st.subheader("💾 Unduh Data Gabungan")
    towrite = io.BytesIO()
    data_gabungan.to_excel(towrite, index=False, engine="openpyxl")
    towrite.seek(0)
    st.download_button(
        "📥 Unduh Excel Gabungan", 
        towrite, 
        file_name="data_gabungan.xlsx"
    )

    # ======== Visualisasi Aman ========
    st.subheader("📈 Visualisasi Data")

    all_cols = [c for c in data_gabungan.columns if c not in ["__SHEET__", "__FILE__", "TERPILIH"]]
    if len(all_cols) >= 2:
        x_axis = st.selectbox("Sumbu X", all_cols)
        y_axis = st.selectbox("Sumbu Y", all_cols, index=1)

        jenis_chart = st.radio("Pilih Jenis Grafik", ["Batang", "Garis", "Sebar"])

        # Konversi numerik jika bisa
        for col in [x_axis, y_axis]:
            try:
                data_gabungan[col] = pd.to_numeric(data_gabungan[col], errors="ignore")
            except:
                pass

        base = alt.Chart(data_gabungan).encode(
            x=alt.X(x_axis, title=x_axis),
            y=alt.Y(y_axis, title=y_axis),
            color=alt.condition(
                alt.datum.TERPILIH == True,
                alt.value("red"), alt.value("steelblue")
            ),
            tooltip=list(data_gabungan.columns)
        )

        if jenis_chart == "Batang":
            chart = base.mark_bar()
        elif jenis_chart == "Garis":
            chart = base.mark_line(point=True)
        else:
            chart = base.mark_circle(size=80)

        st.altair_chart(chart.interactive(), use_container_width=True)
