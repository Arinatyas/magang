import streamlit as st
import pandas as pd
import os
import altair as alt

st.set_page_config(page_title="📊 Gabung & Visualisasi Data", layout="wide")
st.title("📊 Gabung Data Excel/ODS + Visualisasi Aman")

# ======================
# Upload atau Folder
# ======================
mode = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])

data_frames = []

if mode == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)", 
        type=["xlsx", "xls", "ods"], 
        accept_multiple_files=True
    )

    if uploaded_files:
        header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0)
        for file in uploaded_files:
            try:
                sheets = pd.read_excel(file, sheet_name=None, header=header_row, engine="openpyxl")
            except Exception:
                sheets = pd.read_excel(file, sheet_name=None, header=header_row, engine="odf")

            for name, df in sheets.items():
                # Hapus kolom/baris kosong total
                df = df.dropna(axis=1, how='all')
                df = df.dropna(axis=0, how='all')

                # Bersihkan nama kolom
                df.columns = [
                    str(c).strip() if str(c).strip() not in ["", "None", "nan"] else f"Kolom_{i+1}"
                    for i, c in enumerate(df.columns)
                ]

                # Tambahkan info asal
                df["__SHEET__"] = name
                df["__FILE__"] = file.name

                data_frames.append(df)

elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")
    header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0)

    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                try:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=header_row, engine="openpyxl")
                except Exception:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=header_row, engine="odf")

                for name, df in sheets.items():
                    df = df.dropna(axis=1, how='all')
                    df = df.dropna(axis=0, how='all')
                    df.columns = [
                        str(c).strip() if str(c).strip() not in ["", "None", "nan"] else f"Kolom_{i+1}"
                        for i, c in enumerate(df.columns)
                    ]
                    df["__SHEET__"] = name
                    df["__FILE__"] = fname
                    data_frames.append(df)

# ======================
# Gabungan Data Aman
# ======================
if data_frames:
    # Gabungkan semua sheet/file dengan union kolom
    data_gabungan = pd.concat(data_frames, ignore_index=True, sort=True)

    # Bersihkan hanya yang benar-benar kosong
    data_gabungan = data_gabungan.dropna(axis=0, how='all')
    data_gabungan = data_gabungan.dropna(axis=1, how='all')

    # Perbaiki nama kolom
    clean_cols = []
    for i, c in enumerate(data_gabungan.columns):
        name = str(c).strip()
        if name in ["", "None", "nan", "NaN", "Unnamed: 0"]:
            name = f"Kolom_{i+1}"
        clean_cols.append(name)
    data_gabungan.columns = clean_cols

    # Hapus duplikasi baris identik
    data_gabungan = data_gabungan.drop_duplicates()

    st.subheader("📄 Data Gabungan (Hasil Bersih)")
    st.dataframe(data_gabungan, use_container_width=True)

    # ======================
    # Filter Data
    # ======================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

    filtered_df = data_gabungan.copy()

    for kol in filter_columns:
        unique_vals = filtered_df[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
        if pilihan:
            filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df, use_container_width=True)

    # ======================
    # Unduh Data
    # ======================
    st.subheader("💾 Unduh Data Gabungan")
    out_excel = "data_gabungan.xlsx"
    out_ods = "data_gabungan.ods"

    filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
    filtered_df.to_excel(out_ods, index=False, engine="odf")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)

    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)

    # ======================
    # Visualisasi
    # ======================
    if not filtered_df.empty:
        st.subheader("📈 Visualisasi Data")

        # Kolom bersih
        all_cols = filtered_df.columns.tolist()

        x_axis = st.selectbox("Pilih kolom sumbu X", all_cols, key="x_axis")
        y_axis = st.selectbox("Pilih kolom sumbu Y", [c for c in all_cols if c != x_axis], key="y_axis")

        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        # Deteksi tipe data
        def detect_type(col):
            if pd.api.types.is_numeric_dtype(filtered_df[col]):
                return "quantitative"
            elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                return "temporal"
            else:
                return "nominal"

        x_type = detect_type(x_axis)
        y_type = detect_type(y_axis)
        tooltip_cols = [alt.Tooltip(c, type=detect_type(c)) for c in all_cols]

        if chart_type == "Diagram Batang":
            chart = alt.Chart(filtered_df).mark_bar().encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=tooltip_cols
            )
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=tooltip_cols
            )
        else:
            chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=tooltip_cols
            )

        st.altair_chart(chart, use_container_width=True)
                  
