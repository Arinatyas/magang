import streamlit as st
import pandas as pd
import os
import altair as alt
from io import BytesIO
from odf.opendocument import OpenDocumentSpreadsheet
from odf.table import Table, TableRow, TableCell
from odf.text import P

st.set_page_config(page_title="📊 Gabung Data Excel/ODS + Visualisasi", layout="wide")

st.title("📊 Gabung Data Excel/ODS + Visualisasi")

# =====================
# Upload File / Folder
# =====================
source_choice = st.radio("Pilih sumber data:", ["Upload File", "Pilih Folder"])

uploaded_files = []
folder_path = ""

if source_choice == "Upload File":
    uploaded_files = st.file_uploader(
        "Upload file Excel/ODS (bisa banyak)",
        type=["xlsx", "xls", "ods"],
        accept_multiple_files=True
    )
else:
    folder_path = st.text_input("Masukkan path folder yang berisi Excel/ODS")
    if folder_path and os.path.isdir(folder_path):
        uploaded_files = [os.path.join(folder_path, f) for f in os.listdir(folder_path)
                          if f.endswith((".xlsx", ".xls", ".ods"))]

# =====================
# Gabungkan Data
# =====================
df = pd.DataFrame()

if uploaded_files:
    header_row = st.number_input("Pilih baris header (mulai dari 0)", min_value=0, value=0)

    all_data = []
    for file in uploaded_files:
        try:
            excel_obj = pd.ExcelFile(file)
            for sheet_name in excel_obj.sheet_names:
                temp_df = pd.read_excel(excel_obj, sheet_name=sheet_name, header=header_row)

                # --- Ganti kolom kosong jadi "kosong" ---
                temp_df.columns = [
                    str(col) if str(col).strip() not in ["", "None", "nan", "NaN"] else "kosong"
                    for col in temp_df.columns
                ]

                # --- Ganti nilai kosong dengan 0 ---
                temp_df = temp_df.fillna(0)

                # --- Tambah info sumber ---
                temp_df["source_file"] = os.path.basename(str(file))  # hanya nama file
                temp_df["source_sheet"] = sheet_name
                all_data.append(temp_df)
        except Exception as e:
            st.error(f"Gagal membaca {file}: {e}")

    if all_data:
        df = pd.concat(all_data, ignore_index=True)

if not df.empty:
    st.subheader("📄 Data Gabungan")
    st.dataframe(df)

    # =====================
    # Rename Unnamed Columns (SEBELUM FILTER)
    # =====================
    rename_map = {}
    for col in df.columns:
        if "Unnamed" in str(col) or col == "kosong":
            new_name = st.text_input(f"Kolom '{col}' ingin diganti jadi apa?", value="")
            if new_name.strip():
                rename_map[col] = new_name
    if rename_map:
        df = df.rename(columns=rename_map)

    # =====================
    # Filter Data
    # =====================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", df.columns.tolist())

    filtered_df = df.copy()
    for col in filter_columns:
        unique_vals = filtered_df[col].dropna().unique().tolist()
        selected_vals = st.multiselect(f"Pilih nilai untuk {col}", unique_vals)
        if selected_vals:
            filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

    st.write("Data Setelah Penyaringan")
    st.dataframe(filtered_df)

    # =====================
    # Visualisasi Data
    # =====================
    if not filtered_df.empty:
        st.subheader("📈 Visualisasi Data")

        all_cols = [c.strip() for c in filtered_df.columns.tolist()]
        filtered_df.columns = all_cols

        x_axis = st.selectbox("Pilih kolom sumbu X", [""] + all_cols, index=0)
        y_axis = st.selectbox("Pilih kolom sumbu Y", [""] + all_cols, index=0)

        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        if x_axis and y_axis:  # grafik hanya tampil jika sudah pilih X dan Y
            def detect_type(col):
                if pd.api.types.is_numeric_dtype(filtered_df[col]):
                    return "quantitative"
                elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                    return "temporal"
                else:
                    return "nominal"

            x_type = detect_type(x_axis)
            y_type = detect_type(y_axis)

            try:
                if chart_type == "Diagram Batang":
                    chart = alt.Chart(filtered_df).mark_bar().encode(
                        x=alt.X(x_axis, type=x_type),
                        y=alt.Y(y_axis, type=y_type),
                        tooltip=all_cols
                    )
                elif chart_type == "Diagram Garis":
                    chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                        x=alt.X(x_axis, type=x_type),
                        y=alt.Y(y_axis, type=y_type),
                        tooltip=all_cols
                    )
                else:  # Diagram Sebar
                    chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                        x=alt.X(x_axis, type=x_type),
                        y=alt.Y(y_axis, type=y_type),
                        tooltip=all_cols
                    )

                st.altair_chart(chart, use_container_width=True)
            except Exception as e:
                st.error(f"Gagal menampilkan grafik: {e}")
        else:
            st.info("Pilih kolom X dan Y terlebih dahulu untuk menampilkan grafik.")

else:
    st.info("Silakan upload file atau pilih folder terlebih dahulu.")
              
