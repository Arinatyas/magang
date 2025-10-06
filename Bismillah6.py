import streamlit as st
import pandas as pd
import os
import altair as alt

st.title("📊 Gabung Data Excel/ODS + Visualisasi")

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
        for uploaded_file in uploaded_files:
            try:
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")
            except:
                sheets = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="odf")
            for name, df in sheets.items():
                df["__SHEET__"] = name
                df["__FILE__"] = uploaded_file.name
                data_frames.append(df)

elif mode == "Pilih Folder":
    folder = st.text_input("Masukkan path folder (isi file .xlsx/.ods)")

    if folder and os.path.isdir(folder):
        for fname in os.listdir(folder):
            if fname.endswith((".xlsx", ".xls", ".ods")):
                fpath = os.path.join(folder, fname)
                try:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=None, engine="openpyxl")
                except:
                    sheets = pd.read_excel(fpath, sheet_name=None, header=None, engine="odf")
                for name, df in sheets.items():
                    df["__SHEET__"] = name
                    df["__FILE__"] = fname
                    data_frames.append(df)

# ======================
# Gabungan Data
# ======================
if data_frames:
    data_gabungan = pd.concat(data_frames, ignore_index=True)

    st.subheader("📄 Data Gabungan")
    st.dataframe(data_gabungan)

    # ======================
    # Filter Data
    # ======================
    st.subheader("🔍 Penyaringan Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", data_gabungan.columns)

    filtered_df = data_gabungan.copy()
    tampilkan_kolom = []

    for kol in filter_columns:
        unique_vals = filtered_df[kol].dropna().unique().tolist()
        pilihan = st.multiselect(f"Pilih nilai untuk {kol}", unique_vals)
        tampilkan_kolom.append(kol)
        if pilihan:
            filtered_df = filtered_df[filtered_df[kol].isin(pilihan)]

    if tampilkan_kolom:
        filtered_df = filtered_df[tampilkan_kolom]

    st.write("### Data Setelah Penyaringan")
    st.dataframe(filtered_df)

    # ======================
    # Unduh Data
    # ======================
    st.subheader("💾 Unduh Data")
    out_excel = "data_gabungan.xlsx"
    out_ods = "data_gabungan.ods"
    filtered_df.to_excel(out_excel, index=False, engine="openpyxl")
    filtered_df.to_excel(out_ods, index=False, engine="odf")

    with open(out_excel, "rb") as f:
        st.download_button("📥 Unduh Excel (.xlsx)", f, file_name=out_excel)

    with open(out_ods, "rb") as f:
        st.download_button("📥 Unduh ODS (.ods)", f, file_name=out_ods)

    # ======================
    # Visualisasi Grafik
    # =====================
    if filter_columns and not filtered_df.empty:
        st.subheader("📈 Visualisasi Data")

        all_cols = filtered_df.columns.tolist()
        x_axis = st.selectbox("Pilih kolom sumbu X", all_cols, key="x_axis")
        y_axis = st.selectbox("Pilih kolom sumbu Y", [c for c in all_cols if c != x_axis], key="y_axis")

        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        # 🔹 Pembersihan aman sebelum grafik
        df_viz = filtered_df.copy()
        df_viz = df_viz.dropna(subset=[x_axis, y_axis])  # hilangkan baris kosong di X/Y
        df_viz[x_axis] = df_viz[x_axis].astype(str)  # ubah X ke string agar aman untuk label
        df_viz[y_axis] = pd.to_numeric(df_viz[y_axis], errors="coerce")  # coba ubah Y jadi angka
        df_viz = df_viz.dropna(subset=[y_axis])  # buang jika Y bukan angka

        if df_viz.empty:
            st.warning("⚠️ Tidak ada data valid untuk divisualisasikan. Coba pilih kolom lain.")
        else:
            # Fungsi deteksi tipe kolom
            def detect_type(col):
                try:
                    if pd.api.types.is_numeric_dtype(df_viz[col]):
                        return "quantitative"
                    elif pd.api.types.is_datetime64_any_dtype(df_viz[col]):
                        return "temporal"
                    else:
                        return "nominal"
                except Exception:
                    return "nominal"

            x_type = detect_type(x_axis)
            y_type = detect_type(y_axis)
            tooltip_cols = [alt.Tooltip(c, type=detect_type(c)) for c in df_viz.columns]

            if chart_type == "Diagram Batang":
                chart = alt.Chart(df_viz).mark_bar().encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=tooltip_cols
                )
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(df_viz).mark_line(point=True).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=tooltip_cols
                )
            else:
                chart = alt.Chart(df_viz).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=tooltip_cols
                )

            st.altair_chart(chart, use_container_width=True)
