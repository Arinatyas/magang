import streamlit as st
import pandas as pd
import altair as alt
import io

# =====================
# Judul Aplikasi
# =====================
st.title("📂 Aplikasi Pengelolaan Data Dokumen")
st.write("Unggah beberapa file (.csv, .xlsx, .ods), gabungkan, filter, visualisasikan, dan unduh hasilnya.")

# =====================
# Upload Beberapa File
# =====================
uploaded_files = st.file_uploader(
    "Unggah file (boleh lebih dari satu)", 
    type=["csv", "xlsx", "ods"], 
    accept_multiple_files=True
)

if uploaded_files:
    dataframes = []
    for file in uploaded_files:
        if file.name.endswith(".csv"):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file, sheet_name=None)  # bisa ada banyak sheet
            # Gabung semua sheet dalam satu file
            df = pd.concat(df.values(), ignore_index=True)
        dataframes.append(df)

    # Gabungkan semua file jadi satu tabel
    combined_df = pd.concat(dataframes, ignore_index=True)
    st.subheader("📑 Data Gabungan")
    st.dataframe(combined_df)

    # =====================
    # Fitur Filter Dinamis
    # =====================
    st.subheader("🔍 Filter Data")
    filter_columns = st.multiselect("Pilih kolom untuk filter", combined_df.columns)

    filtered_df = combined_df.copy()
    for col in filter_columns:
        unique_vals = filtered_df[col].unique()
        selected_vals = st.multiselect(f"Pilih nilai untuk kolom **{col}**", unique_vals)
        if selected_vals:
            filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

    st.subheader("📋 Hasil Data Setelah Filter")
    st.dataframe(filtered_df)

    # =====================
    # Visualisasi Data
    # =====================
    if not filtered_df.empty:
        st.subheader("📊 Visualisasi Data")

        all_cols = [c.strip() for c in filtered_df.columns.tolist()]
        filtered_df.columns = all_cols

        # Pilih kolom untuk X dan Y
        x_axis = st.selectbox("Pilih kolom sumbu X", all_cols)
        y_axis = st.selectbox("Pilih kolom sumbu Y", all_cols)

        chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        # Tentukan tipe data otomatis
        if pd.api.types.is_numeric_dtype(filtered_df[x_axis]):
            x_type = "quantitative"
        elif pd.api.types.is_datetime64_any_dtype(filtered_df[x_axis]):
            x_type = "temporal"
        else:
            x_type = "nominal"

        if pd.api.types.is_numeric_dtype(filtered_df[y_axis]):
            y_type = "quantitative"
        elif pd.api.types.is_datetime64_any_dtype(filtered_df[y_axis]):
            y_type = "temporal"
        else:
            y_type = "nominal"

        # Buat grafik
        if chart_type == "Diagram Batang":
            chart = alt.Chart(filtered_df).mark_bar().encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filtered_df.columns.tolist()
            )
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(filtered_df).mark_line(point=True).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filtered_df.columns.tolist()
            )
        else:  # Diagram Sebar
            chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                x=alt.X(x_axis, type=x_type),
                y=alt.Y(y_axis, type=y_type),
                tooltip=filtered_df.columns.tolist()
            )

        st.altair_chart(chart, use_container_width=True)

    # =====================
    # Unduh Hasil
    # =====================
    st.subheader("💾 Unduh Hasil Data")

    # Unduh Excel
    towrite = io.BytesIO()
    with pd.ExcelWriter(towrite, engine="xlsxwriter") as writer:
        filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
    towrite.seek(0)
    st.download_button(
        label="📥 Unduh ke Excel (.xlsx)",
        data=towrite,
        file_name="hasil_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # Unduh CSV
    csv = filtered_df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="📥 Unduh ke CSV (.csv)",
        data=csv,
        file_name="hasil_data.csv",
        mime="text/csv"
    )
