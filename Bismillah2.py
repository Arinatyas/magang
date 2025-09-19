import streamlit as st
import pandas as pd
import altair as alt
import io

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# Upload beberapa file
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["csv", "xlsx", "ods"],
    accept_multiple_files=True
)

# Input header row
header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)",
    min_value=0, value=0, step=1
)

all_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name

        if filename.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, header=None)
        elif filename.endswith(".xlsx"):
            df_raw = pd.read_excel(uploaded_file, sheet_name=None, header=None)
            df_raw = pd.concat(df_raw.values(), ignore_index=True)
        elif filename.endswith(".ods"):
            df_raw = pd.read_excel(uploaded_file, engine="odf", sheet_name=None, header=None)
            df_raw = pd.concat(df_raw.values(), ignore_index=True)
        else:
            continue

        # Perbaiki header kosong dengan baris sebelumnya
        headers = df_raw.iloc[header_row].tolist()
        if header_row > 0:
            prev_row = df_raw.iloc[header_row - 1].tolist()
            headers = [
                headers[i] if pd.notna(headers[i]) and headers[i] != "" else prev_row[i]
                for i in range(len(headers))
            ]

        df = df_raw[(header_row + 1):].copy()
        df.columns = headers

        # Ganti data kosong
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = df[col].fillna("")

        all_data.append(df)

    # Gabungkan semua data
    data = pd.concat(all_data, ignore_index=True)

    st.subheader("🔎 Data Gabungan (setelah diproses)")
    st.write(data.head())

    # === FILTER MULTI-KOLOM ===
    st.subheader("🔎 Filter Data")
    filter_cols = st.multiselect("Pilih kolom untuk difilter", data.columns)

    filtered_data = data.copy()
    for col in filter_cols:
        values = st.multiselect(f"Pilih nilai untuk kolom {col}", data[col].unique())
        if values:
            filtered_data = filtered_data[filtered_data[col].isin(values)]

    st.subheader("📌 Hasil Penyaringan")
    st.write(filtered_data)

    # === UNDUH HASIL ===
    st.markdown("### ⬇️ Unduh Hasil Penyaringan")

    # CSV
    csv = filtered_data.to_csv(index=False).encode("utf-8")
    st.download_button("Unduh CSV", csv, file_name="hasil.csv", mime="text/csv")

    # Excel
    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        filtered_data.to_excel(writer, index=False, sheet_name="Hasil")
    st.download_button(
        "Unduh Excel (.xlsx)", 
        excel_buffer.getvalue(), 
        file_name="hasil.xlsx", 
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    # ODS
    ods_buffer = io.BytesIO()
    with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
        filtered_data.to_excel(writer, index=False, sheet_name="Hasil")
    st.download_button(
        "Unduh ODS", 
        ods_buffer.getvalue(), 
        file_name="hasil.ods", 
        mime="application/vnd.oasis.opendocument.spreadsheet"
    )

    # === VISUALISASI ===
    st.subheader("📈 Visualisasi Data")
    x_col = st.selectbox("Pilih kolom sumbu X", data.columns)
    y_col = st.selectbox("Pilih kolom sumbu Y", data.columns)

    chart_type = st.selectbox(
        "Pilih jenis grafik",
        ["Diagram Batang", "Diagram Garis", "Diagram Sebar"]
    )

    # Konversi Y ke numerik jika bisa
    try:
        filtered_data[y_col] = pd.to_numeric(filtered_data[y_col], errors="coerce")
    except:
        pass

    if pd.api.types.is_numeric_dtype(filtered_data[y_col]):
        if chart_type == "Diagram Batang":
            chart = alt.Chart(filtered_data).mark_bar().encode(x=x_col, y=y_col)
        elif chart_type == "Diagram Garis":
            chart = alt.Chart(filtered_data).mark_line().encode(x=x_col, y=y_col)
        else:
            chart = alt.Chart(filtered_data).mark_point().encode(x=x_col, y=y_col)
    else:
        chart = alt.Chart(filtered_data).mark_bar().encode(
            x=x_col,
            y="count()",
            color=y_col
        )

    st.altair_chart(chart, use_container_width=True)
    
