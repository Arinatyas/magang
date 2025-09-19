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

# Input header row (default = 0 kalau tidak diisi)
header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)",
    min_value=0, step=1, value=0
)

all_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        filename = uploaded_file.name

        # === CSV ===
        if filename.endswith(".csv"):
            df_raw = pd.read_csv(uploaded_file, header=None)
            sheets = {"Sheet1": df_raw}

        # === XLSX ===
        elif filename.endswith(".xlsx"):
            xls = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="openpyxl")
            sheets = xls

        # === ODS ===
        elif filename.endswith(".ods"):
            ods = pd.read_excel(uploaded_file, sheet_name=None, header=None, engine="odf")
            sheets = ods
        else:
            continue

        # Proses semua sheet → gabungkan
        for sheet_name, df_raw in sheets.items():
            # Tentukan header (kalau lebih dari jumlah baris, fallback ke 0)
            h = header_row if header_row < len(df_raw) else 0

            # Ambil header dari baris ke-h
            headers = df_raw.iloc[h].tolist()

            # Kalau ada kosong → ganti dengan baris sebelumnya
            if h > 0:
                prev_row = df_raw.iloc[h - 1].tolist()
                headers = [
                    headers[i] if pd.notna(headers[i]) and headers[i] != "" else prev_row[i]
                    for i in range(len(headers))
                ]

            df = df_raw[(h + 1):].copy()
            df.columns = headers

            # Isi kosong (0 untuk numerik, "" untuk kategori)
            for col in df.columns:
                if pd.api.types.is_numeric_dtype(df[col]):
                    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                else:
                    df[col] = df[col].fillna("")

            # Tambahkan metadata sumber
            df["Sumber_File"] = filename
            df["Sumber_Sheet"] = sheet_name

            all_data.append(df)

    # Gabungkan semua sheet + file jadi satu
    if all_data:
        data = pd.concat(all_data, ignore_index=True)

        st.subheader("🔎 Data Gabungan (setelah diproses)")
        st.dataframe(data.head())

        # === FILTER ===
        st.subheader("🔎 Filter Data")

        filter_cols = st.multiselect(
            "Pilih kolom untuk difilter",
            data.columns.tolist()
        )

        filtered_data = data.copy()
        for col in filter_cols:
            values = st.multiselect(f"Pilih nilai untuk {col}", data[col].unique())
            if values:
                filtered_data = filtered_data[filtered_data[col].isin(values)]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_data)

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
        if not filtered_data.empty:
            cols = [c for c in filtered_data.columns if c not in ["Sumber_File", "Sumber_Sheet"]]
            if len(cols) >= 2:
                x_col = st.selectbox("Pilih kolom sumbu X", cols)
                y_col = st.selectbox("Pilih kolom sumbu Y", cols)

                chart_type = st.selectbox(
                    "Pilih jenis grafik",
                    ["Diagram Batang", "Diagram Garis", "Diagram Sebar"]
                )

                try:
                    filtered_data[y_col] = pd.to_numeric(filtered_data[y_col], errors="coerce")
                except:
                    pass

                if chart_type == "Diagram Batang":
                    chart = alt.Chart(filtered_data).mark_bar().encode(x=x_col, y=y_col)
                elif chart_type == "Diagram Garis":
                    chart = alt.Chart(filtered_data).mark_line().encode(x=x_col, y=y_col)
                else:
                    chart = alt.Chart(filtered_data).mark_point().encode(x=x_col, y=y_col)

                st.altair_chart(chart, use_container_width=True)

                # Unduh visualisasi PNG
                try:
                    chart_bytes = chart.save(fp=None, format="png")
                    st.download_button(
                        "Unduh Gambar PNG",
                        chart_bytes,
                        file_name="visualisasi.png",
                        mime="image/png"
                    )
                except:
                    st.info("Gunakan klik kanan grafik untuk unduh gambar.")
                
