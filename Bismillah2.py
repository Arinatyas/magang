import streamlit as st
import pandas as pd
import altair as alt
import io

st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")
st.write("Unggah beberapa file (Excel/CSV/ODS) sekaligus")

# ====================
# Upload file
# ====================
uploaded_files = st.file_uploader(
    "Upload file (Excel/CSV/ODS)", 
    type=["xlsx", "csv", "ods"], 
    accept_multiple_files=True
)

header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)", 
    min_value=0, 
    value=0
)

semua_data = []

if uploaded_files:
    for file in uploaded_files:
        if file.name.endswith(".xlsx"):
            xls = pd.ExcelFile(file)
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row)
                semua_data.append(df)
        elif file.name.endswith(".ods"):
            xls = pd.ExcelFile(file, engine="odf")
            for sheet in xls.sheet_names:
                df = pd.read_excel(file, sheet_name=sheet, header=header_row, engine="odf")
                semua_data.append(df)
        elif file.name.endswith(".csv"):
            df = pd.read_csv(file, header=header_row)
            semua_data.append(df)

if semua_data:
    all_data = pd.concat(semua_data, ignore_index=True)

    # ========================
    # Bersihkan nama kolom
    # ========================
    new_columns = []
    last_valid = "Kolom"
    for col in all_data.columns:
        if str(col).startswith("Unnamed") or str(col).strip() == "":
            new_columns.append(last_valid)  # pakai nama sebelumnya
        else:
            new_columns.append(col)
            last_valid = col

    # Pastikan nama unik
    def make_unique(seq):
        seen = {}
        result = []
        for item in seq:
            if item in seen:
                seen[item] += 1
                result.append(f"{item}_{seen[item]}")
            else:
                seen[item] = 0
                result.append(item)
        return result

    all_data.columns = make_unique(new_columns)

    # ========================
    # Isi data kosong
    # ========================
    for col in all_data.columns:
        if all_data[col].dtype == "object":
            all_data[col] = all_data[col].fillna("Kosong")
        else:
            all_data[col] = all_data[col].fillna(0)

    # ========================
    # Tampilkan data gabungan
    # ========================
    st.subheader("🔎 Data Gabungan")
    st.dataframe(all_data.head())

    # ========================
    # Filter kolom
    # ========================
    st.subheader("Pilih kolom yang ingin difilter & ditampilkan")
    filter_columns = st.multiselect("Pilih kolom:", all_data.columns)

    filtered_data = all_data.copy()
    for col in filter_columns:
        unique_vals = filtered_data[col].unique().tolist()
        selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
        if selected_vals:
            filtered_data = filtered_data[filtered_data[col].isin(selected_vals)]

    st.subheader("📌 Hasil Penyaringan")
    st.dataframe(filtered_data.head())

    # ========================
    # Visualisasi
    # ========================
    st.subheader("📈 Visualisasi Data")
    x_axis = st.selectbox("Pilih kolom sumbu X", all_data.columns)
    y_axis = st.selectbox("Pilih kolom sumbu Y", all_data.columns)
    chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

    chart = None
    if chart_type == "Diagram Batang":
        chart = alt.Chart(filtered_data).mark_bar().encode(
            x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
        )
    elif chart_type == "Diagram Garis":
        chart = alt.Chart(filtered_data).mark_line(point=True).encode(
            x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
        )
    elif chart_type == "Diagram Sebar":
        chart = alt.Chart(filtered_data).mark_circle(size=60).encode(
            x=x_axis, y=y_axis, tooltip=list(filtered_data.columns)
        )

    if chart:
        st.altair_chart(chart, use_container_width=True)

        # Simpan grafik PNG / SVG
        if st.button("💾 Simpan Grafik PNG"):
            chart.save("chart.png")
            with open("chart.png", "rb") as f:
                st.download_button("Unduh PNG", f, "chart.png")

        if st.button("💾 Simpan Grafik SVG"):
            chart.save("chart.svg")
            with open("chart.svg", "rb") as f:
                st.download_button("Unduh SVG", f, "chart.svg")
        )
        filtered_df = all_data.copy()

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

            # tampilkan hanya kolom yang difilter
            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # ========================
        # Visualisasi
        # ========================
        if filter_columns and not filtered_df.empty:
            st.subheader("📈 Visualisasi Data")
            all_cols = filtered_df.columns.tolist()

            # Pilih X dan Y
            x_axis = st.selectbox("Pilih kolom sumbu X", all_cols)
            y_axis = st.selectbox("Pilih kolom sumbu Y", all_cols)

            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

            x_type = tipe_kolom(x_axis, filtered_df)
            y_type = tipe_kolom(y_axis, filtered_df)

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
            else:
                chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )

            st.altair_chart(chart, use_container_width=True)

            # ========================
            # Unduh visualisasi (SVG)
            # ========================
            st.subheader("💾 Unduh Visualisasi")
            try:
                svg = chart.to_svg()
                st.download_button(
                    "Unduh Grafik sebagai SVG",
                    svg,
                    "visualisasi.svg",
                    "image/svg+xml"
                )
            except Exception as e:
                st.warning(f"Gagal mengekspor visualisasi: {e}")

        # ========================
        # Unduh hasil
        # ========================
        st.subheader("💾 Unduh Data")

        if not filtered_df.empty:
            # CSV
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            # Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai Excel (.xlsx)", excel_buffer.getvalue(),
                "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

            # ODS
            ods_buffer = BytesIO()
            with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button(
                "Unduh sebagai ODS (.ods)", ods_buffer.getvalue(),
                "hasil_filter.ods", "application/vnd.oasis.opendocument.spreadsheet"
            )
            
