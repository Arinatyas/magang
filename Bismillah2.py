import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# Fungsi baca file dengan pilihan header baris
def baca_file(file, header_row):
    ext = file.name.split(".")[-1].lower()
    all_dfs = []

    try:
        if ext in ["xlsx", "ods"]:
            xls = pd.ExcelFile(file, engine="openpyxl" if ext=="xlsx" else "odf")
            for sheet in xls.sheet_names:
                try:
                    df_sheet = pd.read_excel(
                        file,
                        sheet_name=sheet,
                        header=header_row,
                        engine="openpyxl" if ext=="xlsx" else "odf"
                    )

                    # Buang baris kosong
                    df_sheet = df_sheet.dropna(how="all")

                    # Isi NaN: numerik → 0, kategorik → "data kosong"
                    for col in df_sheet.columns:
                        if pd.api.types.is_numeric_dtype(df_sheet[col]):
                            df_sheet[col] = df_sheet[col].fillna(0)
                        else:
                            df_sheet[col] = df_sheet[col].fillna("data kosong")

                    df_sheet["Sumber_File"] = file.name
                    df_sheet["Sumber_Sheet"] = sheet
                    all_dfs.append(df_sheet)

                except Exception as e:
                    st.warning(f"Gagal membaca sheet {sheet} dari {file.name}: {e}")
        elif ext == "csv":
            df_csv = pd.read_csv(file, header=header_row)
            df_csv["Sumber_File"] = file.name
            df_csv["Sumber_Sheet"] = "CSV"
            all_dfs.append(df_csv)
        else:
            st.warning(f"Format {file.name} tidak didukung")
    except Exception as e:
        st.error(f"Gagal membaca file {file.name}: {e}")

    return all_dfs


# Upload file
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

# Pilih baris header manual
header_row = st.number_input(
    "Pilih baris header (mulai dari 0, misalnya baris ke-6 berarti 5)",
    min_value=0, value=0, step=1
)

if uploaded_files:
    semua_data = []
    for file in uploaded_files:
        semua_data.extend(baca_file(file, header_row))

    if semua_data:
        all_data = pd.concat(semua_data, ignore_index=True)
        st.subheader("🔎 Data Gabungan")
        st.dataframe(all_data.head())

        # Filter dinamis
        filter_columns = st.multiselect("Pilih kolom yang ingin difilter & ditampilkan", all_data.columns.tolist())
        filtered_df = all_data.copy()

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]

            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # Visualisasi
        if filter_columns and not filtered_df.empty:
            st.subheader("📈 Visualisasi Data")
            all_cols = filtered_df.columns.tolist()

            x_axis = st.selectbox("Pilih kolom sumbu X", all_cols)
            y_axis = st.selectbox("Pilih kolom sumbu Y", all_cols)

            chart_type = st.radio("Pilih jenis grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

            def tipe_kolom(col):
                if pd.api.types.is_numeric_dtype(filtered_df[col]):
                    return "quantitative"
                elif pd.api.types.is_datetime64_any_dtype(filtered_df[col]):
                    return "temporal"
                else:
                    return "nominal"

            x_type = tipe_kolom(x_axis)
            y_type = tipe_kolom(y_axis)

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

        # Unduh hasil
        st.subheader("💾 Unduh Hasil")

        if not filtered_df.empty:
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button("Unduh sebagai Excel (.xlsx)", excel_buffer.getvalue(),
                               "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


