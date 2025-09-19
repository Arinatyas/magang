import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# ===========================
# Fungsi Baca & Bersihkan File
# ===========================
def baca_file(file):
    ext = file.name.split(".")[-1].lower()
    all_dfs = []

    if ext in ["xlsx", "ods"]:
        xls = pd.ExcelFile(file, engine="openpyxl" if ext=="xlsx" else "odf")
        for sheet in xls.sheet_names:
            try:
                # Baca sheet tanpa header
                df_raw = pd.read_excel(
                    file, sheet_name=sheet, header=None,
                    engine="openpyxl" if ext=="xlsx" else "odf"
                )

                # Cari baris dengan isi terbanyak → jadi header
                filled_counts = df_raw.notna().sum(axis=1)
                header_row = filled_counts.idxmax()

                # Ambil baris itu sebagai header
                new_header = df_raw.iloc[header_row].fillna(method="ffill").fillna(method="bfill").tolist()

                # Data mulai setelah header
                df = df_raw.iloc[header_row+1:].copy()
                df.columns = [str(col) if str(col) != "nan" else f"Kolom_{i+1}" 
                              for i, col in enumerate(new_header)]

                # Bersihkan NaN → 0 utk numerik, "data kosong" utk kategorik
                for col in df.columns:
                    if pd.api.types.is_numeric_dtype(df[col]):
                        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
                    else:
                        df[col] = df[col].fillna("data kosong")

                df["Sumber_File"] = file.name
                df["Sumber_Sheet"] = sheet
                all_dfs.append(df)

            except Exception as e:
                st.warning(f"Gagal membaca sheet {sheet} dari {file.name}: {e}")

    elif ext == "csv":
        df = pd.read_csv(file)
        for col in df.columns:
            if pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
            else:
                df[col] = df[col].fillna("data kosong")
        df["Sumber_File"] = file.name
        df["Sumber_Sheet"] = "CSV"
        all_dfs.append(df)

    else:
        st.warning(f"Format {file.name} tidak didukung")

    return all_dfs


# ===========================
# Upload File
# ===========================
uploaded_files = st.file_uploader(
    "Unggah beberapa file (Excel/CSV/ODS) sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

if uploaded_files:
    semua_data = []
    for file in uploaded_files:
        semua_data.extend(baca_file(file))

    if semua_data:
        all_data = pd.concat(semua_data, ignore_index=True)
        st.subheader("🔎 Data Gabungan")
        st.dataframe(all_data.head())

        # ===========================
        # Filter Data
        # ===========================
        filter_columns = st.multiselect("Pilih kolom yang ingin difilter & ditampilkan", all_data.columns.tolist())

        # 🔹 Filter tambahan: pilih sheet
        sheets_filter = st.multiselect("Filter berdasarkan Sheet", all_data["Sumber_Sheet"].unique())
        filtered_df = all_data.copy()

        if sheets_filter:
            filtered_df = filtered_df[filtered_df["Sumber_Sheet"].isin(sheets_filter)]

        if filter_columns:
            for col in filter_columns:
                unique_vals = filtered_df[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Pilih nilai untuk kolom {col}", unique_vals)
                if selected_vals:
                    filtered_df = filtered_df[filtered_df[col].isin(selected_vals)]
            filtered_df = filtered_df[filter_columns]

        st.subheader("📌 Hasil Penyaringan")
        st.dataframe(filtered_df)

        # ===========================
        # Visualisasi
        # ===========================
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

            chart = None
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
            else:  # Diagram Sebar
                chart = alt.Chart(filtered_df).mark_circle(size=60).encode(
                    x=alt.X(x_axis, type=x_type),
                    y=alt.Y(y_axis, type=y_type),
                    tooltip=all_cols
                )

            st.altair_chart(chart, use_container_width=True)

            # 🔹 Unduh grafik sebagai PNG
            try:
                from altair_saver import save
                save(chart, "chart.png", fmt="png")
                with open("chart.png", "rb") as f:
                    st.download_button("📥 Unduh Grafik PNG", f, file_name="visualisasi.png", mime="image/png")
            except Exception:
                st.info("Untuk unduh grafik sebagai PNG, install package `altair_saver` (pip install altair_saver).")

        # ===========================
        # Unduh Data Hasil Filter
        # ===========================
        st.subheader("💾 Unduh Hasil")

        if not filtered_df.empty:
            # CSV
            csv = filtered_df.to_csv(index=False).encode("utf-8")
            st.download_button("Unduh sebagai CSV", csv, "hasil_filter.csv", "text/csv")

            # Excel
            excel_buffer = BytesIO()
            with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button("Unduh sebagai Excel (.xlsx)", excel_buffer.getvalue(),
                               "hasil_filter.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

            # ODS
            ods_buffer = BytesIO()
            with pd.ExcelWriter(ods_buffer, engine="odf") as writer:
                filtered_df.to_excel(writer, index=False, sheet_name="Hasil")
            st.download_button("Unduh sebagai ODS (.ods)", ods_buffer.getvalue(),
                               "hasil_filter.ods", "application/vnd.oasis.opendocument.spreadsheet")

