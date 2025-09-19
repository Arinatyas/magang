import streamlit as st
import pandas as pd
import altair as alt
from io import BytesIO

st.set_page_config(layout="wide")
st.title("📊 Aplikasi Gabung, Filter, dan Visualisasi Data Multi-Sheet")

# -----------------------------
# Helper functions
# -----------------------------
def detect_header_row(df_raw, max_search=10):
    if df_raw is None or len(df_raw) == 0:
        return 0
    n = min(len(df_raw), max_search)
    best_idx, best_score = 0, -1
    for i in range(n):
        row = df_raw.iloc[i]
        non_null = row.notna().sum()
        numeric_count = pd.to_numeric(row.dropna(), errors="coerce").notna().sum()
        string_count = non_null - numeric_count
        score = string_count * 2 + non_null * 0.1
        if score > best_score:
            best_score = score
            best_idx = i
    return int(best_idx)

def fill_header_from_prev(df_raw, h):
    if h < 0:
        h = 0
    if h >= len(df_raw):
        h = 0
    header_row = df_raw.iloc[h].copy().astype(object)
    if h > 0:
        prev_row = df_raw.iloc[h-1].copy().astype(object)
    else:
        prev_row = pd.Series([None]*len(header_row))
    filled = []
    for i, val in enumerate(header_row):
        if pd.isna(val) or str(val).strip() == "":
            prev_val = prev_row.iat[i] if i < len(prev_row) else None
            if pd.isna(prev_val) or str(prev_val).strip() == "":
                filled.append(f"Kolom_{i+1}")
            else:
                filled.append(str(prev_val))
        else:
            filled.append(str(val))
    return filled

def clean_column_values(series):
    non_na = series.notna().sum()
    if non_na == 0:
        return series.astype(str).fillna("data kosong")
    numeric = pd.to_numeric(series, errors="coerce")
    numeric_non_na = numeric.notna().sum()
    if numeric_non_na / max(non_na, 1) >= 0.6:
        return numeric.fillna(0)
    else:
        return series.astype(str).replace("nan", "data kosong").fillna("data kosong")

def make_unique(cols):
    seen, new = {}, []
    for c in cols:
        c = "" if c is None else str(c)
        if c in seen:
            seen[c] += 1
            new.append(f"{c}_{seen[c]}")
        else:
            seen[c] = 0
            new.append(c)
    return new

# -----------------------------
# Upload files
# -----------------------------
uploaded_files = st.file_uploader(
    "Unggah file (Excel/CSV/ODS) — bisa banyak sekaligus",
    type=["xlsx", "csv", "ods"],
    accept_multiple_files=True
)

sheets_info = []
if uploaded_files:
    for f_idx, up in enumerate(uploaded_files):
        fname = up.name
        try:
            if fname.lower().endswith(".csv"):
                df_raw = pd.read_csv(up, header=None)
                detected = detect_header_row(df_raw)
                sheets_info.append({
                    "file_name": fname,
                    "sheet_name": fname,
                    "df_raw": df_raw,
                    "detected": detected
                })
            else:
                engine = "odf" if fname.lower().endswith(".ods") else None
                xls = pd.read_excel(up, sheet_name=None, header=None, engine=engine)
                for sname, df_raw in xls.items():
                    detected = detect_header_row(df_raw)
                    sheets_info.append({
                        "file_name": fname,
                        "sheet_name": sname,
                        "df_raw": df_raw,
                        "detected": detected
                    })
        except Exception as e:
            st.warning(f"Gagal membaca {fname}: {e}")

# -----------------------------
# Pilih header per sheet
# -----------------------------
chosen_headers = {}
if sheets_info:
    st.subheader("1️⃣ Pilih baris header per sheet (auto + manual)")
    for idx, info in enumerate(sheets_info):
        with st.expander(f"{info['file_name']} — {info['sheet_name']}"):
            df_raw = info["df_raw"]
            detected = info["detected"]
            start = max(0, detected - 3)
            end = min(len(df_raw), detected + 4)
            st.write(f"Preview baris {start}..{end-1}")
            st.dataframe(df_raw.iloc[start:end].reset_index(drop=True))
            chosen = st.number_input(
                f"Pilih baris header (deteksi otomatis: {detected})",
                min_value=0,
                max_value=len(df_raw)-1,
                value=detected,
                step=1,
                key=f"header_choice_{idx}"
            )
            chosen_headers[idx] = int(chosen)

combine = st.button("✅ Gabungkan Semua Sheet")

# -----------------------------
# Proses gabungan
# -----------------------------
data_gabungan = None
if sheets_info and combine:
    dfs = []
    for idx, info in enumerate(sheets_info):
        df_raw = info["df_raw"]
        file_name = info["file_name"]
        sheet_name = info["sheet_name"]
        h = chosen_headers.get(idx, info["detected"])
        headers = fill_header_from_prev(df_raw, h)
        df_data = df_raw.iloc[h+1:].copy()
        if df_data.shape[1] < len(headers):
            for c in range(df_data.shape[1], len(headers)):
                df_data[c] = pd.NA
        if df_data.shape[1] > len(headers):
            for i in range(df_data.shape[1] - len(headers)):
                headers.append(f"Kolom_extra_{i+1}")
        df_data.columns = headers[:df_data.shape[1]]
        df_data["Sumber_File"] = file_name
        df_data["Sumber_Sheet"] = sheet_name
        for col in df_data.columns:
            if col not in ["Sumber_File", "Sumber_Sheet"]:
                df_data[col] = clean_column_values(df_data[col])
        dfs.append(df_data)
    data_gabungan = pd.concat(dfs, ignore_index=True)
    data_gabungan.columns = make_unique(list(data_gabungan.columns))
    st.success("✅ Semua sheet berhasil digabung!")

# -----------------------------
# Data Gabungan
# -----------------------------
st.subheader("2️⃣ Data Gabungan")
if data_gabungan is not None:
    st.dataframe(data_gabungan.head(200), use_container_width=True)

    # -----------------------------
    # Filter
    # -----------------------------
    st.subheader("3️⃣ Filter Data")
    filter_cols = st.multiselect("Pilih kolom untuk difilter", list(data_gabungan.columns))
    filtered = data_gabungan.copy()
    for col in filter_cols:
        options = filtered[col].dropna().unique().tolist()
        selected = st.multiselect(f"Pilih nilai untuk `{col}`", options, key=f"val_{col}")
        if selected:
            filtered = filtered[filtered[col].astype(str).isin([str(s) for s in selected])]
    if filter_cols:
        st.dataframe(filtered[filter_cols].reset_index(drop=True), use_container_width=True)

        # Unduh hasil
        csv_bytes = filtered.to_csv(index=False).encode("utf-8")
        st.download_button("⬇️ Unduh CSV", csv_bytes, "hasil.csv", "text/csv")

        buf = BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            filtered.to_excel(writer, index=False, sheet_name="Hasil")
        st.download_button("⬇️ Unduh Excel", buf.getvalue(),
                           "hasil.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # -----------------------------
    # Visualisasi
    # -----------------------------
    st.subheader("4️⃣ Visualisasi Data")
    if filter_cols:
        x_col = st.selectbox("Sumbu X", filter_cols)
        y_col = st.selectbox("Sumbu Y", filter_cols)
        chart_type = st.selectbox("Jenis Grafik", ["Diagram Batang", "Diagram Garis", "Diagram Sebar"])

        try:
            filtered[y_col] = pd.to_numeric(filtered[y_col], errors="coerce")
        except:
            pass

        def col_type(s):
            if pd.api.types.is_numeric_dtype(s): return "quantitative"
            if pd.api.types.is_datetime64_any_dtype(s): return "temporal"
            return "nominal"

        if pd.api.types.is_numeric_dtype(filtered[y_col]):
            if chart_type == "Diagram Batang":
                chart = alt.Chart(filtered).mark_bar().encode(
                    x=alt.X(x_col, type=col_type(filtered[x_col])),
                    y=alt.Y(y_col, type=col_type(filtered[y_col])),
                    tooltip=list(filtered.columns))
            elif chart_type == "Diagram Garis":
                chart = alt.Chart(filtered).mark_line(point=True).encode(
                    x=alt.X(x_col, type=col_type(filtered[x_col])),
                    y=alt.Y(y_col, type=col_type(filtered[y_col])),
                    tooltip=list(filtered.columns))
            else:
                chart = alt.Chart(filtered).mark_circle(size=60).encode(
                    x=alt.X(x_col, type=col_type(filtered[x_col])),
                    y=alt.Y(y_col, type=col_type(filtered[y_col])),
                    tooltip=list(filtered.columns))
        else:
            chart = alt.Chart(filtered).mark_bar().encode(
                x=alt.X(x_col, type=col_type(filtered[x_col])),
                y=alt.Y("count()", type="quantitative"),
                color=alt.Color(y_col, type="nominal"),
                tooltip=list(filtered.columns))

        st.altair_chart(chart, use_container_width=True)
