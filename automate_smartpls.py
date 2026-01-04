import pandas as pd
import numpy as np
import os, re, gc, math
from util import formatting_excel, preview_table

# Bagian ini untuk konfigurasi, disesuaikan dengan kebutuhan
# =============================================================================
# File input
input_file = os.path.join('target', 'smartpls.xlsx')
output_file = os.path.join('result', 'cleaned_smartpls.xlsx')

# Mapping variabel
full_mapping = {
    "P (X1)": "Pelatihan",
    "WB (X2)": "Work-Life Balance",
    "BK (X3)": "Beban Kerja",
    "D (Z)": "Digitalisasi",
    "PK (Y)": "Produktivitas Karyawan"
}

# Mapping variabel
short_mapping = {
    "P (X1)": "P",
    "WB (X2)": "WB",
    "BK (X3)": "BK",
    "D (Z)": "D",
    "PK (Y)": "PK"
}

# Ketentuan level signifikan, sesuaikan one tail atau two tail
SIGNIFICANCE_LEVEL = 0.025 # One tail
# SIGNIFICANCE_LEVEL = 0.05 # Two tail
# =============================================================================

# =============================================================================
# Mulai dari sini untuk proses, jangan diubah
output_flc = os.path.join('result', 'cleaned_flc.xlsx')
output_htmt = os.path.join('result', 'cleaned_htmt.xlsx')
output_val = os.path.join('result', 'cleaned_validity.xlsx')
output_rel = os.path.join('result', 'cleaned_reliability.xlsx')
output_loading = os.path.join('result', 'cleaned_loading_factor.xlsx')
output_boot = os.path.join('result', 'cleaned_bootstrapping.xlsx')
output_r_square = os.path.join("result", "cleaned_r_square.xlsx")
output_blindfold = os.path.join("result", "cleaned_blindfold.xlsx")

def process_flc(df_raw):
    """Proses Fornell-Larcker: Diagonal tetap ada (k=1), nama inisial."""
    valid_labels = [label for label in df_raw.index if label in short_mapping]
    df = df_raw.loc[valid_labels, valid_labels].copy()
    df.index = [short_mapping[l] for l in df.index]
    df.columns = [short_mapping[l] for l in df.columns]
    mask = np.triu(np.ones(df.shape), k=1).astype(bool)
    return df.where(~mask, "")

def process_htmt(df_raw):
    """Proses HTMT: Diagonal dihapus (k=0), baris nama lengkap, kolom inisial."""
    valid_labels = [label for label in df_raw.index if label in full_mapping]
    df = df_raw.loc[valid_labels, valid_labels].copy()
    mask = np.triu(np.ones(df.shape), k=0).astype(bool)
    df_final = df.where(~mask, "")
    df_final.index = [full_mapping[l] for l in df_final.index]
    df_final.columns = [short_mapping[l] for l in df_final.columns]
    return df_final

def process_validity(df_raw):
    """Proses AVE: Filter variabel utama dan ambil kolom AVE saja."""
    column_label = df_raw.columns[0]
    column_ave = 'Average Variance Extracted (AVE)'
    df_filtered = df_raw[df_raw[column_label].isin(full_mapping.keys())].copy()
    df_filtered['Konstruk'] = df_filtered[column_label].map(full_mapping)
    result = df_filtered[['Konstruk', column_ave]].reset_index(drop=True)

    ave_values = result[column_ave].dropna()

    ave_sum = round(ave_values.sum(), 3)
    ave_count = int(ave_values.count())
    ave_avg = round(ave_sum / ave_count, 3) if ave_count > 0 else ""

    summary_rows = pd.DataFrame([
        {"Konstruk": "SUM", column_ave: ave_sum},
        {"Konstruk": "COUNT", column_ave: ave_count},
        {"Konstruk": "AVERAGE (SUM/COUNT)", column_ave: ave_avg},
    ])

    result = pd.concat([result, summary_rows], ignore_index=True)
    return result

def process_reliability(df_raw):
    """
    Proses Reliability:
    - Ambil konstruk utama saja (tanpa jalur mediasi)
    - Ambil Cronbach's Alpha, rho_A, dan Composite Reliability
    - Ganti nama konstruk menjadi nama lengkap
    """
    konstruk_col = df_raw.columns[0]

    df_filtered = df_raw[
        df_raw[konstruk_col].isin(full_mapping.keys())
    ].copy()

    df_filtered['Konstruk'] = df_filtered[konstruk_col].map(full_mapping)

    result = df_filtered[[
        'Konstruk',
        "Cronbach's Alpha",
        'rho_A',
        'Composite Reliability'
    ]].reset_index(drop=True)

    return result

def process_loading_factor(df_raw):
    """Proses Loading Factor: Menampilkan label indikator di kolom pertama."""
    indikator_col = df_raw.columns[0]

    def is_valid_indikator(val):
        if pd.isna(val): return False
        val = str(val)
        if "*" in val: return False
        return bool(re.match(r"^[A-Za-z]+[0-9]+$", val))

    df_valid = df_raw[df_raw[indikator_col].apply(is_valid_indikator)].copy()

    def get_variabel_from_text(text):
        match = re.match(r"[A-Za-z]+", str(text))
        return match.group(0) if match else None

    col_variabel_map = {
        col: get_variabel_from_text(col)
        for col in df_valid.columns[1:]
    }

    unique_variabel = sorted(set(v for v in col_variabel_map.values() if v))

    result = pd.DataFrame(index=df_valid[indikator_col].values, columns=unique_variabel)

    for idx, row in df_valid.iterrows():
        ind_code = row[indikator_col]
        var_prefix = get_variabel_from_text(ind_code)

        for col, var_col in col_variabel_map.items():
            if var_col == var_prefix:
                value = row[col]
                if pd.notna(value):
                    result.loc[ind_code, var_prefix] = value
                break

    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() 
                for text in re.split('([0-9]+)', str(s))]

    result = result.reindex(sorted(result.index, key=natural_sort_key))
    return result

def process_bootstrapping(df_raw):
    """
    Proses Bootstrapping SmartPLS:
    - Jalur langsung & mediasi
    - Tahan terhadap format label SmartPLS
    - SIGNIFICANCE_LEVEL global
    """

    path_col = df_raw.columns[0]
    rows = []

    def normalize_label(label):
        """
        Mengubah:
        - 'VAL1 (X3)' -> 'VAL1'
        - 'VAL1'      -> 'VAL1'
        - 'VAL2 (Y)'  -> 'VAL2'
        """
        label = str(label).strip()
        return label.split("(")[0].strip()

    normalized_short_mapping = {
        normalize_label(k): v
        for k, v in short_mapping.items()
    }

    for _, row in df_raw.iterrows():
        raw_path = str(row[path_col]).strip()

        if "->" not in raw_path:
            continue

        left, right = raw_path.split("->")
        right = normalize_label(right)

        dst = normalized_short_mapping.get(right)
        if not dst:
            continue

        if ">" in left:
            parts = [normalize_label(p) for p in left.split(">")]

            if len(parts) < 2:
                continue

            src_codes = []
            for p in parts[:-1]:
                code = normalized_short_mapping.get(p)
                if code:
                    src_codes.append(code)

            if not src_codes:
                continue

            src = "".join(src_codes)
            jalur = f"{src} → {dst}"

        else:
            src = normalized_short_mapping.get(normalize_label(left))
            if not src:
                continue

            jalur = f"{src} → {dst}"

        pval = float(row['P Values'])

        rows.append({
            "Jalur": jalur,
            "Original Sample (O)": row['Original Sample (O)'],
            "STDEV": row['Standard Deviation (STDEV)'],
            "T-Statistic": row['T Statistics (|O/STDEV|)'],
            "P-Value": pval,
            "Keterangan": "Signifikan" if pval < SIGNIFICANCE_LEVEL else "Tidak signifikan"
        })

    return pd.DataFrame(rows)

def process_r_square(df_raw):
    """
    Proses R-Square SmartPLS:
    - Ambil variabel endogen (Y)
    - Gunakan nama dari full_mapping
    """

    df = df_raw.copy()

    konstruk_col = df.columns[0]

    endogen_keys = [
        k for k in full_mapping.keys()
        if "(Y)" in k
    ]

    df_filtered = df[
        df[konstruk_col].isin(endogen_keys)
    ].copy()

    df_filtered["Variabel"] = df_filtered[konstruk_col].map(
        lambda x: f"{full_mapping[x]} ({x.split()[0]})"
    )

    result = df_filtered[
        ["Variabel", "R Square", "R Square Adjusted"]
    ].reset_index(drop=True)

    return result

def process_blindfold(df_raw):
    """
    Proses Blindfolding SmartPLS:
    - Ambil konstruk utama saja
    - Buang jalur mediasi
    - Skala SSO & SSE (÷1000)
    """

    df = df_raw.copy()

    konstruk_col = df.columns[0]

    df_filtered = df[
        df[konstruk_col].isin(full_mapping.keys())
    ].copy()

    df_filtered["Konstruk"] = df_filtered[konstruk_col].map(full_mapping)

    def scale(val):
        if pd.isna(val):
            return ""
        return round(val / 1000, 3)

    df_filtered["SSO"] = df_filtered["SSO"].apply(scale)
    df_filtered["SSE"] = df_filtered["SSE"].apply(scale)

    if "Q² (=1-SSE/SSO)" in df_filtered.columns:
        df_filtered["Q² (=1-SSE/SSO)"] = df_filtered[
            "Q² (=1-SSE/SSO)"
        ].apply(lambda x: round(x, 3) if pd.notna(x) else "")

    result = df_filtered[
        ["Konstruk", "SSO", "SSE", "Q² (=1-SSE/SSO)"]
    ].reset_index(drop=True)

    return result

def process_gof(df_validity, df_r_square):
    """
    Proses Goodness of Fit (GoF):
    - Average AVE dari sheet validity
    - Average R-Square dari sheet r square
    """

    ave_row = df_validity[
        df_validity["Konstruk"] == "AVERAGE (SUM/COUNT)"
    ]

    if ave_row.empty:
        raise ValueError("Row AVERAGE (SUM/COUNT) tidak ditemukan di sheet validity")

    avg_ave = float(
        ave_row["Average Variance Extracted (AVE)"].values[0]
    )

    avg_r_square = round(
        df_r_square["R Square"].mean(), 3
    )

    gof_value = round(
        math.sqrt(avg_ave * avg_r_square), 3
    )

    result = pd.DataFrame({
        "Komponen": [
            "Average AVE",
            "Average R-Square",
            "Goodness of Fit (GoF)"
        ],
        "Nilai": [
            avg_ave,
            avg_r_square,
            gof_value
        ]
    })

    return result

try:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        df_flc_raw = pd.read_excel(input_file, sheet_name='flc', index_col=0)
        df_flc_final = process_flc(df_flc_raw)
        df_flc_final.to_excel(writer, sheet_name='flc', index_label="Konstruk")

        df_htmt_raw = pd.read_excel(input_file, sheet_name='htmt', index_col=0)
        df_htmt_final = process_htmt(df_htmt_raw)
        df_htmt_final.to_excel(writer, sheet_name='htmt', index_label="Konstruk")

        df_val_raw = pd.read_excel(input_file, sheet_name='validity and reability')
        df_val_final = process_validity(df_val_raw)
        df_val_final.to_excel(writer, sheet_name='validity', index=False)

        df_rel_final = process_reliability(df_val_raw)
        df_rel_final.to_excel(writer, sheet_name='reliability', index=False)

        df_load_raw = pd.read_excel(input_file, sheet_name='loading factor')
        df_load_final = process_loading_factor(df_load_raw)
        df_load_final.to_excel(writer, sheet_name='loading factor', index_label="Indikator")

        df_boot_raw = pd.read_excel(input_file, sheet_name='bootstrapping')
        df_boot_final = process_bootstrapping(df_boot_raw)
        df_boot_final.to_excel(writer, sheet_name='bootstrapping', index=False)

        df_r_square_raw = pd.read_excel(input_file, sheet_name='r square')
        df_r_square_final = process_r_square(df_r_square_raw)
        df_r_square_final.to_excel(writer, sheet_name='r square', index=False)

        df_blind_raw = pd.read_excel(input_file, sheet_name='blindfold')
        df_blind_final = process_blindfold(df_blind_raw)
        df_blind_final.to_excel(writer, sheet_name='blindfold', index=False)

        df_gof = process_gof(df_val_final, df_r_square_final)
        df_gof.to_excel(writer, sheet_name='gof', index=False)

    formatting_excel(output_file)

except Exception as e:
    print("Terjadi kesalahan:", e)