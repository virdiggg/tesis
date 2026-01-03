import pandas as pd
import numpy as np
import os, re, gc
from util import formatting_excel

input_file = os.path.join('target', 'smartpls.xlsx')
output_flc = os.path.join('result', 'flc_cleaned.xlsx')
output_htmt = os.path.join('result', 'htmt_cleaned.xlsx')
output_val_rel = os.path.join('result', 'validity_and_reability_cleaned.xlsx')
output_loading = os.path.join('result', 'loading_factor_cleaned.xlsx')

full_mapping = {
    "P (X1)": "Pelatihan",
    "WB (X2)": "Work-Life Balance",
    "BK (X3)": "Beban Kerja",
    "D (Z)": "Digitalisasi",
    "PK (Y)": "Produktivitas Karyawan"
}

short_mapping = {
    "P (X1)": "P",
    "WB (X2)": "WB",
    "BK (X3)": "BK",
    "D (Z)": "D",
    "PK (Y)": "PK"
}

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
    return result

def process_loading_factor(df_raw):
    """Proses Loading Factor: Ekstraksi indikator dan pembersihan matriks."""
    indikator_col = df_raw.columns[0]

    def is_valid_indikator(val):
        if pd.isna(val): return False
        val = str(val)
        if "*" in val: return False
        return bool(re.match(r"^[A-Za-z]+[0-9]+$", val))

    df_valid = df_raw[df_raw[indikator_col].apply(is_valid_indikator)].copy()

    def get_var_prefix(text):
        match = re.match(r"[A-Za-z]+", str(text))
        return match.group(0) if match else None

    col_variabel_map = {col: get_var_prefix(col) for col in df_valid.columns[1:]}
    unique_variabel = sorted(set(v for v in col_variabel_map.values() if v))

    result = pd.DataFrame(index=df_valid[indikator_col].astype(str), columns=unique_variabel)

    for _, row in df_valid.iterrows():
        ind_code = row[indikator_col]
        var_prefix = get_var_prefix(ind_code)
        for col, col_prefix in col_variabel_map.items():
            if col_prefix == var_prefix and pd.notna(row[col]):
                result.loc[ind_code, col_prefix] = row[col]
                break

    return result

try:
    df_flc_raw = pd.read_excel(input_file, sheet_name='flc', index_col=0)
    df_flc_final = process_flc(df_flc_raw)
    df_flc_final.to_excel(output_flc, index_label="Konstruk")
    formatting_excel(output_flc)

    gc.collect()

    df_htmt_raw = pd.read_excel(input_file, sheet_name='htmt', index_col=0)
    df_htmt_final = process_htmt(df_htmt_raw)
    df_htmt_final.to_excel(output_htmt, index_label="Konstruk")
    formatting_excel(output_htmt)

    gc.collect()

    df_val_raw = pd.read_excel(input_file, sheet_name='validity and reability')
    df_val_final = process_validity(df_val_raw)
    df_val_final.to_excel(output_val_rel, index=False)
    formatting_excel(output_val_rel)

    gc.collect()

    df_load_raw = pd.read_excel(input_file, sheet_name='loading factor')
    df_load_final = process_loading_factor(df_load_raw)
    df_load_final.to_excel(output_loading, index=False)
    formatting_excel(output_loading)

    gc.collect()

except Exception as e:
    print(f"Terjadi kesalahan: {e}")