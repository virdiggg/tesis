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
# SIGNIFICANCE_LEVEL = 0.025 # Two tail
SIGNIFICANCE_LEVEL = 0.05 # One tail
# =============================================================================

# =============================================================================
# Mulai dari sini untuk proses, jangan diubah
# def process_flc(df_raw):
#     """Proses Fornell-Larcker: Diagonal tetap ada (k=1), nama inisial."""
#     valid_labels = [label for label in df_raw.index if label in short_mapping]
#     df = df_raw.loc[valid_labels, valid_labels].copy()
#     df.index = [short_mapping[l] for l in df.index]
#     df.columns = [short_mapping[l] for l in df.columns]
#     mask = np.triu(np.ones(df.shape), k=1).astype(bool)
#     return df.where(~mask, "")
def process_flc(df_raw):
    """
    Proses Fornell-Larcker: Mengembalikan 2 output.
    1. df_short: Hanya variabel utama (BK, D, P, dsb).
    2. df_long: Variabel utama + Moderasi (BK * D, dsb).
    """

    def transform_label(label, mode='short'):
        if label in short_mapping:
            return short_mapping[label]

        if ">" in label and mode == 'long':
            parts = [p.split("(")[0].strip() for p in label.split(">")]
            final_parts = []
            for p in parts[:2]:
                found = False
                for k, v in short_mapping.items():
                    if k.startswith(p):
                        final_parts.append(v)
                        found = True
                        break
                if not found: final_parts.append(p)
            return f"{final_parts[0]} * {final_parts[1]}"

        return None

    labels_short = [l for l in df_raw.index if l in short_mapping]
    df_short = df_raw.loc[labels_short, labels_short].copy()
    df_short.index = [transform_label(l, 'short') for l in df_short.index]
    df_short.columns = [transform_label(l, 'short') for l in df_short.columns]
    mask_s = np.triu(np.ones(df_short.shape), k=1).astype(bool)
    df_short = df_short.where(~mask_s, "")

    labels_long = [l for l in df_raw.index if l in short_mapping or ">" in str(l)]
    df_long = df_raw.loc[labels_long, labels_long].copy()

    new_labels_long = [transform_label(l, 'long') for l in df_long.index]
    df_long.index = new_labels_long
    df_long.columns = new_labels_long
    mask_l = np.triu(np.ones(df_long.shape), k=1).astype(bool)
    df_long = df_long.where(~mask_l, "")

    return df_short, df_long

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

            src = " → ".join(src_codes)
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

def process_hypotheses(df_boot_raw):
    """
    Menghasilkan tabel hipotesis secara dinamis berdasarkan full_mapping.
    Format: H1-H4 (Langsung), H5-H7 (Moderasi/Mediasi)
    """
    path_col = df_boot_raw.columns[0]
    results = []

    var_x = [k for k in full_mapping.keys() if "(X" in k]
    var_z = [k for k in full_mapping.keys() if "(Z)" in k]
    var_y = [k for k in full_mapping.keys() if "(Y)" in k]

    if not var_z or not var_y:
        return pd.DataFrame()

    z_key = var_z[0]
    y_key = var_y[0]
    z_short = short_mapping[z_key]

    dynamic_hypo = []

    for i, x_key in enumerate(var_x, 1):
        h_code = f"H{i}"
        label = f"{full_mapping[x_key]} berpengaruh terhadap {full_mapping[y_key]}"
        dynamic_hypo.append((h_code, label, x_key))

    next_idx = len(var_x) + 1
    dynamic_hypo.append((f"H{next_idx}", f"{full_mapping[z_key]} berpengaruh terhadap {full_mapping[y_key]}", z_key))

    start_mod = next_idx + 1
    for i, x_key in enumerate(var_x, start_mod):
        h_code = f"H{i}"
        x_short = short_mapping[x_key]
        label = f"{full_mapping[z_key]} memoderasi pengaruh {full_mapping[x_key]} terhadap {full_mapping[y_key]}"
        search_term = f"{x_short} > {z_short}"
        dynamic_hypo.append((h_code, label, search_term))

    for code, label_display, search_term in dynamic_hypo:
        row_data = df_boot_raw[df_boot_raw[path_col].str.contains(search_term, regex=False, na=False)]

        if not row_data.empty:
            row = row_data.iloc[0]
            p_val = float(row['P Values'])

            is_sig = p_val < SIGNIFICANCE_LEVEL
            keterangan = f"P-Value {'<' if is_sig else '>'} {SIGNIFICANCE_LEVEL}"

            if "memoderasi" in label_display:
                kesimpulan = "Moderasi Berpengaruh" if is_sig else "Moderasi tidak berpengaruh"
            else:
                kesimpulan = "Hipotesis Diterima" if is_sig else "Hipotesis Ditolak"

            results.append({
                "No": code,
                "Pengembangan Hipotesis": label_display,
                "P-value": round(p_val, 3),
                "Keterangan": keterangan,
                "Kesimpulan": kesimpulan
            })

    return pd.DataFrame(results)

def process_r_square(df_raw):
    """
    Proses R-Square SmartPLS:
    - Ambil variabel endogen (Y)
    - Tambahkan kolom Kontribusi (%) dan Faktor Lain (%)
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

    result["Kontribusi (%)"] = (result["R Square"] * 100).round(1).astype(str) + "%"

    result["Faktor Lain (%)"] = ((1 - result["R Square"]) * 100).round(1).astype(str) + "%"

    # untuk keperluan chart/grafik di Excel
    # result["Kontribusi (%)"] = (result["R Square"] * 100).round(1)
    # result["Faktor Lain (%)"] = (100 - result["Kontribusi (%)"]).round(1)

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

def process_vif(df_raw):
    """
    Proses VIF SmartPLS (FINAL):
    - Konstruk utama  → MAX VIF indikator
    - Interaksi       → Ambil langsung
    """

    label_col = df_raw.columns[0]
    vif_col = df_raw.columns[1]

    indikator_vif = {}
    interaction_rows = []

    prefix_to_konstruk = {
        short_mapping[k]: full_mapping[k]
        for k in short_mapping
    }

    for _, row in df_raw.iterrows():
        label = str(row[label_col]).strip()
        vif = row[vif_col]

        if pd.isna(vif):
            continue

        if "*" in label:
            left, right = [p.strip() for p in label.split("*")]

            if left in full_mapping and right in full_mapping:
                interaction_rows.append({
                    "Konstruk": f"{full_mapping[left]} × {full_mapping[right]}",
                    "VIF": round(float(vif), 3)
                })

        elif re.match(r"^[A-Za-z]+[0-9]+$", label):
            prefix = re.match(r"[A-Za-z]+", label).group()

            konstruk = prefix_to_konstruk.get(prefix)
            if konstruk:
                indikator_vif.setdefault(konstruk, []).append(float(vif))

    konstruk_rows = [
        {
            "Konstruk": konstruk,
            "VIF": round(max(vifs), 3)
        }
        for konstruk, vifs in indikator_vif.items()
    ]

    result = pd.DataFrame(konstruk_rows + interaction_rows)

    return result

try:
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:

        df_flc_raw = pd.read_excel(input_file, sheet_name='flc', index_col=0)
        # df_flc_final = process_flc(df_flc_raw)
        df_flc_final, df_flc_long = process_flc(df_flc_raw)
        df_flc_final.to_excel(writer, sheet_name='flc', index_label="Konstruk")
        df_flc_long.to_excel(writer, sheet_name='flc_long', index_label="Konstruk")

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

        df_mod_final = process_hypotheses(df_boot_raw)
        df_mod_final.to_excel(writer, sheet_name='hypotheses', index=False)

        df_r_square_raw = pd.read_excel(input_file, sheet_name='r square')
        df_r_square_final = process_r_square(df_r_square_raw)
        df_r_square_final.to_excel(writer, sheet_name='r square', index=False)

        df_blind_raw = pd.read_excel(input_file, sheet_name='blindfold')
        df_blind_final = process_blindfold(df_blind_raw)
        df_blind_final.to_excel(writer, sheet_name='blindfold', index=False)

        df_gof = process_gof(df_val_final, df_r_square_final)
        df_gof.to_excel(writer, sheet_name='gof', index=False)

        df_vif_raw = pd.read_excel(input_file, sheet_name='vif', header=None)
        df_vif_final = process_vif(df_vif_raw)
        df_vif_final.to_excel(writer, sheet_name='vif', index=False)

        df_nfi = pd.read_excel(input_file, sheet_name='nfi', index_col=0)
        df_nfi.to_excel(writer, sheet_name='nfi', index_label="")

        df_penelitian = pd.read_excel(input_file, sheet_name='penelitian terdahulu', index_col=0)
        df_penelitian.to_excel(writer, sheet_name='penelitian terdahulu', index=False)

    gc.collect()
    formatting_excel(output_file)

except Exception as e:
    print("Terjadi kesalahan:", e)