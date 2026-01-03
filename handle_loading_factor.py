import pandas as pd
from util import formatting_excel
import re, os

input_file = os.path.join('target', 'loading_factor.xlsx')
output_file = os.path.join('result', 'loading_factor_cleaned.xlsx')
sheet_name = 0

try:
    df = pd.read_excel(input_file, sheet_name=sheet_name)

    indikator_col = df.columns[0]

    def is_valid_indikator(val):
        if pd.isna(val):
            return False
        val = str(val)
        if "*" in val:
            return False
        return bool(re.match(r"^[A-Za-z]+[0-9]+$", val))

    df_valid = df[df[indikator_col].apply(is_valid_indikator)].copy()

    indikator = df_valid[indikator_col].astype(str)

    def get_variabel_from_indikator(kode):
        return re.match(r"[A-Za-z]+", kode).group(0)

    def get_variabel_from_column(col):
        match = re.match(r"[A-Za-z]+", str(col))
        return match.group(0) if match else None

    col_variabel_map = {
        col: get_variabel_from_column(col)
        for col in df_valid.columns[1:]
    }

    unique_variabel = sorted(set(col_variabel_map.values()))

    result = pd.DataFrame(index=indikator, columns=unique_variabel)

    for idx, ind in df_valid.iterrows():
        var_ind = get_variabel_from_indikator(ind[indikator_col])

        for col, var_col in col_variabel_map.items():
            if var_col == var_ind:
                value = ind[col]
                if pd.notna(value):
                    result.loc[ind[indikator_col], var_ind] = value
                break

    result.index.name = None

    result.to_excel(output_file)

    formatting_excel(output_file)

except FileNotFoundError:
    print(f"Error: File '{input_file}' tidak ditemukan. Pastikan nama file sudah benar.")
except Exception as e:
    print(f"Terjadi kesalahan: {e}")