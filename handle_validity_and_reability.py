import pandas as pd
from util import formatting_excel
import os

input_file = os.path.join('target', 'validity_and_reability.xlsx')
output_file = os.path.join('result', 'validity_and_reability_cleaned.xlsx')

try:
    df = pd.read_excel(input_file)

    column_label = df.columns[0]
    column_ave = 'Average Variance Extracted (AVE)'

    name_mapping = {
        "P (X1)": "Pelatihan",
        "WB (X2)": "Work-Life Balance",
        "BK (X3)": "Beban Kerja",
        "D (Z)": "Digitalisasi",
        "PK (Y)": "Produktivitas Karyawan"
    }

    df_filtered = df[df[column_label].isin(name_mapping.keys())].copy()

    df_filtered['Konstruk'] = df_filtered[column_label].map(name_mapping)

    result_table = df_filtered[['Konstruk', column_ave]].reset_index(drop=True)
    result_table.columns = ['Konstruk', 'Average Variance Extracted (AVE)']

    # print("Preview Tabel:")
    # print("-" * 55)
    # print(result_table.to_string(index=False))

    result_table.to_excel(output_file, index=False)
    print(f"File '{output_file}' berhasil disimpan.")

    formatting_excel(output_file)

except FileNotFoundError:
    print(f"Error: File '{input_file}' tidak ditemukan. Pastikan nama file sudah benar.")
except Exception as e:
    print(f"Terjadi kesalahan: {e}")