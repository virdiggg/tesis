import pandas as pd
import numpy as np
import os
from util import formatting_excel

input_file = os.path.join('target', 'discriminant.xlsx')
output_file = os.path.join('result', 'discriminant_cleaned.xlsx')

name_mapping = {
    "P (X1)": "P",
    "WB (X2)": "WB",
    "BK (X3)": "BK",
    "D (Z)": "D",
    "PK (Y)": "PK"
}

try:
    df = pd.read_excel(input_file, index_col=0)

    valid_labels = [label for label in df.index if label in name_mapping]
    df_filtered = df.loc[valid_labels, valid_labels]

    df_filtered.index = [name_mapping[label] for label in df_filtered.index]
    df_filtered.columns = [name_mapping[label] for label in df_filtered.columns]

    mask = np.triu(np.ones(df_filtered.shape), k=1).astype(bool)
    df_final = df_filtered.where(~mask, "")

    df_final.to_excel(output_file, index_label="Konstruk")

    # print("\nPreview Tabel:")
    # print(df_final)
    formatting_excel(output_file)

except Exception as e:
    print(f"Terjadi kesalahan saat parsing: {e}")