import pandas as pd
import numpy as np
import os

df = pd.read_csv(os.path.join('source', 'pernyataan.csv'), header=0, names=['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10', 'WB1', 'WB2', 'WB3', 'WB4', 'WB5', 'WB6', 'BK1', 'BK2', 'BK3', 'BK4', 'BK5', 'BK6', 'D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11', 'PK1', 'PK2', 'PK3', 'PK4', 'PK5', 'PK6', 'PK7', 'PK8'])

total = 58
new_rows = []

for _ in range(total):
    row = []
    for col in df.columns:
        if col.startswith('D') or col.startswith('PK'):
            row.append(np.random.choice([1, 2, 3, 4, 5], p=[0.05, 0.05, 0.1, 0.45, 0.35]))
        elif col.startswith('P') or col.startswith('WB') or col.startswith('BK'):
            row.append(np.random.choice([1, 2, 3, 4, 5], p=[0.05, 0.45, 0.35, 0.1, 0.05]))
        else:
            row.append(np.random.randint(1, 6))
    new_rows.append(row)

new_df = pd.DataFrame(new_rows, columns=df.columns)

output_path = os.path.join('target', f"pernyataan_{total}.csv")
new_df.to_csv(output_path, index=False)