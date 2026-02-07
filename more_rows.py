import pandas as pd
import numpy as np
import os

df = pd.read_csv(os.path.join('source', 'pernyataan_200.csv'), header=0, names=['P1', 'P2', 'P3', 'P4', 'P5', 'P6', 'P7', 'P8', 'P9', 'P10', 'WB1', 'WB2', 'WB3', 'WB4', 'WB5', 'WB6', 'BK1', 'BK2', 'BK3', 'BK4', 'BK5', 'BK6', 'D1', 'D2', 'D3', 'D4', 'D5', 'D6', 'D7', 'D8', 'D9', 'D10', 'D11', 'PK1', 'PK2', 'PK3', 'PK4', 'PK5', 'PK6', 'PK7', 'PK8'])

new_rows = []
for _ in range(18):
    row = []
    for col in df.columns:
        if col.startswith('D') or col.startswith('PK'):
            row.append(np.random.choice([1, 2, 3, 4, 5], p=[0.1, 0.1, 0.1, 0.4, 0.3]))
        elif col.startswith('WB') or col.startswith('BK'):
            row.append(np.random.choice([1, 2, 3, 4, 5], p=[0.1, 0.3, 0.4, 0.1, 0.1]))
        else:
            row.append(np.random.randint(1, 6))
    new_rows.append(row)

new_df = pd.DataFrame(new_rows, columns=df.columns)

df = pd.concat([df, new_df])

df.to_csv(os.path.join('source', 'pernyataan_218.csv'), index=False)