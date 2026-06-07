import json
from pathlib import Path

import pandas as pd

path = Path('/Users/sunyitong/coding/voucher/代码资料/凭证信息_temp - 2026-06-02T140354.252.xlsx')
book = pd.ExcelFile(path)
print(json.dumps({'sheet_names': book.sheet_names}, ensure_ascii=False, indent=2))
for sheet in book.sheet_names:
    df = pd.read_excel(path, sheet_name=sheet, nrows=12)
    print(f'\nSHEET: {sheet}')
    print('columns:', json.dumps([str(c) for c in df.columns], ensure_ascii=False))
    print(df.head(8).to_string(index=False))
