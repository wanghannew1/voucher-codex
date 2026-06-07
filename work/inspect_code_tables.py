from pathlib import Path
import json
import pandas as pd

base = Path('/Users/sunyitong/coding/voucher/代码资料/化简代码表')
files = sorted(base.glob('*.xlsx'))

result = []
for path in files:
    item = {'file': str(path), 'size': path.stat().st_size}
    try:
        book = pd.ExcelFile(path)
        item['sheets'] = book.sheet_names
        sheets = []
        for sheet in book.sheet_names:
            df = pd.read_excel(path, sheet_name=sheet, dtype=object)
            df = df.dropna(how='all')
            sheets.append({
                'sheet': sheet,
                'rows': int(len(df)),
                'cols': int(len(df.columns)),
                'columns': [str(c) for c in df.columns],
                'sample': df.head(5).astype(str).replace('nan', '').to_dict(orient='records'),
            })
        item['sheet_infos'] = sheets
    except Exception as e:
        item['error'] = repr(e)
    result.append(item)

print(json.dumps(result, ensure_ascii=False, indent=2))
