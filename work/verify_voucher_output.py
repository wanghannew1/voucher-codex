from pathlib import Path
import json

import pandas as pd

path = Path('/Users/sunyitong/Documents/Codex/2026-06-06/new-chat/outputs/voucher_bank_analysis/银行存款1002凭证筛选与规律分析.xlsx')
book = pd.ExcelFile(path)
checks = {'sheets': book.sheet_names, 'exists': path.exists(), 'size': path.stat().st_size}
for sheet in book.sheet_names:
    df = pd.read_excel(path, sheet_name=sheet)
    checks[sheet] = {'rows': len(df), 'cols': len(df.columns), 'columns': [str(c) for c in df.columns[:8]]}

listing = pd.read_excel(path, sheet_name='1002凭证清单')
full = pd.read_excel(path, sheet_name='筛选后完整分录')
checks['direction_counts'] = listing['方向'].value_counts().to_dict()
checks['contains_non_1002_full_lines'] = bool((full['accsubjcode'].astype(str).str.replace(r'\.0$', '', regex=True) != '1002').any())
checks['all_listed_have_1002_amount'] = bool(((listing['银行借方金额'].fillna(0) != 0) | (listing['银行贷方金额'].fillna(0) != 0)).all())
print(json.dumps(checks, ensure_ascii=False, indent=2, default=str))
