from pathlib import Path
import json
import pandas as pd

bank_dir = Path('/Users/sunyitong/coding/voucher/代码资料/银行流水')
invoice_path = Path('/Users/sunyitong/coding/voucher/代码资料/4月901张发票.xlsx')

def preview_excel(path, max_sheets=10):
    info = {'path': str(path), 'exists': path.exists(), 'size': path.stat().st_size if path.exists() else None}
    try:
        book = pd.ExcelFile(path)
        info['sheets'] = book.sheet_names
        sheet_infos = []
        for sheet in book.sheet_names[:max_sheets]:
            raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=20, dtype=object)
            non_empty_rows = []
            for idx, row in raw.iterrows():
                values = [str(x).strip() for x in row.tolist() if pd.notna(x) and str(x).strip()]
                if values:
                    non_empty_rows.append({'row': int(idx + 1), 'values': values[:16]})
            sheet_infos.append({'sheet': sheet, 'preview_rows': non_empty_rows[:12]})
        info['sheet_infos'] = sheet_infos
    except Exception as e:
        info['error'] = repr(e)
    return info

all_info = {
    'bank_files': [preview_excel(p) for p in sorted(bank_dir.glob('*')) if p.is_file()],
    'invoice_file': preview_excel(invoice_path),
}

print(json.dumps(all_info, ensure_ascii=False, indent=2))
