from pathlib import Path
import json
import pandas as pd

bank_path = Path('/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹吉林银行4月对账单.xlsx')
invoice_path = Path('/Users/sunyitong/coding/voucher/代码资料/4月901张发票.xlsx')

bank = pd.read_excel(bank_path, sheet_name='账户明细查询', header=8, dtype=object)
bank = bank.dropna(how='all')
invoice = pd.read_excel(invoice_path, sheet_name='Sheet1', dtype=object)
invoice = invoice.dropna(how='all')

def profile(df, cols):
    out = {}
    for col in cols:
        if col not in df.columns:
            continue
        s = df[col]
        non_null = s.notna() & (s.astype(str).str.strip() != '')
        samples = s[non_null].astype(str).drop_duplicates().head(5).tolist()
        out[col] = {
            'non_empty': int(non_null.sum()),
            'coverage': round(float(non_null.mean()), 4),
            'unique': int(s[non_null].astype(str).nunique()),
            'samples': samples,
        }
    return out

bank['交易金额_num'] = pd.to_numeric(bank.get('交易金额', pd.Series(dtype=object)).astype(str).str.replace(',', ''), errors='coerce')
invoice['价税合计_num'] = pd.to_numeric(invoice.get('价税合计', pd.Series(dtype=object)), errors='coerce')
invoice['金额_num'] = pd.to_numeric(invoice.get('金额', pd.Series(dtype=object)), errors='coerce')
invoice['税额_num'] = pd.to_numeric(invoice.get('税额', pd.Series(dtype=object)), errors='coerce')

result = {
    'jilin_bank_shape': bank.shape,
    'jilin_bank_columns': bank.columns.astype(str).tolist(),
    'jilin_bank_direction_counts': bank.get('借贷标志', pd.Series(dtype=object)).value_counts(dropna=False).to_dict(),
    'jilin_bank_amount_by_direction': bank.groupby('借贷标志', dropna=False)['交易金额_num'].sum().round(2).to_dict() if '借贷标志' in bank.columns else {},
    'jilin_bank_field_profile': profile(bank, ['交易时间', '交易对手户名', '交易对手账号', '交易对手行名', '借贷标志', '交易金额', '交易后余额', '币种', '交易类型', '用途', '摘要']),
    'invoice_shape': invoice.shape,
    'invoice_columns': invoice.columns.astype(str).tolist(),
    'invoice_status_counts': invoice.get('发票状态', pd.Series(dtype=object)).value_counts(dropna=False).to_dict(),
    'invoice_positive_counts': invoice.get('是否正数发票', pd.Series(dtype=object)).value_counts(dropna=False).to_dict(),
    'invoice_type_counts': invoice.get('发票票种', pd.Series(dtype=object)).value_counts(dropna=False).to_dict(),
    'invoice_amounts': {
        '金额合计': round(float(invoice['金额_num'].sum()), 2),
        '税额合计': round(float(invoice['税额_num'].sum()), 2),
        '价税合计': round(float(invoice['价税合计_num'].sum()), 2),
        '正数价税合计': round(float(invoice.loc[invoice.get('是否正数发票') == '是', '价税合计_num'].sum()), 2),
        '负数价税合计': round(float(invoice.loc[invoice.get('是否正数发票') == '否', '价税合计_num'].sum()), 2),
    },
    'invoice_field_profile': profile(invoice, ['数电发票号码', '销方识别号', '销方名称', '购方识别号', '购买方名称', '开票日期', '货物或应税劳务名称', '金额', '税率', '税额', '价税合计', '发票票种', '发票状态', '是否正数发票', '备注']),
}

print(json.dumps(result, ensure_ascii=False, indent=2, default=str))
