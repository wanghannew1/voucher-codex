from __future__ import annotations

import csv
import cgi
import io
import json
import math
import re
import sys
from dataclasses import dataclass, asdict
from datetime import date, datetime, timedelta
from decimal import Decimal, ROUND_HALF_UP
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, urlparse

import pandas as pd

ROOT = Path(__file__).resolve().parent
PYDEPS = ROOT.parent / 'pydeps'
if PYDEPS.exists():
    sys.path.insert(0, str(PYDEPS))

BANK_PATH = Path('/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹吉林银行4月对账单.xlsx')
INVOICE_PATH = Path('/Users/sunyitong/coding/voucher/代码资料/4月901张发票.xlsx')
CODE_DIR = Path('/Users/sunyitong/coding/voucher/代码资料/化简代码表')


def money(value) -> float:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return 0.0
    text = str(value).replace(',', '').replace('¥', '').strip()
    if not text:
        return 0.0
    try:
        return float(Decimal(text).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))
    except Exception:
        return 0.0


def round2(value: float) -> float:
    return float(Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def norm_name(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)) or str(value).lower() == 'nan':
        return ''
    text = str(value or '').strip().lower()
    for token in ['有限公司', '有限责任公司', '股份有限公司', '集团', '分公司', '公司', '（', '）', '(', ')', ' ', '\u3000']:
        text = text.replace(token, '')
    return text


def parse_memo_amount(label: str, text: str) -> float | None:
    if not text:
        return None
    pattern = rf'{label}\s*[:：]\s*(-?\d+(?:\.\d+)?)'
    match = re.search(pattern, text)
    return money(match.group(1)) if match else None


def clean_text(value) -> str:
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return ''
    text = str(value).strip()
    return '' if text.lower() == 'nan' else text


@dataclass
class Entry:
    summary: str
    subject_code: str
    subject_name: str
    debit: float = 0.0
    credit: float = 0.0
    customer_code: str = ''
    customer_name: str = ''
    supplier_code: str = ''
    supplier_name: str = ''
    department_code: str = ''
    project_code: str = ''
    bank_code: str = ''
    cashflow_code: str = ''


@dataclass
class Voucher:
    id: str
    date: str
    number: int | None
    source: str
    status: str
    confidence: str
    bank_ref: str
    bank_time: str
    bank_account: str
    counterparty: str
    amount: float
    direction: str
    description: str
    invoice_refs: list[str]
    entries: list[Entry]


class DemoData:
    def __init__(self):
        self.import_results = []
        self.bank = self._load_initial_bank()
        self.invoices = self._load_invoices()
        self.subjects = self._load_subjects()
        self.customers = self._load_customers()
        self.suppliers = self._load_suppliers()
        self.bank_code = '2801:银行档案'

    def _load_initial_bank(self) -> pd.DataFrame:
        df = self._parse_bank_file(BANK_PATH, source_file=BANK_PATH.name)
        self.import_results = [{
            'file': BANK_PATH.name,
            'status': '已预加载',
            'bank': '吉林银行',
            'account': '8936374879000001',
            'rows': int(len(df)),
            'income': round2(df.loc[df['借贷标志'] == '贷', '交易金额_num'].sum()),
            'expense': round2(df.loc[df['借贷标志'] == '借', '交易金额_num'].sum()),
            'date_range': self._date_range_for(df),
        }]
        return df

    def _excel_kwargs(self, source_file: str) -> dict:
        return {'engine': 'xlrd'} if source_file.lower().endswith('.xls') else {}

    def _read_raw_first_sheet(self, path_or_file, source_file: str) -> pd.DataFrame:
        kwargs = self._excel_kwargs(source_file)
        book = pd.ExcelFile(path_or_file, **kwargs)
        return pd.read_excel(path_or_file, sheet_name=book.sheet_names[0], header=None, dtype=object, **kwargs)

    def _find_header_row(self, raw: pd.DataFrame, required_terms: list[str]) -> int:
        for idx, row in raw.iterrows():
            values = {clean_text(v) for v in row.tolist() if clean_text(v)}
            if all(term in values for term in required_terms):
                return int(idx)
        raise ValueError('无法识别表头：' + '、'.join(required_terms))

    def _parse_bank_file(self, path_or_file, source_file: str) -> pd.DataFrame:
        raw = self._read_raw_first_sheet(path_or_file, source_file)
        flat_values = {clean_text(v) for v in raw.head(12).to_numpy().ravel().tolist() if clean_text(v)}
        if {'交易时间', '交易对手户名', '借贷标志', '交易金额'}.issubset(flat_values):
            return self._parse_jilin_bank(raw, source_file)
        if {'日期', '交易类型', '对方户名', '借方发生额', '贷方发生额'}.issubset(flat_values):
            return self._parse_icbc(raw, source_file)
        if {'账号', '账户名称', '交易时间', '借方发生额（支取）', '贷方发生额（收入）'}.issubset(flat_values):
            return self._parse_ccb(raw, source_file)
        raise ValueError('暂未识别该银行流水格式，请检查是否为吉林银行、工行或建行导出模板。')

    def _with_standard_columns(self, df: pd.DataFrame, source_file: str, bank_name: str, account_no: str, bank_code: str, account_label: str) -> pd.DataFrame:
        df = df.dropna(how='all').copy()
        df['交易时间_dt'] = pd.to_datetime(df['交易时间'], errors='coerce')
        df['交易日期'] = df['交易时间_dt'].dt.strftime('%Y-%m-%d')
        df['交易金额_num'] = df['交易金额'].map(money)
        df['来源文件'] = source_file
        df['来源银行'] = bank_name
        df['本方账号'] = account_no
        df['银行档案代码'] = bank_code
        df['银行账户名称'] = account_label
        prefix = re.sub(r'\W+', '', source_file)[-10:] or bank_name
        df['流水号'] = [f'{bank_name}-{prefix}-{i + 1:04d}' for i in range(len(df))]
        return df

    def _parse_jilin_bank(self, raw: pd.DataFrame, source_file: str) -> pd.DataFrame:
        header = self._find_header_row(raw, ['交易时间', '交易对手户名', '借贷标志', '交易金额'])
        df = raw.iloc[header + 1:].copy()
        df.columns = raw.iloc[header].tolist()
        required = {'交易时间', '交易对手户名', '借贷标志', '交易金额', '用途', '摘要'}
        missing = required - set(df.columns.astype(str))
        if missing:
            raise ValueError('缺少吉林银行流水字段：' + '、'.join(sorted(missing)))
        return self._with_standard_columns(df, source_file, '吉林银行', '8936374879000001', '2801:银行档案', '吉林银行亚泰大街支行')

    def _parse_icbc(self, raw: pd.DataFrame, source_file: str) -> pd.DataFrame:
        account_no = clean_text(raw.iloc[0, 1]) if raw.shape[0] > 0 and raw.shape[1] > 1 else ''
        header = self._find_header_row(raw, ['日期', '交易类型', '对方户名', '借方发生额', '贷方发生额'])
        source = raw.iloc[header + 1:].copy()
        source.columns = raw.iloc[header].tolist()
        source = source.dropna(how='all')
        debit = source['借方发生额'].map(money)
        credit = source['贷方发生额'].map(money)
        out = pd.DataFrame({
            '交易时间': source['日期'],
            '交易对手户名': source.get('对方户名', ''),
            '交易对手账号': source.get('对方账号', ''),
            '交易对手行名': '',
            '借贷标志': credit.gt(0).map(lambda x: '贷' if x else '借'),
            '交易金额': [c if c > 0 else d for d, c in zip(debit, credit)],
            '交易后余额': source.get('余额', ''),
            '币种': '人民币',
            '交易类型': source.get('交易类型', ''),
            '用途': source.get('摘要', ''),
            '摘要': source.get('摘要', ''),
        })
        out = out[(debit > 0) | (credit > 0)]
        return self._with_standard_columns(out, source_file, '工商银行', account_no, '0107:银行档案', '工行人民广场支行')

    def _parse_ccb(self, raw: pd.DataFrame, source_file: str) -> pd.DataFrame:
        header = self._find_header_row(raw, ['账号', '账户名称', '交易时间', '借方发生额（支取）', '贷方发生额（收入）'])
        source = raw.iloc[header + 1:].copy()
        source.columns = raw.iloc[header].tolist()
        source = source.dropna(how='all')
        debit = source['借方发生额（支取）'].map(money)
        credit = source['贷方发生额（收入）'].map(money)
        out = pd.DataFrame({
            '交易时间': source['交易时间'],
            '交易对手户名': source.get('对方户名', ''),
            '交易对手账号': source.get('对方账号', ''),
            '交易对手行名': source.get('对方开户机构', ''),
            '借贷标志': credit.gt(0).map(lambda x: '贷' if x else '借'),
            '交易金额': [c if c > 0 else d for d, c in zip(debit, credit)],
            '交易后余额': source.get('余额', ''),
            '币种': source.get('币种', ''),
            '交易类型': source.get('摘要', ''),
            '用途': source.get('备注', ''),
            '摘要': source.get('备注', ''),
        })
        out = out[(debit > 0) | (credit > 0)]
        account_no = clean_text(source['账号'].dropna().iloc[0]) if '账号' in source.columns and not source['账号'].dropna().empty else ''
        return self._with_standard_columns(out, source_file, '建设银行', account_no, '0402:银行档案', '建行富豪花园支行')

    def _date_range_for(self, df: pd.DataFrame) -> str:
        dates = sorted([d for d in df.get('交易日期', pd.Series(dtype=object)).dropna().unique().tolist() if d])
        return f'{dates[0]} 至 {dates[-1]}' if dates else ''

    def bank_summary(self) -> dict:
        grouped = []
        if not self.bank.empty:
            for (source_file, bank_name, account), df in self.bank.groupby(['来源文件', '来源银行', '本方账号'], dropna=False):
                grouped.append({
                    'file': source_file,
                    'bank': bank_name,
                    'account': account,
                    'rows': int(len(df)),
                    'income': round2(df.loc[df['借贷标志'] == '贷', '交易金额_num'].sum()),
                    'expense': round2(df.loc[df['借贷标志'] == '借', '交易金额_num'].sum()),
                    'date_range': self._date_range_for(df),
                })
        return {'accounts': grouped, 'imports': self.import_results[-20:]}

    def transactions(self, day: str = '') -> dict:
        df = self.bank.copy()
        if day:
            df = df[df['交易日期'] == day]
        df = df.sort_values(['交易时间_dt', '来源银行', '流水号'], na_position='last')
        rows = []
        for _, row in df.iterrows():
            rows.append({
                'id': clean_text(row.get('流水号')),
                'source_file': clean_text(row.get('来源文件')),
                'bank': clean_text(row.get('来源银行')),
                'account': clean_text(row.get('本方账号')),
                'bank_code': clean_text(row.get('银行档案代码')),
                'time': clean_text(row.get('交易时间')),
                'date': clean_text(row.get('交易日期')),
                'direction': '收入' if clean_text(row.get('借贷标志')) == '贷' else '支出',
                'amount': money(row.get('交易金额_num')),
                'balance': clean_text(row.get('交易后余额')),
                'counterparty': clean_text(row.get('交易对手户名')),
                'counterparty_account': clean_text(row.get('交易对手账号')),
                'counterparty_bank': clean_text(row.get('交易对手行名')),
                'type': clean_text(row.get('交易类型')),
                'purpose': clean_text(row.get('用途')),
                'summary': clean_text(row.get('摘要')),
            })
        return {'rows': rows, 'count': len(rows)}

    def import_bank_files(self, files: list[tuple[str, bytes]]) -> list[dict]:
        results = []
        frames = []
        for filename, content in files:
            suffix = Path(filename).suffix.lower()
            if suffix not in {'.xlsx', '.xls'}:
                results.append({'file': filename, 'status': '跳过', 'message': '仅支持 .xlsx 银行流水。'})
                continue
            try:
                frame = self._parse_bank_file(io.BytesIO(content), source_file=filename)
                frames.append(frame)
                results.append({
                    'file': filename,
                    'status': '已导入',
                    'bank': clean_text(frame['来源银行'].iloc[0]) if not frame.empty else '',
                    'account': clean_text(frame['本方账号'].iloc[0]) if not frame.empty else '',
                    'rows': int(len(frame)),
                    'income': round2(frame.loc[frame['借贷标志'] == '贷', '交易金额_num'].sum()),
                    'expense': round2(frame.loc[frame['借贷标志'] == '借', '交易金额_num'].sum()),
                    'date_range': self._date_range_for(frame),
                })
            except Exception as exc:
                results.append({'file': filename, 'status': '失败', 'message': str(exc)})
        if frames:
            combined = pd.concat([self.bank] + frames, ignore_index=True)
            combined['dedupe_key'] = combined.apply(lambda r: '|'.join([
                clean_text(r.get('本方账号')),
                clean_text(r.get('交易时间')),
                clean_text(r.get('借贷标志')),
                f"{money(r.get('交易金额_num')):.2f}",
                clean_text(r.get('交易对手户名')),
                clean_text(r.get('用途')),
                clean_text(r.get('摘要')),
            ]), axis=1)
            before = len(combined)
            combined = combined.drop_duplicates('dedupe_key', keep='first').drop(columns=['dedupe_key'])
            self.bank = combined
            removed = before - len(combined)
            if removed:
                results.append({'file': '去重', 'status': '完成', 'message': f'已移除重复流水 {removed} 笔。'})
        self.import_results.extend(results)
        return results

    def _load_invoices(self) -> pd.DataFrame:
        df = pd.read_excel(INVOICE_PATH, sheet_name='Sheet1', dtype=object).dropna(how='all')
        for col in ['金额', '税额', '价税合计']:
            df[f'{col}_num'] = df[col].map(money)
        df['开票日期_dt'] = pd.to_datetime(df['开票日期'], errors='coerce')
        df['购买方_norm'] = df['购买方名称'].map(norm_name)
        df['备注_text'] = df['备注'].fillna('').astype(str)
        df['管理费'] = df['备注_text'].map(lambda x: parse_memo_amount('管理费', x))
        df['扣除额'] = df['备注_text'].map(lambda x: parse_memo_amount('扣除额', x))
        return df

    def _load_subjects(self) -> dict[str, str]:
        raw = pd.read_excel(CODE_DIR / '会计科目_472.xlsx', header=None, dtype=object)
        fields = str(raw.iloc[1, 0]).strip().strip('"').split(',')
        df = raw.iloc[2:, :len(fields)].copy()
        df.columns = fields
        df = df.dropna(how='all')
        return {str(row['code']).strip(): str(row['name']).strip() for _, row in df.iterrows()}

    def _load_customers(self) -> list[dict]:
        df = pd.read_excel(CODE_DIR / '客户基本信息列表 (476)-化简版.xlsx', dtype=object).dropna(how='all')
        out = []
        for _, row in df.iterrows():
            out.append({
                'code': str(row.get('客户编码', '') or '').strip(),
                'name': str(row.get('客户名称', '') or '').strip(),
                'tax_no': str(row.get('统一社会信用代码', '') or '').strip(),
                'norm': norm_name(row.get('客户名称', '')),
            })
        return out

    def _load_suppliers(self) -> list[dict]:
        df = pd.read_excel(CODE_DIR / '供应商基本信息列表 (325条)-化简版.xlsx', dtype=object).dropna(how='all')
        out = []
        for _, row in df.iterrows():
            out.append({
                'code': str(row.get('供应商编码', '') or '').strip(),
                'name': str(row.get('供应商名称', '') or '').strip(),
                'norm': norm_name(row.get('供应商名称', '')),
            })
        return out

    def subject(self, code: str) -> str:
        return self.subjects.get(code, code)

    def customer_for(self, name: str, tax_no: str = '') -> tuple[str, str]:
        norm = norm_name(name)
        for item in self.customers:
            if tax_no and item['tax_no'] and tax_no == item['tax_no']:
                return item['code'], item['name']
        for item in self.customers:
            if norm and (norm == item['norm'] or norm in item['norm'] or item['norm'] in norm):
                return item['code'], item['name']
        return '', str(name or '')

    def supplier_for(self, name: str) -> tuple[str, str]:
        norm = norm_name(name)
        for item in self.suppliers:
            if norm and (norm == item['norm'] or norm in item['norm'] or item['norm'] in norm):
                return item['code'], item['name']
        return '', str(name or '')

    def dates(self) -> list[str]:
        present = sorted([d for d in self.bank['交易日期'].dropna().unique().tolist() if d])
        if not present:
            return []
        start = datetime.strptime(present[0], '%Y-%m-%d').date()
        end = datetime.strptime(present[-1], '%Y-%m-%d').date()
        days = []
        current = start
        while current <= end:
            days.append(current.strftime('%Y-%m-%d'))
            current += timedelta(days=1)
        return days

    def match_invoices(self, row: pd.Series) -> pd.DataFrame:
        if row['借贷标志'] != '贷':
            return self.invoices.iloc[0:0]
        amount = row['交易金额_num']
        counterparty = norm_name(row.get('交易对手户名', ''))
        day = row['交易时间_dt']
        candidates = self.invoices.copy()
        if counterparty:
            candidates = candidates[candidates['购买方_norm'].map(lambda x: counterparty in x or x in counterparty if x else False)]
        if day is not pd.NaT:
            candidates = candidates[(candidates['开票日期_dt'] - day).abs().dt.days <= 45]
        exact = candidates[(candidates['价税合计_num'] - amount).abs() <= 0.01]
        if not exact.empty:
            return exact.head(1)
        same_amount = self.invoices[(self.invoices['价税合计_num'] - amount).abs() <= 0.01]
        return same_amount.head(1)

    def voucher_for_bank_row(self, row: pd.Series, number: int | None = None) -> Voucher:
        dt = row['交易时间_dt']
        voucher_date = dt.strftime('%Y-%m-%d') if pd.notna(dt) else str(row.get('交易日期', ''))
        amount = row['交易金额_num']
        direction = 'in' if row['借贷标志'] == '贷' else 'out'
        text = ' '.join([str(row.get('用途', '') or ''), str(row.get('摘要', '') or '')]).strip()
        counterparty = clean_text(row.get('交易对手户名', ''))
        entries: list[Entry] = []
        invoice_refs: list[str] = []
        status = '通用待确认'
        confidence = '低'
        bank_entry = Entry(
            summary=text or ('银行进账' if direction == 'in' else '银行出账'),
            subject_code='1002',
            subject_name=self.subject('1002'),
            debit=amount if direction == 'in' else 0.0,
            credit=amount if direction == 'out' else 0.0,
            bank_code=clean_text(row.get('银行档案代码')) or self.bank_code,
            cashflow_code='1111' if direction == 'in' else '1121',
        )
        entries.append(bank_entry)
        if direction == 'in':
            matched = self.match_invoices(row)
            if not matched.empty:
                inv = matched.iloc[0]
                invoice_refs = [str(inv.get('数电发票号码', ''))]
                cust_code, cust_name = self.customer_for(inv.get('购买方名称', ''), str(inv.get('购方识别号', '') or ''))
                service_name = str(inv.get('货物或应税劳务名称', '') or '')
                memo = str(inv.get('备注', '') or '')
                is_outsource = '外包' in service_name
                has_dispatch_split = inv.get('管理费') is not None and inv.get('扣除额') is not None and not is_outsource
                summary = service_name or text or '银行收款'
                entries[0].summary = summary
                entries[0].customer_code = cust_code
                entries[0].customer_name = cust_name
                if has_dispatch_split:
                    fee_gross = money(inv.get('管理费'))
                    deduction = money(inv.get('扣除额'))
                    revenue = round2(fee_gross / 1.06)
                    tax = round2(fee_gross - revenue)
                    entries.extend([
                        Entry(summary, '6001', self.subject('6001'), credit=revenue, customer_code=cust_code, customer_name=cust_name, department_code='01', cashflow_code='1111'),
                        Entry(summary, '22210102', self.subject('22210102'), credit=tax, customer_code=cust_code, customer_name=cust_name),
                        Entry(summary, '2241', self.subject('2241'), credit=round2(amount - revenue - tax), customer_code=cust_code, customer_name=cust_name),
                    ])
                    status = '自动匹配-派遣拆分'
                    confidence = '中'
                else:
                    revenue = abs(money(inv.get('金额')))
                    tax = abs(money(inv.get('税额')))
                    if inv.get('价税合计_num', 0) < 0:
                        revenue = -revenue
                        tax = -tax
                    entries.extend([
                        Entry(summary, '6001', self.subject('6001'), credit=round2(revenue), customer_code=cust_code, customer_name=cust_name, department_code='01', cashflow_code='1111'),
                        Entry(summary, '22210102', self.subject('22210102'), credit=round2(tax), customer_code=cust_code, customer_name=cust_name),
                    ])
                    status = '自动匹配-外包全额' if is_outsource else '自动匹配-发票全额'
                    confidence = '中'
            else:
                cust_code, cust_name = self.customer_for(counterparty)
                entries.append(Entry(text or '未匹配收款', '1122', self.subject('1122'), credit=amount, customer_code=cust_code, customer_name=cust_name, cashflow_code='1111'))
        else:
            supplier_code, supplier_name = self.supplier_for(counterparty)
            debit_code = '660304' if ('手续费' in text or '收费' in text) else '2241'
            entries.append(Entry(text or '未匹配付款', debit_code, self.subject(debit_code), debit=amount, supplier_code=supplier_code, supplier_name=supplier_name, cashflow_code='1121'))
        return Voucher(
            id=str(row['流水号']),
            date=voucher_date,
            number=number,
            source='银行流水',
            status=status,
            confidence=confidence,
            bank_ref=str(row['流水号']),
            bank_time=str(row.get('交易时间', '')),
            bank_account=clean_text(row.get('银行账户名称')) or '吉林银行亚泰大街支行',
            counterparty=counterparty,
            amount=amount,
            direction=direction,
            description=text,
            invoice_refs=invoice_refs,
            entries=entries,
        )

    def generate_day(self, day: str, start_no: int) -> list[Voucher]:
        rows = self.bank[self.bank['交易日期'] == day].copy().sort_values('交易时间_dt')
        vouchers = []
        no = start_no
        for _, row in rows.iterrows():
            vouchers.append(self.voucher_for_bank_row(row, no))
            no += 1
        return vouchers


DATA = DemoData()


def voucher_to_dict(v: Voucher) -> dict:
    d = asdict(v)
    d['entries'] = [asdict(e) for e in v.entries]
    return d


def validate_vouchers(items: list[dict]) -> list[dict]:
    issues = []
    for v in items:
        debit = round2(sum(money(e.get('debit')) for e in v.get('entries', [])))
        credit = round2(sum(money(e.get('credit')) for e in v.get('entries', [])))
        if abs(debit - credit) > 0.01:
            issues.append({'voucher': v.get('number'), 'level': 'error', 'message': f'借贷不平：借 {debit:.2f}，贷 {credit:.2f}'})
        for i, e in enumerate(v.get('entries', []), start=1):
            if not str(e.get('subject_code', '')).strip():
                issues.append({'voucher': v.get('number'), 'level': 'error', 'message': f'第 {i} 行缺科目编码'})
        if v.get('status') == '通用待确认':
            issues.append({'voucher': v.get('number'), 'level': 'warn', 'message': '通用凭证待财务确认'})
    nums = [v.get('number') for v in items if v.get('number') is not None]
    if nums and nums != list(range(min(nums), min(nums) + len(nums))):
        issues.append({'voucher': '', 'level': 'error', 'message': '凭证号不连续'})
    return issues


def export_csv(items: list[dict]) -> bytes:
    buf = io.StringIO()
    writer = csv.writer(buf)
    writer.writerow(['凭证日期', '凭证类别', '凭证号', '摘要', '科目编码', '科目名称', '借方金额', '贷方金额', '客户编码', '客户名称', '供应商编码', '供应商名称', '部门', '项目', '银行档案', '现金流量项目', '来源流水', '关联发票'])
    for v in items:
        for e in v.get('entries', []):
            writer.writerow([
                v.get('date', ''), '记账凭证', v.get('number', ''), e.get('summary', ''), e.get('subject_code', ''), e.get('subject_name', ''),
                f"{money(e.get('debit')):.2f}" if money(e.get('debit')) else '',
                f"{money(e.get('credit')):.2f}" if money(e.get('credit')) else '',
                e.get('customer_code', ''), e.get('customer_name', ''), e.get('supplier_code', ''), e.get('supplier_name', ''),
                e.get('department_code', ''), e.get('project_code', ''), e.get('bank_code', ''), e.get('cashflow_code', ''),
                v.get('bank_ref', ''), ';'.join(v.get('invoice_refs', [])),
            ])
    return buf.getvalue().encode('utf-8-sig')


class Handler(BaseHTTPRequestHandler):
    def send_json(self, payload, status=200):
        body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed = urlparse(self.path)
        if parsed.path == '/':
            body = (ROOT / 'index.html').read_bytes()
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        if parsed.path == '/api/dates':
            self.send_json({'dates': DATA.dates()})
            return
        if parsed.path == '/api/banks':
            self.send_json(DATA.bank_summary())
            return
        if parsed.path == '/api/transactions':
            qs = parse_qs(parsed.query)
            day = qs.get('date', [''])[0]
            self.send_json(DATA.transactions(day))
            return
        if parsed.path == '/api/generate':
            qs = parse_qs(parsed.query)
            day = qs.get('date', [DATA.dates()[0]])[0]
            start_no = int(qs.get('start', ['1'])[0] or 1)
            vouchers = [voucher_to_dict(v) for v in DATA.generate_day(day, start_no)]
            self.send_json({'date': day, 'start': start_no, 'vouchers': vouchers, 'issues': validate_vouchers(vouchers)})
            return
        self.send_response(404)
        self.end_headers()

    def do_POST(self):
        length = int(self.headers.get('Content-Length', '0'))
        if self.path == '/api/import':
            ctype = self.headers.get('Content-Type', '')
            if 'multipart/form-data' not in ctype:
                self.send_json({'results': [{'file': '', 'status': '失败', 'message': '请用文件上传方式导入。'}]}, status=400)
                return
            env = {
                'REQUEST_METHOD': 'POST',
                'CONTENT_TYPE': ctype,
                'CONTENT_LENGTH': str(length),
            }
            form = cgi.FieldStorage(fp=self.rfile, headers=self.headers, environ=env)
            file_fields = form['files'] if 'files' in form else []
            if not isinstance(file_fields, list):
                file_fields = [file_fields]
            files = []
            for field in file_fields:
                if getattr(field, 'filename', None):
                    files.append((field.filename, field.file.read()))
            results = DATA.import_bank_files(files)
            self.send_json({'results': results, 'summary': DATA.bank_summary(), 'dates': DATA.dates()})
            return

        payload = json.loads(self.rfile.read(length) or b'{}')
        if self.path == '/api/validate':
            self.send_json({'issues': validate_vouchers(payload.get('vouchers', []))})
            return
        if self.path == '/api/export':
            body = export_csv(payload.get('vouchers', []))
            self.send_response(200)
            self.send_header('Content-Type', 'text/csv; charset=utf-8')
            self.send_header('Content-Disposition', 'attachment; filename="voucher-demo-export.csv"')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return
        self.send_response(404)
        self.end_headers()


if __name__ == '__main__':
    server = ThreadingHTTPServer(('127.0.0.1', 8765), Handler)
    print('Voucher demo running at http://127.0.0.1:8765')
    server.serve_forever()
