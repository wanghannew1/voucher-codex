from pathlib import Path
import re

paths = [
    Path('/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹工行4月对账单.xls'),
    Path('/Users/sunyitong/coding/voucher/代码资料/银行流水/彩虹建行4月对账单.xls'),
]

allowed_extra = set('，。；：、（）()-_/.:0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ')

def is_useful_char(ch):
    return ('\u4e00' <= ch <= '\u9fff') or ch in allowed_extra

def extract_utf16le_strings(data, min_chars=2):
    text = data.decode('utf-16le', errors='ignore')
    runs = []
    buf = []
    for ch in text:
        if is_useful_char(ch):
            buf.append(ch)
        else:
            if len(buf) >= min_chars:
                runs.append(''.join(buf))
            buf = []
    if len(buf) >= min_chars:
        runs.append(''.join(buf))
    cleaned = []
    for s in runs:
        s = re.sub(r'\s+', ' ', s).strip()
        if s and any(('\u4e00' <= ch <= '\u9fff') or ch.isdigit() for ch in s):
            cleaned.append(s)
    return cleaned

for path in paths:
    data = path.read_bytes()
    print('\n==== ' + path.name + ' ====')
    strings = extract_utf16le_strings(data)
    seen = []
    for s in strings:
        if s not in seen:
            seen.append(s)
    for s in seen[:260]:
        print(s)
