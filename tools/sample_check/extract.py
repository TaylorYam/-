"""從 order sample 資料夾的 .msg 抽出起單段落，供 run_fill.ps1 回填與 check.py 比對。

用法：python -I -X utf8 extract.py <sample 資料夾> <輸出資料夾>

- 起單段落：從玉山的 "We buy" 起，到下一個簽名檔或 From: 為止（和 FCN 樣本的貼法相同）。
- 上手回覆裡有日期時，從信的開頭貼到起單段落結束，和 FCN 樣本的貼法相同：
  - HSBC：回覆 final TS 時的日期區（Strike Date／Issue Date／Final Valuation／Maturity Date，左右一組）。
  - GS：回覆裡的 Order Summary 表（TDCC Code／Issue Date／Final Valuation Date／Maturity Date）。
- 沒有 "We buy" 的信（例：BARC 的 BARX COMET 成交確認）不是起單內容，略過。

輸出：<輸出資料夾>/html/NNN.htm 與 manifest.json（每封信的表格、標題: 數值、TDCC）。
輸出含客戶資料，請放在 repo 以外（例如暫存資料夾）。
"""
import sys, os, re, json, html as htmlmod
from html.parser import HTMLParser
import extract_msg
import RTFDE.text_extraction as _rtf

# 部分 BARC 信的 RTF 裡有不合 cp950 的位元組，extract_msg 會整封讀不到；
# 這些字元改用替代字元（U+FFFD），起單段落裡有替代字元時記在 manifest 的 replaced_chars
_decode_hex_char = _rtf.decode_hex_char
def _lenient_decode_hex_char(item, codec):
    try:
        return _decode_hex_char(item, codec)
    except UnicodeDecodeError:
        return item.decode(codec or 'CP1252', errors='replace').encode()
_rtf.decode_hex_char = _lenient_decode_hex_char

ISSUER_BY_PREFIX = {'029': 'BARC', '037': 'SG', '076': 'JPM', '078': 'DBS', '082': 'BNP', '083': 'SC',
                    '086': 'CA', '092': 'UBS', '147': 'MS', '148': 'GS', '156': 'Nomura', '159': 'Citi',
                    '325': 'HSBC', '326': 'BBVA', '328': 'Natixis'}
SIGNATURES = ('玉山證券 E.SUN', 'From:', '寄件者:')
HSBC_DATE_LABELS = ('Strike Date', 'Issue Date', 'Final Valuation(?: Date)?', 'Maturity(?: Date)?')


def has_reply_dates(issuer, before):
    """上手回覆（"We buy" 之前）有沒有 FCN 樣本也會一起貼的日期區"""
    if issuer == 'HSBC':
        return all(re.search(lb + r' \d{1,2}/\d{1,2}/\d{4}', before) for lb in HSBC_DATE_LABELS)
    if issuer == 'GS':
        return 'Order Summary' in before and 'Final Valuation Date' in before
    return False


class _Rows(HTMLParser):
    """把 HTML 表格讀成列；同一格裡的多段文字以空白連接。"""
    def __init__(self):
        super().__init__(); self.rows = []; self.row = None; self.cell = None
    def handle_starttag(self, tag, attrs):
        if tag == 'tr': self.row = []
        elif tag in ('td', 'th') and self.row is not None: self.cell = []
        elif tag in ('p', 'br', 'div') and self.cell is not None: self.cell.append(' ')
    def handle_endtag(self, tag):
        if tag in ('td', 'th') and self.cell is not None:
            self.row.append(re.sub(r'\s+', ' ', ''.join(self.cell).replace('　', ' ')).strip())
            self.cell = None
        elif tag == 'tr' and self.row is not None:
            self.rows.append(self.row); self.row = None
    def handle_data(self, d):
        if self.cell is not None: self.cell.append(d)


def html_text(h):
    h = re.sub(r'(?is)<(style|head)\b.*?</\1>', ' ', h)
    h = re.sub(r'(?i)<br[^>]*>|</p>|</div>|</tr>', '\n', h)
    h = re.sub(r'(?i)</t[dh]>', '\t', h)
    t = htmlmod.unescape(re.sub(r'<[^>]+>', '', h))
    return '\n'.join(re.sub(r'[ \xa0　]+', ' ', l).strip(' \t') for l in t.splitlines())


def main(src, out):
    os.makedirs(os.path.join(out, 'html'), exist_ok=True)
    manifest = []
    files = sorted(f for f in os.listdir(src) if f.lower().endswith('.msg'))
    for i, fn in enumerate(files, 1):
        rec = {'no': i, 'file': fn, 'html': f'{i:03d}'}
        name_tdcc = (re.findall(r'(\d{12})', fn) or [''])[-1]
        rec['issuer'] = ISSUER_BY_PREFIX.get(name_tdcc[:3], '')
        try:
            m = extract_msg.Message(os.path.join(src, fn)); h = m.htmlBody; m.close()
        except Exception as e:
            rec['skip'] = f'讀不到 .msg：{e}'; manifest.append(rec); continue
        if isinstance(h, bytes): h = h.decode('utf-8', 'replace')
        h = h or ''
        p = re.search(r'We\s+buy', h, re.I)
        if not p:
            rec['skip'] = '信裡沒有 "We buy"（不是起單內容，例如成交確認）'; manifest.append(rec); continue
        p = p.start()
        ends = [x for x in (h.find(s, p) for s in SIGNATURES) if x > 0]
        end = min(ends) if ends else len(h)
        start = max(h.rfind('<p', 0, p), h.rfind('<div', 0, p))
        before = re.sub(r'\s+', ' ', html_text(h[:p]))
        rec['reply_dates'] = has_reply_dates(rec['issuer'], before)
        b = re.search(r'<body[^>]*>', h, re.I)
        body_start = b.end() if b else 0
        if rec['reply_dates']: start = body_start
        head = h[:body_start] if b else '<html><body>'
        if '<meta charset' not in head.lower(): head = '<meta charset="utf-8">' + head
        seg = h[start:end]
        rec['replaced_chars'] = seg.count('\ufffd')
        with open(os.path.join(out, 'html', rec['html'] + '.htm'), 'w', encoding='utf-8') as f:
            f.write(head + seg + '</body></html>')

        text = html_text(seg)
        rec['text'] = text
        rows = _Rows(); rows.feed(seg)
        rec['rows'] = [r for r in rows.rows if any(r)]
        manifest.append(rec)

    with open(os.path.join(out, 'manifest.json'), 'w', encoding='utf-8') as f:
        json.dump(manifest, f, ensure_ascii=False, indent=1)
    print('信件', len(manifest), '起單段落', sum('skip' not in r for r in manifest),
          '略過', sum('skip' in r for r in manifest))


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
