"""比對 run_fill.ps1 的回填結果和信件表格直接算出的應填值，逐欄列出差異。

用法：python -I -X utf8 check.py <manifest.json> <result.txt>

應填值由這支程式自己讀表格計算（不呼叫 VBA），規則照 docs/order-fill.md。結果分成：
- 一致：包含用到推算的信。信裡沒有日期、工具有提醒時（O5 填今天、O6～O8 由 Term 推算），
  這些欄位不算不一致，另外統計在「推算」。
- 不一致：值不同，需要人工判斷是工具錯、信件特殊，還是這支程式的規則不足。
- 略過：不是起單內容（extract.py 已標記）。
- 沒有填入：工具沒有回填（例：判斷不出上手），要看摘要。
新增商品時，在 expected() 補上該商品的欄位規則，並把商品名稱加進 PRODUCTS 與 FLOOR_ZERO；
表頭偵測認不得該商品的表頭時，再補 HEADER_WORDS。
目前沒有 Citi（Daily KO／Memory Autocall、Non Callable + 1、100 − Upfront）與 SG 推定的完整條件；
遇到這些信時差異會列為不一致（不會漏報），請看原信判斷。
"""
import sys, json, re, collections, datetime


def norm(s): return re.sub(r'[^a-z0-9?]', '', (s or '').lower())


def last_value(s):
    """改價取最後的值：97.50%→97.45%、99.10->99.05"""
    s = (s or '').strip()
    parts = re.split(r'→|->|=>|>>|(?<=[\d%])\s*>\s*(?=\d)', s)
    return parts[-1].strip()


def num(s):
    s = last_value(s).replace('%', '').replace('\xa0', '').strip()
    s = re.sub(r'(?i)m$', '', s).strip()
    if not re.fullmatch(r'-?\d+(\.\d+)?', s): return None
    v = float(s)
    return int(v) if v == int(v) else round(v, 4)


MON = {m: i for i, m in enumerate('jan feb mar apr may jun jul aug sep oct nov dec'.split(), 1)}


def date(s):
    s = (s or '').strip()
    for pat, order in ((r'(\d{4})[-/](\d{1,2})[-/](\d{1,2})', 'ymd'), (r'(\d{1,2})-([A-Za-z]{3})-(\d{2,4})', 'dMy'),
                       (r'(\d{1,2})([A-Za-z]{3})(\d{2})\b', 'dMy'), (r'(\d{1,2}) ([A-Za-z]{3}) (\d{4})', 'dMy'),
                       (r'(\d{1,2})/(\d{1,2})/(\d{4})', 'dmy')):
        m = re.match(pat, s)
        if not m: continue
        if order == 'ymd': y, mo, d = int(m[1]), int(m[2]), int(m[3])
        elif order == 'dmy': y, mo, d = int(m[3]), int(m[2]), int(m[1])
        else:
            y, mo, d = int(m[3]), MON.get(m[2].lower()), int(m[1])
            if y < 100: y += 2000
        if mo: return f'{y}/{mo}/{d}'
    return None


HEADER_WORDS = ('product', 'currency', 'curr', 'strike', 'putstrike', 'kotype', 'tenorm', 'bbgcode1',
                'underlying1', 'stock1', 'strikedate', 'couponpa')


def table_of(rows):
    """第一個表頭列（3 格以上認得）和下一列組成 {表頭: 值}；回傳 (dict, 表格數)"""
    tables = []
    for i, r in enumerate(rows):
        if sum(norm(c) in HEADER_WORDS or norm(c).startswith('strike') for c in r) >= 3 and i + 1 < len(rows):
            v = rows[i + 1]
            tables.append({norm(c): (v[j] if j < len(v) else '') for j, c in enumerate(r) if c})
    return (tables[0] if tables else {}), len(tables)


def inline_of(text):
    """「標題: 數值」與 GS 的 [ 標題: 數值 ]"""
    d = {}
    for line in text.splitlines():
        for seg in re.split(r'[\[\]]', line):
            m = re.match(r'\s*([^:：]{1,40})[:：]\s*(.+?)\s*$', seg)
            if m: d.setdefault(norm(m[1]), m[2])
    return d


DATE_KEYS = ('tradedate', 'strikedate', 'initialvaluationdate', 'issuedate', 'issuancedate', 'finalvaluationdate',
             'finalfixingdate', 'redemptionvaluationdate', 'finalvaluation', 'maturitydate', 'redemptiondate',
             'finalmaturitydate')


def reply_dates_of(rows):
    """上手回覆的日期：左右一組（HSBC 的日期區）、表頭列加數值列（GS 的 Order Summary）"""
    d = {}
    for i, r in enumerate(rows):
        for j, c in enumerate(r):
            k = norm(c)
            if k not in DATE_KEYS: continue
            if j + 1 < len(r) and date(r[j + 1]): d.setdefault(k, r[j + 1])
            elif i + 1 < len(rows) and j < len(rows[i + 1]) and date(rows[i + 1][j]): d.setdefault(k, rows[i + 1][j])
    return d


def pick(d, *keys):
    for k in keys:
        v = d.get(k, '')
        if v and v.strip() and v.strip().upper() not in ('N/A', 'NA', '-'): return v.strip()
    return ''


# 信裡的 Product 寫法 → Variable!D 的商品名稱（對照 OrderFill.bas 的 NormProduct）
PRODUCTS = {'fcn': 'FCN', 'fca': 'FCN', 'dac': 'DACN', 'dra': 'DACN', 'vmran': 'DACN', 'dacn': 'DACN',
            'express fcn': 'FCN_Express', 'superbull': 'Super Bull', 'super bull': 'Super Bull',
            'range accrual': 'Range Accrual'}
# V2（Floor Coupon）填 0 的商品
FLOOR_ZERO = ('FCN', 'DACN', 'FCN_Express', 'Super Bull')


def ko_type(t, issuer):
    k = (t.get('kotype') or '').strip().lower()
    if issuer == 'MS':
        mem = (t.get('komemory') or '').strip().lower()
        base = {'daily': 'Daily', 'daily memory': 'Daily', 'observation': 'Period End'}.get(k)
        if k == 'daily memory' and mem != 'yes': return ''   # 和 KO Memory 矛盾，工具留空
        if base and mem in ('yes', 'no'): return base + (' Memory' if mem == 'yes' else '')
        return ''
    return {'daily': 'Daily', 'daily memory': 'Daily Memory', 'period end': 'Period End',
            'period end memory': 'Period End Memory'}.get(re.sub(r'\s+', ' ', k), '')


def barrier(t):
    v = pick(t, 'barriertype', 'putki?', 'kitype').lower()
    if v in ('european', 'at maturity', 'eki'): return 'EKI'
    if v in ('american', 'aki'): return 'AKI'
    if v in ('mki',): return 'MKI'
    if v in ('none', 'no'): return 'NONE'
    if v == '':   # 沒有 Barrier Type：KI 也沒有才是 NONE；有 KI 時工具留空並提醒
        return 'NONE' if num(pick(t, 'kibarrier', 'putki', 'ki')) is None else ''
    return '?' + v


def expected(rec, t, inl):
    issuer = rec['issuer']
    prod_raw = pick(t, 'product')
    prod = PRODUCTS.get(re.sub(r'\s+', ' ', prod_raw.lower()), prod_raw)
    if issuer == 'MS' and prod == 'Range Accrual' and pick(t, 'putstrike'): prod = 'DACN'
    if not prod_raw and issuer == 'SG':
        prod = {'first period': 'DACN', 'first two periods': 'DACN', 'all periods': 'FCN'}.get(pick(t, 'fixedcoupons').lower(), '')
    both = {**inl, **t}
    tdcc = re.sub(r'\D', '', pick(both, 'tdcc', 'tdcccode'))
    e = {'A2': prod, 'B2': tdcc if len(tdcc) == 12 else '', 'C2': issuer,
         'D2': pick(t, 'currency', 'curr').upper(), 'Q2': 'M', 'V2': 0 if prod in FLOOR_ZERO else None}
    for k in range(1, 6):
        e['FGHIJ'[k - 1] + '2'] = pick(t, f'bbgcode{k}', f'underlying{k}', f'stock{k}').upper()
    pg = re.search(r'PG rate\s*:\s*([\d.]+)\s*%', rec['text'], re.I)
    e['T2'] = num(pg[1]) if pg else None
    if prod == 'Super Bull':
        hc = issuer in ('HSBC', 'CA')
        e['E2'] = num(pick(t, 'autocallfrequencym', 'autocallfreqm', 'autocalldateonetime'))
        e['K2'] = num(pick(t, 'strike') if hc else pick(t, 'superbulllevel'))
        e['L2'] = 'Period End'
        e['M2'] = num(pick(t, 'autocallbarrier') if hc else pick(t, 'superbulllevel'))
        e['N2'] = num(pick(t, 'autocallcouponflat', 'bonuscouponflat', 'autocallflatcoupon'))
        e['R2'] = ''; e['S2'] = ''
    else:   # FCN、DACN、FCN_Express
        e['E2'] = 0 if prod == 'FCN_Express' else num(pick(t, 'guaranteedperiodsm', 'guaranteedperiods',
                                                           'noncallperiodsm', 'noncallm', 'noncallableperiods'))
        e['K2'] = num(pick(t, 'strike', 'putstrike'))
        e['L2'] = ko_type(t, issuer)
        e['M2'] = num(pick(t, 'kobarrier', 'autocall', 'ko'))
        e['N2'] = num(pick(t, 'couponpa', 'coupon'))
        bt = barrier(t)
        e['R2'] = bt
        e['S2'] = None if bt == 'NONE' else num(pick(t, 'kibarrier', 'putki', 'ki'))
    e['O2'] = num(pick(t, 'upfrontnoteprice', 'noteprice', 'cost', 'reoffer', 'offerprice'))
    tenor = num(pick(t, 'tenorm', 'tenor'))
    if tenor is None and pick(t, 'settlementfrequency').lower() == 'monthly': tenor = num(pick(t, 'noofperiods'))
    e['P2'] = tenor
    e['O5'] = date(pick(both, 'tradedate', 'strikedate', 'initialvaluationdate'))
    e['O6'] = date(pick(both, 'issuedate', 'issuancedate', 'issuedatetcalendardays'))
    e['O7'] = date(pick(both, 'finalvaluationdate', 'finalfixingdate', 'redemptionvaluationdate', 'finalvaluation'))
    e['O8'] = date(pick(both, 'maturitydate', 'redemptiondate', 'finalmaturitydate'))
    ip = re.search(r'Initial price\s*:\s*(vwap|opening price|closing price|[A-Za-z]+)', rec['text'], re.I)
    e['O9'] = {'vwap': 'VWAP', 'opening price': '開盤價', 'closing price': '收盤價'}.get(
        re.sub(r'\s+', ' ', ip[1]).strip().lower() if ip else '', '')
    return e


def same(exp, act):
    if exp is None or exp == '': return act == ''
    if isinstance(exp, (int, float)):
        try: return abs(float(act) - exp) < 1e-6
        except ValueError: return False
    return act == exp


def parse_results(path):
    res = {}
    for block in open(path, encoding='utf-8-sig').read().split('=== ')[1:]:
        lines = block.splitlines()
        cells = dict(kv.split('=', 1) for kv in lines[1].split(' | ')) if len(lines) > 1 and '=' in lines[1] else {}
        res[lines[0].strip()] = {'cells': cells, 'notes': [l[2:] for l in lines[2:] if l.startswith('- ')],
                                 'text': ' '.join(lines[2:])}
    return res


INFER = {'O5': '找不到 Trade date', 'O6': '找不到 Issue date', 'O7': '找不到 Final valuation date',
         'O8': '找不到 Maturity date'}


def main(manifest_path, result_path):
    manifest = json.load(open(manifest_path, encoding='utf-8'))
    results = parse_results(result_path)
    today = datetime.date.today()
    today_s = f'{today.year}/{today.month}/{today.day}'
    stat = collections.Counter(); infer = collections.Counter(); notes = collections.Counter()
    for rec in manifest:
        tag = f"{rec['no']:03d} [{rec['issuer'] or '?'}] {rec['file'][:90]}"
        if 'skip' in rec:
            stat['略過'] += 1; print(f'[略過] {tag}\n       {rec["skip"]}'); continue
        res = results.get(rec['html'])
        if not res or not res['cells']:
            stat['沒有填入'] += 1; print(f'[沒有填入] {tag}\n       {res["text"] if res else "沒有結果"}'); continue
        t, n_tables = table_of(rec['rows'])
        e = expected(rec, t, {**reply_dates_of(rec['rows']), **inline_of(rec['text'])})
        diffs = []
        for k, v in e.items():
            act = res['cells'].get(k, '')
            if k in INFER and v is None and INFER[k] in res['text']:
                if k != 'O5' or act == today_s: infer[k] += 1; continue
            if not same(v, act): diffs.append(f'{k} 應為 [{v}] 實際 [{act}]')
        for nt in res['notes']: notes[re.sub(r'\d[\d./%,]*', '#', nt)] += 1
        flags = []
        if n_tables > 1: flags.append(f'{n_tables} 張表')
        if rec.get('replaced_chars'): flags.append(f'起單段落有 {rec["replaced_chars"]} 個替代字元')
        if rec.get('reply_dates'): flags.append('含上手回覆的日期')
        if diffs:
            stat['不一致'] += 1
            print(f'[不一致] {tag}  {"；".join(flags)}')
            for d in diffs: print('       ' + d)
            for nt in res['notes']: print('       提醒：' + nt)
        else:
            stat['一致'] += 1
            if flags: print(f'[一致] {tag}  {"；".join(flags)}')
    print('\n== 統計', dict(stat))
    print('== 推算（信裡沒有、工具有提醒）', dict(infer))
    print('== 提醒種類（數字改成 #）')
    for nt, c in notes.most_common(): print(f'{c:4d}  {nt}')


if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
