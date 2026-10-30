"""比對 run_fill.ps1 的回填結果和信件表格直接算出的應填值，逐欄列出差異。

用法：python -I -X utf8 check.py <manifest.json> <result.txt> [保本型商品]

保本型商品（例：Range Accrual）的樣本，run_fill.ps1 用 -ProtectedProduct 在 E5 選商品，這裡也要給同一個商品名稱。

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


def norm(s):
    """轉小寫，只留英數與「?」；玉山格式的 Issue date (T+7)：T+ 後面 1～2 位數的天數不影響表頭（和 OrderFill.bas 的 NormKey 一致）"""
    k = re.sub(r'[^a-z0-9?]', '', (s or '').lower())
    return 'issuedate' if re.fullmatch(r'issuedatet\d{1,2}', k) else k


def last_value(s):
    """改價取最後的值：97.50%→97.45%、99.10->99.05"""
    s = (s or '').strip()
    parts = re.split(r'→|->|=>|>>|(?<=[\d%])\s*>\s*(?=\d)', s)
    return parts[-1].strip()


def num(s):
    s = re.sub(r'\s+', '', last_value(s).replace('%', ''))   # 數字中間的空白不算（例：HSBC 的 82 .00%）
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
                'underlying1', 'stock1', 'strikedate', 'couponpa',
                'tenory', 'condition1', 'callfeature', 'rarate', 'lowerbarrier', 'upperbarrier')   # Range Accrual


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
            'range accrual': 'Range Accrual', 'ben': 'Regular BEN', 'regular ben': 'Regular BEN',
            'ben with lookback': 'Regular BEN', 'digital note': 'Digital BEN', 'dcn': 'Digital BEN',
            'digit': 'Digital BEN', 'digit with lookback': 'Digital BEN', 'bonus rcn': 'Digital BEN',
            'tw dra': 'Range Accrual', 'rates cran': 'Range Accrual', 'rates autocall ran': 'Range Accrual'}
# V2（Floor Coupon）填 0 的商品
FLOOR_ZERO = ('FCN', 'DACN', 'FCN_Express', 'Super Bull', 'Regular BEN', 'Digital BEN')
BEN = ('Regular BEN', 'Digital BEN')
RA_COUPON = 'couponcouponfrequencycallablefrequency'   # 玉山詢價格式的配息欄


def ra_noncall(t):
    """Range Accrual 的 Non-call（月）：Non call period (Y) × 12；NC 1Y*3M、NC 0.25Y X 3M、1Y 取 NC 後面的期間。
    Is Callable 是 No、或不是整月數時工具留空"""
    if pick(t, 'iscallable').lower() == 'no': return None
    y = num(pick(t, 'noncallperiody'))
    if y is not None: m = y * 12
    else:
        v = re.sub(r'\s+', '', pick(t, 'callfeature', 'autocallfeature', 'nc')).upper()
        g = re.match(r'(?:NC)?(\d+(?:\.\d+)?)([YM])', v)
        if not g: return None
        m = float(g[1]) * (12 if g[2] == 'Y' else 1)
    return int(m) if m == int(m) else None


def ra_underlying(t):
    """只有 USD SOFR CMS 10Y、USD CMT 10Y；rate 前面只能有 USD、SOFR 或條件符號（EUR CMS 10Y 不算）"""
    v = re.sub(r'\s+', '', pick(t, 'underlying', 'index', 'condition1')).upper()
    n = len(re.findall(r'CMS|CMT', v))
    def usd(rate): return n == 1 and re.search(r'(?:^|[^A-Z])(?:USD)?(?:SOFR)?' + rate, v) is not None
    if v == 'USS10Y' or usd('CMS10Y'): return 'USD SOFR CMS 10Y'
    if v == 'UCMT10' or usd('CMT10Y'): return 'USD CMT 10Y'
    return ''


def ra_tenor(t):
    y = num(pick(t, 'tenory'))
    if y is not None: return y
    g = re.fullmatch(r'(\d+(?:\.\d+)?)Y', re.sub(r'\s+', '', pick(t, 'tenorm', 'tenor', 'tenormonths')).upper())
    return num(g[1]) if g else None


def ra_floor(t):
    """Floor@ 的值；沒有 Floor 填 0，Floor 後面讀不到數字時工具留空"""
    c = t.get(RA_COUPON, '')
    if not re.search(r'(?i)floor', c): return 0
    g = re.search(r'(?i)floor\s*@?\s*(\d+(?:\.\d+)?)', c)
    return num(g[1]) if g else None


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


def expected(rec, t, inl, protected=None):
    issuer = rec['issuer']
    prod_raw = pick(t, 'product')
    prod = PRODUCTS.get(re.sub(r'\s+', ' ', prod_raw.lower()), prod_raw)
    if issuer == 'MS' and prod == 'Range Accrual' and pick(t, 'putstrike'): prod = 'DACN'
    if not prod_raw and issuer == 'SG':
        prod = {'first period': 'DACN', 'first two periods': 'DACN', 'all periods': 'FCN'}.get(
            re.sub(r'\s+', ' ', pick(t, 'fixedcoupons').lower()), '')
    if protected: prod = protected   # 信裡的 Product 是別的商品時工具不填入，會列為不一致
    both = {**inl, **t}
    tdcc = re.sub(r'\D', '', pick(both, 'tdcc', 'tdcccode'))
    e = {'A2': prod, 'B2': tdcc if len(tdcc) == 12 else '', 'C2': issuer,
         'D2': pick(t, 'currency', 'curr', 'ccy').upper(), 'Q2': 'M', 'V2': 0 if prod in FLOOR_ZERO else None}
    for k in range(1, 6):
        e['FGHIJ'[k - 1] + '2'] = pick(t, f'bbgcode{k}', f'underlying{k}', f'stock{k}').upper()
    pg = re.search(r'PG rate\s*:\s*([\d.]+)\s*%', rec['text'], re.I)
    e['T2'] = num(pg[1]) if pg else None
    if prod == 'Range Accrual':   # 只填 A～F、O～Q、T、V
        e.update({c + '2': None for c in 'GHIJKLMNRS'})
        e.update({'E2': ra_noncall(t), 'F2': ra_underlying(t), 'Q2': 'Y', 'V2': ra_floor(t)})
    elif prod == 'Super Bull':
        hc = issuer in ('HSBC', 'CA')
        e['E2'] = num(pick(t, 'autocallfrequencym', 'autocallfreqm', 'autocalldateonetime'))
        e['K2'] = num(pick(t, 'strike') if hc else pick(t, 'superbulllevel'))
        e['L2'] = 'Period End'
        e['M2'] = num(pick(t, 'autocallbarrier') if hc else pick(t, 'superbulllevel'))
        e['N2'] = num(pick(t, 'autocallcouponflat', 'bonuscouponflat', 'autocallflatcoupon', 'bonuscoupon',
                           'couponflat'))
        e['R2'] = ''; e['S2'] = ''
    elif prod in BEN:   # E2 = Tenor（下面算），KO Barrier 有值時留空；M2 沒有 Coupon Barrier 時 = Strike
        e['K2'] = num(pick(t, 'strike', 'putstrike'))
        e['L2'] = 'Period End'
        cb = pick(t, 'couponbarrier')
        e['M2'] = num(cb) if cb else e['K2']
        e['N2'] = num(pick(t, 'bonuscouponflat', 'bonuscoupon', 'couponflat'))
        bt = barrier(t)
        e['R2'] = bt
        e['S2'] = None if bt == 'NONE' else num(pick(t, 'kibarrier', 'putki', 'ki'))
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
    e['O2'] = num(pick(t, 'upfrontnoteprice', 'noteprice', 'cost', 'reoffer', 'offerprice', 'issueprice'))
    if e['O2'] is None and prod == 'Range Accrual':   # 玉山格式只有 UF
        uf = num(pick(t, 'uf'))
        e['O2'] = None if uf is None else round(100 - uf, 4)
    tenor = num(pick(t, 'tenorm', 'tenor', 'tenormonths'))
    if tenor is None and pick(t, 'settlementfrequency').lower() == 'monthly': tenor = num(pick(t, 'noofperiods'))
    e['P2'] = ra_tenor(t) if prod == 'Range Accrual' else tenor
    if prod in BEN: e['E2'] = None if pick(t, 'kobarrier', 'autocall', 'ko') else tenor
    e['O5'] = date(pick(both, 'tradedate', 'strikedate', 'initialvaluationdate'))
    e['O6'] = date(pick(both, 'issuedate', 'issuancedate', 'issuedatetcalendardays'))
    e['O7'] = date(pick(both, 'finalvaluationdate', 'finalfixingdate', 'redemptionvaluationdate', 'finalvaluation'))
    e['O8'] = date(pick(both, 'maturitydate', 'redemptiondate', 'finalmaturitydate'))
    ip = re.search(r'Initial price\s*:\s*(vwap|opening price|closing price|[A-Za-z]+)', rec['text'], re.I)
    e['O9'] = {'vwap': 'VWAP', 'opening price': '開盤價', 'closing price': '收盤價'}.get(
        re.sub(r'\s+', ' ', ip[1]).strip().lower() if ip else '', '')
    if prod == 'Range Accrual':   # O7、O9 不填（O7 信裡有也不填）
        e['O7'] = e['O9'] = ''
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


# 日期找不到或格式不認得（例：BARC 的 Issue Date 是 1w）時，工具填今天或由 Term 推算
INFER = {'O5': 'Trade date', 'O6': 'Issue date', 'O7': 'Final valuation date', 'O8': 'Maturity date'}


def inferred(k, text):
    return f'找不到 {INFER[k]}' in text or f'{INFER[k]} 格式不認得' in text


def main(manifest_path, result_path, protected=None):
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
        e = expected(rec, t, {**reply_dates_of(rec['rows']), **inline_of(rec['text'])}, protected)
        diffs = []
        for k, v in e.items():
            act = res['cells'].get(k, '')
            if k in INFER and v is None and inferred(k, res['text']):
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
    main(*sys.argv[1:4])
