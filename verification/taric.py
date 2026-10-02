"""EU TARIC (official EU tariff database) reader. For a CN8 code, loads measures.jsp (creates the session), then every
measures_details.jsp sub-table (one per 10-digit goods code shown) and parses the duty rows."""
import requests, re, html, json, time, sys, io, os
H = {'User-Agent': 'Mozilla/5.0'}
CACHE = 'taric_cache.json'
_cache = json.load(open(CACHE, encoding='utf-8')) if os.path.exists(CACHE) else {}
TYPES = ('Third country duty', 'Tariff preference', 'Preferential tariff quota', 'Non preferential tariff quota',
         'Autonomous tariff suspension', 'Airworthiness tariff suspension', 'Customs Union Duty', 'Preferential suspension')


def _clean(x):
    x = x.replace(chr(13), ' ').replace(chr(10), ' ')
    b = re.sub(r'<script.*?</script>', ' ', x, flags=re.S); b = re.sub(r'<style.*?</style>', ' ', b, flags=re.S)
    b = re.sub(r'</tr>', chr(10), b); b = re.sub(r'</t[dh]>', ' | ', b); b = re.sub(r'<[^>]+>', ' ', b); b = html.unescape(b)
    b = b.replace(chr(160), ' ')
    b = re.sub(r'[ \t]+', ' ', b); b = re.sub(chr(10) + r'\s*' + chr(10) + '+', chr(10), b)
    return b


_S = requests.Session()


def _get(url, params=None, retries=3):
    for a in range(retries):
        try:
            r = _S.get(url, params=params, timeout=90, headers=H)
            if r.status_code == 200:
                return r.text
        except Exception:
            pass
        time.sleep(4)
    return None


def measures(code, area='', simdate='20230701'):
    key = f'{code}|{area}|{simdate}'
    if key in _cache:
        return _cache[key]
    code10 = (code + '00') if len(code) == 8 else code
    page = _get('https://ec.europa.eu/taxation_customs/dds2/taric/measures.jsp',
                {'Lang': 'en', 'SimDate': simdate, 'Area': area, 'Taric': code10, 'LangDescr': 'en', 'search_text': 'goods',
                 'MeasType': '', 'StartPub': '', 'EndPub': '', 'MeasText': '', 'GoodsText': '', 'op': '', 'textSearch': '',
                 'OrderNum': '', 'Regulation': '', 'measStartDat': '', 'measEndDat': ''})
    if page is None or 'measures_details.jsp' not in page:
        return None
    urls = re.findall(r"expandCollapseIFrame\('taric_(\d{10})_\d+', '(measures_details\.jsp\?[^']+)'", page)
    descs = {}
    for m in re.finditer(r'(\d{4}) (\d{2}) (\d{2}) (\d{2})\s*</', page):
        gc = ''.join(m.groups())
        tail = page[m.end(): m.end() + 3000]
        dm = re.search(r'<td[^>]*class="[^"]*description[^"]*"[^>]*>(.*?)</td>', tail, flags=re.S)
        if dm:
            descs.setdefault(gc, re.sub(r'\s+', ' ', _clean(dm.group(1))).strip(' |'))
    seen, goods = set(), []
    for gcode, u in urls:
        if gcode in seen:
            continue
        seen.add(gcode)
        raw = _get('https://ec.europa.eu/taxation_customs/dds2/taric/' + html.unescape(u))
        g = {'goods_code': gcode, 'description': descs.get(gcode, '')[:200], 'third_country_duty': None, 'third_country_conditions': None, 'measures': []}
        if raw:
            # condition-dependent duties (entry-price system etc.): the duty is empty inline and lives in measures_conditions.jsp
            # positional scan: each conditions link belongs to the nearest preceding measure-type occurrence
            type_pos = sorted((mm.start(), mm.group(0)) for mm in re.finditer('|'.join(re.escape(t) for t in TYPES), raw))
            cond_sids = {}
            for mm in re.finditer(r'measures_conditions\.jsp\?MeasureSid=(\d+)', raw):
                prev = [tp for tp in type_pos if tp[0] < mm.start()]
                if prev:
                    cond_sids.setdefault(prev[-1][1], []).append(mm.group(1))
            for k in cond_sids:      # de-duplicate (show/hide links repeat the same sid)
                seen_s, uniq = set(), []
                for s_ in cond_sids[k]:
                    if s_ not in seen_s:
                        seen_s.add(s_); uniq.append(s_)
                cond_sids[k] = uniq
            txt = _clean(raw)
            cur_area = None
            for line in txt.split(chr(10)):
                m = re.search(r'(?:^|\|)\s*([A-Za-z][A-Za-z ,\-\.]*?) \(([A-Z][A-Za-z ]*?)(?: (\d{4}))?\)\s*\|', line)
                if m:
                    cur_area = m.group(1).strip()
                m2 = re.search(r'(' + '|'.join(TYPES) + r')[^:(]*\(([\d\-]+) - ([\d\-]*)\)\s*(:\s*\|\s*([^|]*)\|)?', line)
                if m2:
                    duty = (m2.group(5) or '').strip()
                    entry = {'type': m2.group(1), 'start': m2.group(2), 'end': m2.group(3), 'duty': duty, 'area': cur_area}
                    if not duty and cond_sids.get(m2.group(1)):
                        sid = cond_sids[m2.group(1)].pop(0)
                        # fetch condition tiers only for the erga-omnes duty and the Mercosur preference (each fetch ~4 s)
                        if (cur_area or '').startswith('ERGA') or (cur_area or '').lower().startswith('mercosur'):
                            entry['conditions'] = conditions(sid, simdate)
                    if m2.group(1) == 'Third country duty' and (cur_area or '').startswith('ERGA'):
                        g['third_country_duty'] = duty or None
                        g['third_country_conditions'] = entry.get('conditions')
                    g['measures'].append(entry)
        goods.append(g)
    out = {'code': code, 'area': area, 'simdate': simdate, 'goods': goods}
    _cache[key] = out
    json.dump(_cache, open(CACHE, 'w', encoding='utf-8'))
    return out


def conditions(sid, simdate):
    """Entry-price (or other) condition tiers for a measure: list of {threshold_eur_100kg, duty}."""
    t = _get(f'https://ec.europa.eu/taxation_customs/dds2/taric/measures_conditions.jsp?MeasureSid={sid}&Lang=en&LangDescr=&SimDate={simdate}')
    if not t:
        return None
    out = []
    for line in _clean(t).split(chr(10)):
        m = re.search(r'entry price \(see components\) ([\d\.]+) EUR / (100 kg|hl|tonne) \| Apply the amount of the action \(see components\) ([^|]+)\|', line)
        if m:
            out.append({'threshold': float(m.group(1)), 'unit': m.group(2), 'duty': m.group(3).strip()})
    return out or [{'raw': _clean(t)[:1500]}]


def duty_at_unit_value(g, uv_eur_100kg, cn8=None):
    """Third-country duty expression applicable at the Greek unit value: inline duty, or the entry-price tier met.
    Entry prices per hl are compared with the unit value per hl (density by heading)."""
    if g.get('third_country_duty'):
        return g['third_country_duty']
    conds = g.get('third_country_conditions') or []
    tiers = [c for c in conds if 'threshold' in c]
    if tiers and uv_eur_100kg:
        def uv_for(c):
            if c['unit'] == 'hl':
                return uv_hl(cn8 or g.get('goods_code', '')[:8], uv_eur_100kg)
            if c['unit'] == 'tonne':
                return uv_eur_100kg * 10
            return uv_eur_100kg
        met = [c for c in tiers if uv_for(c) is not None and uv_for(c) >= c['threshold']]
        if met:
            return max(met, key=lambda c: c['threshold'])['duty']
        return min(tiers, key=lambda c: c['threshold'])['duty']
    return None


def parse_duty(expr):
    e = expr.replace(',', '.')
    d = {'adval': None, 'specific': None, 'specific_unit': None, 'min': None, 'max': None, 'agri_component': False, 'raw': expr}
    m = re.search(r'(\d+(?:\.\d+)?) %', e)
    if m:
        d['adval'] = float(m.group(1)) / 100
    if re.search(r'\bEA\b|\bAD S/Z\b|\bAD F/M\b|agricultural component', e):
        d['agri_component'] = True
    m = re.search(r'MIN (\d+(?:\.\d+)?) EUR / ([\w ]+?)(?= MAX|$)', e)
    if m:
        d['min'] = float(m.group(1)); d['specific_unit'] = m.group(2).strip()
    m = re.search(r'MAX (\d+(?:\.\d+)?) EUR / ([\w ]+)$', e)
    if m:
        d['max'] = float(m.group(1)); d['specific_unit'] = m.group(2).strip()
    m = re.search(r'(?:^|\+ )(\d+(?:\.\d+)?) EUR / ([\w ]+?)(?= MIN| MAX|$)', e)
    if m:
        d['specific'] = float(m.group(1)); d['specific_unit'] = m.group(2).strip()
    if e.strip().lower().startswith('free'):
        d['adval'] = 0.0
    return d


def hl_per_100kg(cn8):
    """Hectolitres in 100 kg of product, by heading (density kg/l): ethanol 0.79; wine, vermouth, cider 0.99; spirits
    0.95; concentrated juices with Brix > 67 (CN 2009 xx 11 / xx 19) 1.3; other juices 1.05; otherwise water."""
    h4 = cn8[:4]
    if h4 == '2207':
        return 1 / 0.79
    if h4 in ('2204', '2205', '2206'):
        return 1 / 0.99
    if h4 == '2208':
        return 1 / 0.95
    if h4 == '2009':
        return 1 / 1.3 if cn8[6:8] in ('11', '19') else 1 / 1.05
    return 1.0


def uv_hl(cn8, uv_eur_per_100kg):
    """Unit value per hl from the COMEXT value per 100 kg."""
    return uv_eur_per_100kg / hl_per_100kg(cn8) if uv_eur_per_100kg else None


def ave(expr, uv_eur_per_100kg=None, uv_eur_per_hl=None):
    """Ad valorem equivalent at a unit value (EUR per 100 kg, or per hl). None if not computable."""
    d = parse_duty(expr)
    if d['agri_component']:
        return None
    if d['adval'] is None and d['specific'] is None and d['min'] is None and d['max'] is None:
        return None       # nothing parseable (e.g. '% vol/hl' spirits duties): let the caller fall back
    av = d['adval'] or 0.0
    unit = (d['specific_unit'] or '').lower().replace(' ', '')
    if '100kg' in unit:
        uv = uv_eur_per_100kg
    elif unit.startswith('hl'):
        uv = uv_eur_per_hl
    elif unit.startswith('tonne') or unit.startswith('1000kg') or unit == 't':
        uv = uv_eur_per_100kg * 10 if uv_eur_per_100kg else None
    elif unit.startswith('kg'):
        uv = uv_eur_per_100kg / 100 if uv_eur_per_100kg else None
    else:
        uv = None
    if d['specific'] is not None:
        if not uv:
            return None
        av += d['specific'] / uv
    if d['min'] is not None or d['max'] is not None:
        if not uv:
            return None
        duty = av * uv
        if d['min'] is not None:
            duty = max(duty, d['min'])
        if d['max'] is not None:
            duty = min(duty, d['max'])
        av = duty / uv
    return av


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    for code, area, date in [('23040000', '', '20230701'), ('09011100', '', '20230701'), ('24012085', '', '20230701'),
                             ('23040000', 'BR', '20260930'), ('24012085', 'BR', '20260930'), ('08055010', 'AR', '20260930'), ('20091998', 'BR', '20260930')]:
        m = measures(code, area, date)
        if m is None:
            print(code, area, date, 'FAILED'); continue
        for g in m['goods']:
            print(code, area or 'ERGA', date, g['goods_code'], g['description'][:40], '| TCD:', g['third_country_duty'],
                  '| other:', [(e['area'], e['type'][:18], e['duty'], e['start'][-4:]) for e in g['measures'] if e['type'] != 'Third country duty'][:3])
    print(parse_duty('18.40 % MIN 22.00 EUR / 100 kg MAX 24.00 EUR / 100 kg'), ave('18.40 % MIN 22.00 EUR / 100 kg MAX 24.00 EUR / 100 kg', 605))
    print(parse_duty('12.80 % + 304.10 EUR / 100 kg'), ave('12.80 % + 304.10 EUR / 100 kg', 345))
    print(parse_duty('10.20 EUR / hl'), ave('10.20 EUR / hl', uv_eur_per_hl=63))
    print(parse_duty('Free'), ave('Free'), parse_duty('0.00 %'), ave('0.00 %', 100))
