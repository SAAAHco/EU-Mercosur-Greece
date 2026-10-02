"""WITS / UNCTAD TRAINS tariff API: MFN applied rates (reported) and ad valorem equivalents (aveestimated) at HS6."""
import requests, json, time, os, sys, io
H = {'User-Agent': 'Mozilla/5.0'}
CACHE = 'wits_cache.json'
_cache = json.load(open(CACHE)) if os.path.exists(CACHE) else {}


def fetch(reporter, products, year=2023, datatype='reported', partner='000', retries=3):
    """reporter: '918' EU, '076' BRA, '032' ARG, '858' URY, '600' PRY. products: list of HS6 strings (<=25 per call)."""
    out = {}
    todo = [p for p in products if f'{reporter}|{p}|{year}|{datatype}' not in _cache]
    for p in products:
        k = f'{reporter}|{p}|{year}|{datatype}'
        if k in _cache:
            out[p] = _cache[k]
    for i in range(0, len(todo), 20):
        chunk = todo[i:i + 20]
        u = f'https://wits.worldbank.org/API/V1/SDMX/V21/datasource/TRN/reporter/{reporter}/partner/{partner}/product/{";".join(chunk)}/year/{year}/datatype/{datatype}?format=JSON'
        for a in range(retries):
            try:
                r = requests.get(u, timeout=400, headers=H)
                if r.status_code == 200 and 'dataSets' in r.text:
                    j = r.json()
                    dims = j['structure']['dimensions']['series']
                    prod_vals = [d for d in dims if d['id'] == 'PRODUCTCODE'][0]['values']
                    pidx = 2  # empirically the product index sits at position 2 of the series key
                    attrs = [a_['id'] for a_ in j['structure']['attributes']['observation']]
                    got = set()
                    for key, ser in j['dataSets'][0]['series'].items():
                        p = prod_vals[int(key.split(':')[pidx])]['id']
                        obs = ser['observations']['0']
                        rec = {'rate': obs[0], 'attrs': dict(zip(attrs, obs[1:]))}
                        _cache[f'{reporter}|{p}|{year}|{datatype}'] = rec; out[p] = rec; got.add(p)
                    for p in chunk:
                        if p not in got:
                            _cache[f'{reporter}|{p}|{year}|{datatype}'] = None; out[p] = None
                    break
                elif r.status_code == 404:
                    for p in chunk:
                        _cache[f'{reporter}|{p}|{year}|{datatype}'] = None; out[p] = None
                    break
                else:
                    time.sleep(5)
            except Exception as e:
                time.sleep(10)
        json.dump(_cache, open(CACHE, 'w'))
    return out


if __name__ == '__main__':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    print(fetch('918', ['230400', '090111', '240120', '120190'], datatype='aveestimated'))
    print(fetch('076', ['150910', '200570', '080620'], datatype='reported'))
