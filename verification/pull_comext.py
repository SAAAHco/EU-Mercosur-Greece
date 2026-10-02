import requests, json, time, os, sys, io, gzip
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
s = requests.Session(); s.mount('https://', HTTPAdapter(max_retries=Retry(total=4, backoff_factor=3, status_forcelist=[429,500,502,503,504])))
H = {'User-Agent': 'Mozilla/5.0', 'Accept-Encoding': 'gzip, deflate'}
BASE = 'https://ec.europa.eu/eurostat/api/comext/dissemination/statistics/1.0/data/ds-045409'
PARTNERS = ['BR', 'AR', 'UY', 'PY', 'BO']
FLOWS = {'1': 'M', '2': 'X'}
log = []
for year in range(2014, 2025):
    for p in PARTNERS:
        for fl, fname in FLOWS.items():
            out = f'comext_raw/gr_{fname}_{year}_{p}.json.gz'
            if os.path.exists(out):
                continue
            params = [('reporter', 'GR'), ('partner', p), ('flow', fl), ('freq', 'A'), ('time_period', str(year)), ('format', 'JSON'),
                      ('indicators', 'VALUE_IN_EUROS'), ('indicators', 'QUANTITY_IN_100KG')]
            for attempt in range(4):
                try:
                    t = time.time(); r = s.get(BASE, params=params, timeout=600, headers=H)
                    if r.status_code != 200:
                        print(year, p, fname, 'HTTP', r.status_code, r.text[:200]); time.sleep(10); continue
                    j = r.json()
                    # keep only non-null value cells + dimension index to shrink storage
                    with gzip.open(out, 'wt', encoding='utf-8') as f:
                        json.dump(j, f)
                    nv = len(j.get('value', {}))
                    print(year, p, fname, 'ok', nv, 'cells', round(time.time() - t, 1), 's', flush=True)
                    log.append((year, p, fname, nv))
                    break
                except Exception as e:
                    print(year, p, fname, 'ERR', type(e).__name__, str(e)[:150], flush=True); time.sleep(15)
json.dump(log, open('comext_raw/_pull_log.json', 'w'))
print('DONE', len(log))
