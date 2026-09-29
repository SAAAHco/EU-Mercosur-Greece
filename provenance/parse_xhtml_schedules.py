"""Parse EU (Appendix 2-A-1) and Mercosur (Appendix 2-A-2) tariff schedules
from the OJ L 2026/184 XHTML (Cellar manifestation .0006.03/DOC_1) into CSV."""
import csv, re, sys
from lxml import etree
SRC = 'OJ_L_202600184_iTA_EN.xhtml'
parser = etree.XMLParser(recover=True, huge_tree=True)
tree = etree.parse(SRC, parser)
root = tree.getroot()
ns = {'x': 'http://www.w3.org/1999/xhtml'}
def txt(el):
    t = ''.join(el.itertext())
    t = t.replace('\u00a0', ' ').replace('\u2014', '').replace('\u2013', '-')
    return ' '.join(t.split())
cur_title = None
out = {'2-A-1': [], '2-A-2': []}
headers = {}
for el in root.iter('{http://www.w3.org/1999/xhtml}p', '{http://www.w3.org/1999/xhtml}table'):
    tag = etree.QName(el).localname
    if tag == 'p' and 'oj-doc-ti' in (el.get('class') or ''):
        t = txt(el)
        if t.startswith('Appendix 2-A-1'): cur_title = '2-A-1'
        elif t.startswith('Appendix 2-A-2'): cur_title = '2-A-2'
        elif t.startswith('ANNEX') or t.startswith('Appendix'): cur_title = None
        continue
    if tag == 'table' and cur_title:
        if any(etree.QName(a).localname == 'td' for a in el.iterancestors()):
            continue  # skip tables nested inside a description cell
        # only direct rows of this table
        rows = el.findall('./x:tbody/x:tr', ns) or el.findall('./x:tr', ns)
        for tr in rows:
            cells = [txt(td) for td in tr.findall('./x:td', ns)]
            if not cells: continue
            if any(c in ('CN 2013', 'NCM', 'NCM 2012') for c in cells[:1]):
                headers.setdefault(cur_title, cells); continue
            out[cur_title].append(cells)
for k, rows in out.items():
    fn = f'appendix_{k}_{"EU" if k=="2-A-1" else "MERCOSUR"}_schedule_from_OJ_xhtml.csv'
    with open(fn, 'w', newline='', encoding='utf-8') as f:
        w = csv.writer(f)
        if k in headers: w.writerow(headers[k])
        for r in rows: w.writerow(r)
    lens = {}
    for r in rows: lens[len(r)] = lens.get(len(r), 0) + 1
    print(k, 'header:', headers.get(k), 'rows:', len(rows), 'cell-count dist:', lens)
