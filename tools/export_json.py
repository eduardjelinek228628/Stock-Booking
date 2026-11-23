"""Export an .xlsx into the raw-range JSON the viewer's data sources return, for tests/real_structure.test.js.
Output contains real data: keep it in data/ (git-ignored). Needs: pip install openpyxl"""
import openpyxl, json, sys, datetime
from openpyxl.utils import range_boundaries, get_column_letter
src, out = sys.argv[1], sys.argv[2]
wv = openpyxl.load_workbook(src, data_only=True)
wf = openpyxl.load_workbook(src, data_only=False)
EPOCH = datetime.datetime(1899, 12, 30)
def conv(v):
    if isinstance(v, datetime.datetime): return (v - EPOCH).days + (v - EPOCH).seconds / 86400
    if isinstance(v, datetime.date): return (datetime.datetime(v.year, v.month, v.day) - EPOCH).days
    if v is None: return ''
    if isinstance(v, (int, float, str, bool)): return v
    return str(v)
def block(wsv, wsf, ref):
    c1, r1, c2, r2 = range_boundaries(ref)
    V, T, N, F = [], [], [], []
    for r in range(r1, r2 + 1):
        vr, tr, nr, fr = [], [], [], []
        for c in range(c1, c2 + 1):
            cv = wsv.cell(r, c); cf = wsf.cell(r, c)
            v = conv(cv.value); nf = cv.number_format or 'General'
            if isinstance(cv.value, (datetime.datetime, datetime.date)): t = f"{cv.value.day}.{cv.value.month}.{cv.value.year}"; nf = 'd.m.yyyy' if 'd' not in nf.lower() else nf
            else: t = '' if v == '' else str(v)
            f = cf.value
            f = f if isinstance(f, str) and f.startswith('=') else ('=ARRAY()' if type(f).__name__ == 'ArrayFormula' else v)
            vr.append(v); tr.append(t); nr.append(nf); fr.append(f)
        V.append(vr); T.append(tr); N.append(nr); F.append(fr)
    q = wsv.title if wsv.title.replace('_','').isalnum() else "'" + wsv.title.replace("'", "''") + "'"
    return {"address": f"{q}!{ref}", "values": V, "text": T, "numberFormat": N, "formulas": F}
book = {"sheets": []}
for i, wsv in enumerate(wv.worksheets):
    wsf = wf[wsv.title]
    s = {"id": "{s%d}" % i, "name": wsv.title, "tables": [], "used": None}
    for t in wsv.tables.values():
        s["tables"].append({"name": t.name, "range": block(wsv, wsf, t.ref)})
    dim = wsv.calculate_dimension()
    if dim and ':' in dim: s["used"] = block(wsv, wsf, dim)
    book["sheets"].append(s)
json.dump(book, open(out, 'w'), ensure_ascii=False)
print("ok", out)
