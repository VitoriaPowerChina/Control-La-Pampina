"""
Servidor do Dashboard de Construccion — La Pampina (PowerChina)
Porta: 8080  →  http://127.0.0.1:8080
"""
import json, re, warnings
from io import BytesIO
from datetime import date, datetime, timedelta
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse
warnings.filterwarnings("ignore")
import cgi, openpyxl

ROOT     = Path(__file__).resolve().parent
PORT     = 8080
XLSM     = ROOT / "Anexo 01-LPA.xlsm"
SHEET    = "LPA CONSTR."
DATA_DATE = datetime(2026, 5, 17)

# ── Column indices in LPA CONSTR. (0-based, row 10 = header) ──────────────────
C = dict(fila=0, nivel=1, resumen=2, edt=6, tarea=7, duracion=8,
         inicio=9, fin=10, hh=11, incidencia=13,
         pct_plan=15, pct_real=16, pct_comp_plan=17, pct_comp_real=18,
         desv_pond=19, desviacion=20)
PLAN_START, PLAN_END = 22, 90   # cols for PLAN weekly
REAL_START, REAL_END = 92, 160  # cols for REAL weekly
CURR_IDX = 28                   # index of current week (S20-26) in the 68-series
DATA_WEEK = "S20-26"


def parse_constr(source):
    if isinstance(source, Path):
        wb = openpyxl.load_workbook(str(source), read_only=True, data_only=True, keep_vba=True)
    else:
        wb = openpyxl.load_workbook(source, read_only=True, data_only=True, keep_vba=True)

    ws  = wb[SHEET]
    rows = list(ws.iter_rows(values_only=True))

    # Row 9 (index 8) = dates;  Row 10 (index 9) = headers;  Row 11+ = data
    date_row   = rows[8]
    header_row = rows[9]

    plan_weeks = [header_row[c] for c in range(PLAN_START, PLAN_END) if c < len(header_row)]
    plan_dates = [_d(date_row[c]) for c in range(PLAN_START, PLAN_END) if c < len(date_row)]

    def _v(row, col, default=None):
        return row[col] if col < len(row) else default

    records, scurve_plan, scurve_real = [], [], []
    total_row = None

    for row in rows[10:]:
        if _v(row, 0) is None and _v(row, 5) is None:
            continue
        edt = str(_v(row, C['edt']) or '').strip()
        tarea = str(_v(row, C['tarea']) or '').strip()
        if not tarea:
            continue

        ps = [(_v(row, c) or 0) for c in range(PLAN_START, PLAN_END)]
        rs = [(_v(row, c) or 0) for c in range(REAL_START, REAL_END)]

        rec = dict(
            edt       = edt,
            nivel     = int(_v(row, C['nivel']) or 0),
            resumen   = (_v(row, C['resumen']) == 'Si' or
                         _v(row, C['resumen']) == 'Sí' or
                         str(_v(row, C['resumen']) or '').lower() in ('si','sí','yes','true')),
            tarea     = tarea,
            duracion  = str(_v(row, C['duracion']) or ''),
            inicio    = _d(_v(row, C['inicio'])),
            fin       = _d(_v(row, C['fin'])),
            hh        = float(_v(row, C['hh']) or 0),
            incidencia= float(_v(row, C['incidencia']) or 0),
            pctPlan   = float(_v(row, C['pct_plan']) or 0),
            pctReal   = float(_v(row, C['pct_real']) or 0),
            pctCompPlan = float(_v(row, C['pct_comp_plan']) or 0),
            pctCompReal = float(_v(row, C['pct_comp_real']) or 0),
            desvPond  = float(_v(row, C['desv_pond']) or 0),
            desviacion= float(_v(row, C['desviacion']) or 0),
            planSeries= ps,
            realSeries= rs,
        )
        rec['status'] = _status(rec)
        records.append(rec)

        if edt == '4.5':
            total_row = rec
            scurve_plan = ps
            scurve_real = rs

    wb.close()

    # Build S-curve
    scurve = []
    for i in range(68):
        wk  = plan_weeks[i] if i < len(plan_weeks) else f'W{i}'
        dt  = plan_dates[i] if i < len(plan_dates) else None
        p   = scurve_plan[i] if i < len(scurve_plan) else 0
        r   = scurve_real[i] if i < len(scurve_real) else 0
        scurve.append(dict(week=wk, date=dt, plan=p, real=r if r else None,
                           isCurrent=(i==CURR_IDX)))

    # Areas (level-3)
    areas = [r for r in records if r['resumen'] and r['nivel'] in (3, 4)]

    # Leaves
    leaves = [r for r in records if not r['resumen']]

    top_dev = sorted([r for r in leaves if r['incidencia'] > 0.0001],
                     key=lambda r: r['desviacion'])[:50]

    critical = sorted([r for r in leaves
                       if r['incidencia'] > 0.003 and r['desviacion'] < -0.05],
                      key=lambda r: r['incidencia'], reverse=True)

    sin_avance = sorted([r for r in leaves
                         if r['pctCompPlan'] > 0 and r['pctCompReal'] == 0],
                        key=lambda r: r['incidencia'], reverse=True)

    ranking = sorted([r for r in leaves if r['incidencia'] > 0],
                     key=lambda r: abs(r['desvPond']), reverse=True)[:50]

    future = sorted([r for r in leaves
                     if r['pctCompReal'] < 1.0 and r['fin'] and r['fin'] > DATA_DATE.date().isoformat()
                     and r['incidencia'] > 0.0005],
                    key=lambda r: r['incidencia'] * (1 - r['pctCompReal']), reverse=True)[:30]

    avg_incid = (sum(r['incidencia'] for r in leaves if r['incidencia'] > 0) /
                 max(1, sum(1 for r in leaves if r['incidencia'] > 0)))

    tr = total_row or {}
    return dict(
        meta = dict(
            dataDate  = DATA_DATE.date().isoformat(),
            dataWeek  = DATA_WEEK,
            startLB   = tr.get('inicio'),
            endLB     = tr.get('fin'),
            totalHH   = tr.get('hh', 0),
            pctPlan   = tr.get('pctPlan', 0),
            pctReal   = tr.get('pctReal', 0),
            desvio    = tr.get('desvPond', 0),
            actTotal  = len(leaves),
            actSinAvance = len(sin_avance),
            avgIncidencia = avg_incid,
        ),
        scurve    = scurve,
        areas     = areas,
        topDesvios= top_dev,
        critical  = critical,
        sinAvance = sin_avance,
        ranking   = ranking,
        future    = future,
        allLeaves = leaves,
    )


def _d(v):
    if isinstance(v, datetime): return v.date().isoformat()
    if isinstance(v, date):     return v.isoformat()
    return None

def _status(r):
    if r['pctCompReal'] >= 0.995:
        return 'completed'
    if r['pctCompReal'] > 0:
        return 'inProgress'
    if r['pctCompPlan'] > 0:
        return 'late'
    return 'notStarted'


# ── Pre-load on startup ────────────────────────────────────────────────────────
_cache = None
def get_data():
    global _cache
    if _cache is None and XLSM.exists():
        try:
            _cache = parse_constr(XLSM)
        except Exception as e:
            _cache = {"error": str(e)}
    return _cache or {"error": "Arquivo nao encontrado"}


class Handler(SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(ROOT), **kw)

    def log_message(self, fmt, *args):
        pass  # silence logs

    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/api/data':
            self.send_json(get_data())
        else:
            super().do_GET()

    def do_POST(self):
        if urlparse(self.path).path != '/api/import':
            self.send_error(404); return
        global _cache
        try:
            form = cgi.FieldStorage(
                fp=self.rfile, headers=self.headers,
                environ={"REQUEST_METHOD":"POST",
                         "CONTENT_TYPE": self.headers.get("Content-Type",""),
                         "CONTENT_LENGTH": self.headers.get("Content-Length","0")})
            up = form.get("file")
            if up is None or not getattr(up, 'filename', None):
                raise ValueError("Nenhum arquivo enviado.")
            _cache = parse_constr(BytesIO(up.file.read()))
            self.send_json({"ok": True, "rows": len(_cache.get('allLeaves', []))})
        except Exception as e:
            self.send_json({"error": str(e)}, 400)

    def send_json(self, payload, status=200):
        data = json.dumps(payload, ensure_ascii=False, default=str).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.send_header('Content-Length', str(len(data)))
        self.send_header('Access-Control-Allow-Origin', '*')
        self.end_headers()
        self.wfile.write(data)


if __name__ == '__main__':
    print("Carregando dados da planilha...")
    get_data()
    srv = ThreadingHTTPServer(('127.0.0.1', PORT), Handler)
    print(f"Dashboard disponivel em  http://127.0.0.1:{PORT}")
    srv.serve_forever()
