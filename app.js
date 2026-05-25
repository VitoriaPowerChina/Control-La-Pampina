/* Dashboard Construccion — La Pampina | PowerChina | SheetJS */
'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let D = null;
let charts = {};
const collapsedNodes    = new Set();
let _wbsFilterEdt     = ''; // EDT selecionado na cascata WBS; '' = sem filtro
let _scurveFilterEdt  = ''; // EDT selecionado no filtro da Curva S filtrada
let _consCache = null;
let _consolTree = null;
let _simRows    = new Map(); // edt → delta (number: PBs or percentage points)  — Escenarios tab
let _simTabRows = [];        // { edt, delta, mode:'pb'|'pct' }                   — Simulador tab
let _simTabMode = 'pb';      // current add-form mode in Simulador tab

// ── Boot ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupTheme();
  setupTabs();
  setupFileUpload();
  setupSimTab();
  setupArbol();
  setupConsolidado();
  setupTableFilters();
  setupTabFilters();
});

// ── Theme toggle ─────────────────────────────────────────────────────────────
function setupTheme() {
  const btn   = document.getElementById('themeToggle');
  const saved = localStorage.getItem('theme') || 'light';
  applyTheme(saved);
  btn?.addEventListener('click', () => {
    const next = document.documentElement.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    localStorage.setItem('theme', next);
  });
}
function applyTheme(t) {
  document.documentElement.setAttribute('data-theme', t);
  const btn = document.getElementById('themeToggle');
  if (btn) btn.textContent = t === 'dark' ? '☀️' : '🌙';
}

// ── File loading (SheetJS) ────────────────────────────────────────────────────
function setupFileUpload() {
  on('fileInput', 'change', handleFile);
}

async function handleFile(e) {
  const file = e.target.files[0];
  if (!file) return;
  showToast('Leyendo archivo…', false, 60000);
  try {
    const buf = await file.arrayBuffer();
    // Yield ao browser para mostrar o toast antes do parsing pesado
    await new Promise(r => setTimeout(r, 30));
    showToast('Procesando datos…', false, 60000);
    await new Promise(r => setTimeout(r, 30));
    const wb = XLSX.read(buf, {
      type: 'array',
      cellDates: false,         // datas como serial numérico — mais rápido
      bookVBA: false,           // ignora código VBA (maior ganho em .xlsm)
      sheets: 'LPA CONSTR.',    // só lê a aba necessária
    });
    D = parseConstr(wb);
    document.getElementById('welcomeCard').style.display    = 'none';
    document.getElementById('resumenContent').style.display = 'block';
    populateAreaDropdowns();
    render();
    showToast('Cargado: ' + D.meta.dataWeek + ' · ' + D.meta.actTotal + ' actividades');
  } catch(err) {
    showToast('Error: ' + err.message, true);
    console.error(err);
  }
  e.target.value = '';
}

// ── Excel Parsing ─────────────────────────────────────────────────────────────
function parseConstr(wb) {
  const ws = wb.Sheets['LPA CONSTR.'];
  if (!ws) throw new Error('Aba "LPA CONSTR." não encontrada no arquivo.');

  const rows = XLSX.utils.sheet_to_json(ws, {
    header: 1, raw: true, cellDates: true, defval: null
  });

  const PLAN_START = 22, PLAN_END = 90;
  const REAL_START = 92, REAL_END = 160;

  const dateRow   = rows[8] || [];
  const headerRow = rows[9] || [];
  const planWeeks = headerRow.slice(PLAN_START, PLAN_END);
  const planDates = dateRow.slice(PLAN_START, PLAN_END).map(xlsxDateToIso);

  const C = {
    nivel:1, resumen:2, edt:6, tarea:7, duracion:8,
    inicio:9, fin:10, hh:11, incidencia:13,
    pct_plan:15, pct_real:16, pct_comp_plan:17, pct_comp_real:18,
    desv_pond:19, desviacion:20,
  };

  const records = [];
  let totalRow = null;

  for (const row of rows.slice(10)) {
    if (!row || (row[0] == null && row[5] == null)) continue;
    const edt   = String(row[C.edt]   ?? '').trim();
    const tarea = String(row[C.tarea] ?? '').trim();
    if (!tarea) continue;

    const rec = {
      edt, tarea,
      nivel:        +(row[C.nivel]         ?? 0),
      resumen:      isResumenVal(row[C.resumen]),
      duracion:     String(row[C.duracion] ?? ''),
      inicio:       xlsxDateToIso(row[C.inicio]),
      fin:          xlsxDateToIso(row[C.fin]),
      hh:           +(row[C.hh]            ?? 0),
      incidencia:   +(row[C.incidencia]    ?? 0),
      pctPlan:      +(row[C.pct_plan]      ?? 0),
      pctReal:      +(row[C.pct_real]      ?? 0),
      pctCompPlan:  +(row[C.pct_comp_plan] ?? 0),
      pctCompReal:  +(row[C.pct_comp_real] ?? 0),
      desvPond:     +(row[C.desv_pond]     ?? 0),
      desviacion:   +(row[C.desviacion]    ?? 0),
      planSeries:   row.slice(PLAN_START, PLAN_END).map(v => +v || 0),
      realSeries:   row.slice(REAL_START, REAL_END).map(v => +v || 0),
    };
    rec.status = calcStatus(rec);
    records.push(rec);
    if (edt === '4.5') totalRow = rec;
  }

  // Auto-detect semana actual: último índice com real > 0 na linha 4.5
  let currIdx = 0;
  if (totalRow) {
    for (let i = totalRow.realSeries.length - 1; i >= 0; i--) {
      if (totalRow.realSeries[i] > 0) { currIdx = i; break; }
    }
  }
  const dataWeek = planWeeks[currIdx] || `W${currIdx}`;
  const dataDate = planDates[currIdx] || null;

  const scurve = planWeeks.map((wk, i) => ({
    week: wk || `W${i}`,
    date: planDates[i],
    plan: totalRow ? (totalRow.planSeries[i] || 0) : 0,
    real: (totalRow && totalRow.realSeries[i] > 0) ? totalRow.realSeries[i] : null,
    isCurrent: i === currIdx,
  }));

  const areas     = records.filter(r => r.resumen && (r.nivel === 3 || r.nivel === 4));
  const leaves    = records.filter(r => !r.resumen);

  const topDesvios = leaves.filter(r => r.incidencia > 0.0001)
    .sort((a, b) => a.desviacion - b.desviacion).slice(0, 50);
  const critical  = leaves.filter(r => r.incidencia > 0.003 && r.desviacion < -0.05)
    .sort((a, b) => b.incidencia - a.incidencia);
  const sinAvance = leaves.filter(r => r.pctCompPlan > 0 && r.pctCompReal === 0)
    .sort((a, b) => b.incidencia - a.incidencia);
  const ranking   = leaves.filter(r => r.incidencia > 0)
    .sort((a, b) => Math.abs(b.desvPond) - Math.abs(a.desvPond)).slice(0, 50);
  const future    = leaves
    .filter(r => r.pctCompReal < 1 && r.fin && r.fin > dataDate && r.incidencia > 0.0005)
    .sort((a, b) => (b.incidencia*(1-b.pctCompReal)) - (a.incidencia*(1-a.pctCompReal)))
    .slice(0, 30);

  const leavesW = leaves.filter(r => r.incidencia > 0);
  const avgIncidencia = leavesW.length
    ? leavesW.reduce((s, r) => s + r.incidencia, 0) / leavesW.length : 0;

  const tr = totalRow || {};
  return {
    meta: {
      dataDate, dataWeek,
      startLB:      tr.inicio      ?? null,
      endLB:        tr.fin         ?? null,
      totalHH:      tr.hh          || 0,
      pctPlan:      tr.pctPlan     || 0,
      pctReal:      tr.pctReal     || 0,
      desvio:       tr.desvPond    || 0,
      actTotal:     leaves.length,
      actSinAvance: sinAvance.length,
      avgIncidencia,
    },
    scurve, areas, topDesvios, critical, sinAvance, ranking, future, allRecords: records, allLeaves: leaves,
  };
}

function xlsxDateToIso(v) {
  if (v == null) return null;
  // Serial numérico do Excel (ex: 45794 = 17/05/2026)
  if (typeof v === 'number' && v > 1) {
    const d = new Date((v - 25569) * 86400000);
    return d.toISOString().split('T')[0];
  }
  if (v instanceof Date && !isNaN(v)) return v.toISOString().split('T')[0];
  return null;
}
function isResumenVal(v) {
  if (v == null) return false;
  const s = String(v).toLowerCase().trim();
  return s === 'si' || s === 'sí' || s === 'yes' || s === 'true' || v === true || v === 1;
}
function calcStatus(r) {
  if (r.pctCompReal >= 0.995) return 'completed';
  if (r.pctCompReal > 0)      return 'inProgress';
  if (r.pctCompPlan > 0)      return 'late';
  return 'notStarted';
}

// ── Render ────────────────────────────────────────────────────────────────────
function render() {
  renderKPIs();
  renderResumen();
  renderAreaChart();
  renderStatusDonut();
  renderWeeklyBar();
  renderTop5();
  renderScurve();
  renderAreasBar();
  renderAreasTable();
  renderDesviosBar(D.topDesvios);
  renderDesviosTable(D.topDesvios);
  renderCriticasBar(D.critical);
  renderCriticasTable(D.critical);
  renderSinAvanceCharts(D.sinAvance);
  renderSinAvanceTable(D.sinAvance);
  renderRankingBar(D.ranking.slice(0, 20));
  renderRankingTable(D.ranking);
  renderConsolidado();             // sets _consCache
  _consolTree = buildConsolTree(); // builds virtual tree using _consCache
  buildCascadeFilters();
  buildScurveCascadeFilters();
  renderScurveFiltered();
  initSimTab();
  initArbol();
  renderArbol();
  renderFuture();
  renderTabla();
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
function renderKPIs() {
  const m = D.meta;
  set('kPlan',    pct(m.pctPlan));
  set('kReal',    pct(m.pctReal));
  set('kDev',     (m.desvio >= 0 ? '+' : '') + pct(m.desvio));
  set('kHH',      (m.totalHH / 1000).toFixed(0) + 'K h');
  set('kAct',     m.actTotal.toLocaleString());
  set('kSin',     m.actSinAvance);
  set('metaDate', 'Semana ' + m.dataWeek + '  |  ' + fmtDate(m.dataDate));

  const devEl = document.getElementById('kpi-dev');
  devEl.classList.toggle('kpi-alert', m.desvio < 0);
  devEl.classList.toggle('positive',  m.desvio >= 0);

  const realPts = D.scurve.filter(s => s.real != null && s.real > 0);
  let forecast = '—';
  if (realPts.length >= 2) {
    const last = realPts[realPts.length - 1];
    const prev = realPts[realPts.length - 2];
    const rate = last.real - prev.real;
    if (rate > 0) {
      const weeks = Math.round((1 - last.real) / rate);
      const d = new Date(last.date);
      d.setDate(d.getDate() + weeks * 7);
      forecast = d.toLocaleDateString('es-CL', { month: 'short', year: 'numeric' });
    }
  }
  set('kForecast', forecast);
}

// ── Resumen ejecutivo ─────────────────────────────────────────────────────────
function renderResumen() {
  const m = D.meta;
  const dev = m.desvio;
  const devStr = (dev >= 0 ? '+' : '') + pct(dev);
  const status = dev < -0.02 ? '<span class="alert">PROYECTO EN RETRASO CRÍTICO</span>'
               : dev < 0    ? '<span class="alert">Proyecto con retraso</span>'
               :               '<span class="ok">Proyecto sin atraso</span>';

  const areas = D.areas.filter(a => a.nivel === 3 && a.edt !== '4.5.1' && a.edt !== '4.5.9');
  const areaRows = areas.map(a => {
    const d = a.desvPond;
    const cls = d < -0.005 ? 'alert' : d < 0 ? '' : 'ok';
    return `<span class="${cls}">• ${a.tarea.trim()}: Plan ${pct(a.pctPlan)} / Real ${pct(a.pctReal)} / Desvío ${(d >= 0 ? '+' : '') + pct(d)}</span>`;
  }).join('<br>');

  document.getElementById('resumenText').innerHTML = `
    <strong>Proyecto:</strong> LA PAMPINA &nbsp;|&nbsp; <strong>Disciplina:</strong> CONSTRUCCIÓN<br>
    <strong>Fecha de control:</strong> ${fmtDate(m.dataDate)} (${m.dataWeek})<br>
    <strong>Período LB:</strong> ${fmtDate(m.startLB)} → ${fmtDate(m.endLB)}<br>
    <strong>H-H Totales:</strong> ${Math.round(m.totalHH).toLocaleString()} horas-hombre<br>
    <br>
    <strong>Situación:</strong> ${status}<br>
    Avance real: <strong>${pct(m.pctReal)}</strong> vs planeado: <strong>${pct(m.pctPlan)}</strong><br>
    Desvío acumulado ponderado: <strong>${devStr}</strong><br>
    <br>
    <strong>Avance por área:</strong><br>
    ${areaRows}<br>
    <br>
    <strong>Actividades sin avance real:</strong> ${m.actSinAvance}<br>
    <strong>Top desvío:</strong> ${D.topDesvios[0] ? D.topDesvios[0].tarea.trim() + ' (' + pct(D.topDesvios[0].desviacion) + ')' : '—'}
  `;
}

// ── Status Donut ──────────────────────────────────────────────────────────────
function renderStatusDonut() {
  const counts = { completed:0, inProgress:0, late:0, notStarted:0 };
  D.allLeaves.forEach(r => { counts[r.status] = (counts[r.status] || 0) + 1; });
  const total = D.allLeaves.length;

  destroyChart('statusDonut');
  charts['statusDonut'] = new Chart(document.getElementById('statusDonut'), {
    type: 'doughnut',
    data: {
      labels: ['Completada', 'En progreso', 'Atrasada', 'No iniciada'],
      datasets: [{
        data: [counts.completed, counts.inProgress, counts.late, counts.notStarted],
        backgroundColor: ['#00844a','#f0a500','#c00000','#a0a8c0'],
        borderWidth: 2, borderColor: '#fff'
      }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { position: 'right', labels: { font:{ size:11 }, padding:8 } },
        tooltip: { callbacks: { label: ctx =>
          `${ctx.label}: ${ctx.parsed} (${((ctx.parsed/total)*100).toFixed(1)}%)`
        }}
      }
    }
  });
}

// ── Weekly Bar ────────────────────────────────────────────────────────────────
function renderWeeklyBar() {
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  const slice = sc.slice(Math.max(0, currIdx - 9), currIdx + 1);

  destroyChart('weeklyBar');
  charts['weeklyBar'] = new Chart(document.getElementById('weeklyBar'), {
    type: 'bar',
    data: {
      labels: slice.map(s => s.week),
      datasets: [
        { label:'% Plan', data: slice.map(s => +(s.plan*100).toFixed(2)), backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real', data: slice.map(s => s.real != null ? +(s.real*100).toFixed(2) : null), backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top', labels:{ font:{size:11} } } },
      scales: { y: { ticks:{ callback: v => v+'%', font:{size:10} } } }
    }
  });
}

// ── Top 5 ─────────────────────────────────────────────────────────────────────
function renderTop5() {
  document.getElementById('top5List').innerHTML = D.topDesvios.slice(0, 5).map((r, i) => `
    <div class="top5-item">
      <span class="top5-rank">${i+1}</span>
      <div class="top5-info">
        <div class="top5-name">${r.tarea.trim().length > 44 ? r.tarea.trim().slice(0,44)+'…' : r.tarea.trim()}</div>
        <div class="top5-meta">${r.edt} · incid. ${pct(r.incidencia,3)}</div>
      </div>
      <span class="top5-dev dev-neg">${signPct(r.desviacion)}</span>
    </div>
  `).join('');
}

// ── Area bar chart ────────────────────────────────────────────────────────────
function renderAreaChart() {
  const areas = D.areas.filter(a => a.nivel===3 && a.incidencia > 0.001);
  const plan  = areas.map(a => +(a.pctPlan*100).toFixed(2));
  const real  = areas.map(a => +(a.pctReal*100).toFixed(2));

  destroyChart('areaChart');
  charts['areaChart'] = new Chart(document.getElementById('areaChart'), {
    type: 'bar',
    data: {
      labels: areas.map(a => a.tarea.trim()),
      datasets: [
        { label:'% Plan', data:plan, backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real', data:real, backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top' } },
      scales: { y: { ticks:{ callback: v => v+'%' }, max: Math.max(...plan,...real,5)*1.2 } }
    }
  });
}

// ── S-Curve shared chart config ───────────────────────────────────────────────
// planData and realData are arrays of % values (0-100). realData may contain nulls.
function _scurveChartConfig(labels, planData, realData) {
  return {
    type: 'line',
    data: {
      labels,
      datasets: [
        {
          label: '% Planeado',
          data: planData,
          borderColor: '#0054a6',
          backgroundColor: 'rgba(0,84,166,.07)',
          pointRadius: 0, fill: true, tension: 0.3, borderWidth: 2,
          yAxisID: 'y'
        },
        {
          label: '% Real',
          data: realData,
          borderColor: '#00a36c',
          backgroundColor: 'transparent',
          pointRadius: 1,
          // Fill area between Real and Plan: green when ahead, red when behind
          fill: { target: 0, above: 'rgba(0,163,108,.18)', below: 'rgba(192,0,0,.18)' },
          tension: 0.3, borderWidth: 2.5, spanGaps: false,
          yAxisID: 'y'
        },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top' },
        tooltip: {
          callbacks: {
            label: ctx => {
              const v = ctx.parsed.y ?? 0;
              return ctx.dataset.label + ': ' + v.toFixed(2) + '%';
            },
            footer: items => {
              if (!items.length) return [];
              const idx = items[0].dataIndex;
              const p = planData[idx];
              const r = realData[idx];
              // Only show when both plan and real exist at this point
              if (p == null || r == null) return [];
              const dev = +(r - p).toFixed(2);
              return ['Desvío: ' + (dev >= 0 ? '+' : '') + dev.toFixed(2) + '%'];
            }
          },
          footerColor: '#d97706',
          footerFont: { weight: '600' }
        }
      },
      scales: {
        y: { ticks: { callback: v => v + '%' }, min: 0, max: 100 }
      }
    }
  };
}

// ── S-Curve ───────────────────────────────────────────────────────────────────
function renderScurve() {
  const sc = D.scurve;
  const slice = sc;

  const planData = slice.map(s => +(s.plan*100).toFixed(3));
  const realData = slice.map(s => s.real != null ? +(s.real*100).toFixed(3) : null);

  destroyChart('scurveChart');
  charts['scurveChart'] = new Chart(
    document.getElementById('scurveChart'),
    _scurveChartConfig(slice.map(s => s.week), planData, realData)
  );
}

// ── S-Curve filtered — leaf resolver ─────────────────────────────────────────
// For virtual PV discipline nodes (EDT = "4.5.5.2.1.{discSeg}"), a simple prefix
// filter on D.allLeaves only captures PB-1 (the display template).  We must
// instead pull every raw leaf that belongs to that discipline across ALL PBs.
// For BESS disc nodes the prefix filter is fine (BESS is disc-first, so all PBs
// live under the disc EDT).  Non-zone nodes also use the default prefix filter.
function _scurveLeavesForEdt(filtEdt) {
  if (!filtEdt) return D.allLeaves;

  // Check if this EDT falls inside a consolidation zone.
  // We work directly with D.allLeaves (no dependency on _consCache.leaves which
  // is not persisted by buildConsolidated).
  const parts = filtEdt.split('.');
  for (const z of CONS_ZONES) {
    const zp = z.prefix.split('.');
    if (!zp.every((s, i) => parts[i] === s) || parts.length <= zp.length) continue;

    // filtEdt is inside this zone.
    if (parts.length <= z.discIdx) {
      // Zone root — return every raw leaf in the zone
      return D.allLeaves.filter(l => l.edt.startsWith(z.prefix + '.'));
    }

    const discSeg = parts[z.discIdx];
    if (!DISC_LABELS[discSeg]) break; // unknown disc — fall through to default

    const minDepthForTask = Math.max(z.pbIdx, z.discIdx) + 1;
    if (parts.length > minDepthForTask) {
      // Task-level node: match raw leaves using the wildcard key pattern.
      // Build a pattern with '*' at the PB position and exact matches elsewhere.
      const pattern = [...parts];
      pattern[z.pbIdx] = '*'; // '*' = match any PB
      return D.allLeaves.filter(leaf => {
        const lp = leaf.edt.split('.');
        return lp.length === pattern.length &&
               pattern.every((seg, i) => seg === '*' || seg === lp[i]);
      });
    }

    // Disc-level node — all raw leaves for this zone + discipline
    return D.allLeaves.filter(leaf => {
      const lp = leaf.edt.split('.');
      return zp.every((s, i) => lp[i] === s) && lp[z.discIdx] === discSeg;
    });
  }

  // Default: filter D.allLeaves by EDT prefix
  return D.allLeaves.filter(r => r.edt === filtEdt || r.edt.startsWith(filtEdt + '.'));
}

// ── S-Curve filtered ─────────────────────────────────────────────────────────
function renderScurveFiltered() {
  if (!D) return;

  const filtEdt  = _scurveFilterEdt;
  const leaves   = _scurveLeavesForEdt(filtEdt);
  const totalInc = leaves.reduce((s, r) => s + (r.incidencia || 0), 0);

  // Label in subtitle — prefer _consolTree name (shows virtual disc labels too)
  const filtLabel = filtEdt
    ? (_consolTree?.find(r => r.edt === filtEdt)?.tarea?.trim()
        || D.allRecords.find(r => r.edt === filtEdt)?.tarea?.trim()
        || filtEdt)
    : 'Total proyecto';
  const labelEl = document.getElementById('scurveFilterLabel');
  if (labelEl) labelEl.textContent = 'Área: ' + filtLabel;

  if (!totalInc || !leaves.length) {
    destroyChart('scurveFilteredChart');
    return;
  }

  const nWeeks  = D.scurve.length;
  const currIdx = D.scurve.findIndex(s => s.isCurrent);

  // ── Per-leaf normalisation (weeks ≤ currIdx) ─────────────────────────────────
  // planSeries[i] values are anchored so that series[currIdx] → pctCompPlan,
  // matching exactly the WBS-grid value. Real is treated identically up to currIdx.
  const planArr = new Array(nWeeks).fill(0);
  const realArr = new Array(nWeeks).fill(0);

  // ── Semanas passadas + atual: normalização por folha ─────────────────────────
  // Folhas individuais só têm dados confiáveis até currIdx (após ficam flat).
  // A normalização ps/pCurr*pctCompPlan ancora exatamente no valor WBS.
  for (const leaf of leaves) {
    const w     = leaf.incidencia || 0;
    const pCurr = leaf.planSeries[currIdx] || 0;
    const rCurr = leaf.realSeries[currIdx]  || 0;

    for (let i = 0; i <= currIdx; i++) {
      const ps = leaf.planSeries[i] || 0;
      planArr[i] += pCurr > 0 ? w * ps / pCurr * leaf.pctCompPlan : 0;

      const rs = leaf.realSeries[i] || 0;
      realArr[i] += w * (
        rCurr > 0 ? rs / rCurr * leaf.pctCompReal
        : i === currIdx ? leaf.pctCompReal
        : 0
      );
    }
  }

  // ── Semanas futuras: forma da linha resumo, ancorada em currIdx → 100% ───────
  // A linha resumo (ou a linha total 4.5 como fallback) tem o planSeries completo
  // com os dados futuros. Interpolamos do valor WBS atual até totalInc (= 100%)
  // seguindo a mesma curva S da linha resumo — assim o gráfico bate com o principal.
  const refRow = filtEdt
    ? (D.allRecords.find(r => r.edt === filtEdt && r.resumen)
       || D.allRecords.find(r => r.edt === '4.5'))
    : D.allRecords.find(r => r.edt === '4.5');

  const refCurr    = refRow?.planSeries?.[currIdx] || 0;
  const refEnd     = refRow?.planSeries?.slice(currIdx).reduce((m, v) => Math.max(m, v || 0), refCurr) || 1;
  const planAtCurr = planArr[currIdx];
  const remaining  = totalInc - planAtCurr;

  for (let i = currIdx + 1; i < nWeeks; i++) {
    const refI     = refRow?.planSeries?.[i] || 0;
    const progress = refCurr < refEnd
      ? Math.max(0, Math.min(1, (refI - refCurr) / (refEnd - refCurr)))
      : 0;
    planArr[i] = planAtCurr + remaining * progress;
  }

  const planPct = planArr.map(v => +(v / totalInc * 100).toFixed(3));
  const realPct = planArr.map((_, i) => {
    if (i > currIdx) return null;
    const v = realArr[i];
    return v > 0 ? +(v / totalInc * 100).toFixed(3) : null;
  });

  // Mostrar sempre todas as semanas (0 → 100% planejado)
  const labels    = D.scurve.map(s => s.week);
  const planSlice = planPct;
  const realSlice = realPct;

  destroyChart('scurveFilteredChart');
  charts['scurveFilteredChart'] = new Chart(
    document.getElementById('scurveFilteredChart'),
    _scurveChartConfig(labels, planSlice, realSlice)
  );
}

// ── S-Curve cascade filter ────────────────────────────────────────────────────
// Like _getCascadeChildren but also exposes isConsolidated task leaves so the
// user can drill down to individual task types inside a discipline.
function _getScurveCascadeChildren(parentEdt) {
  if (!parentEdt) {
    // Root level: depth-3 resumen nodes (same as WBS cascade)
    return D.allRecords.filter(r => r.resumen && r.edt.split('.').length === 3);
  }
  const tree   = _consolTree || D.allRecords;
  const prefix = parentEdt + '.';
  // All descendants that are resumen nodes OR consolidated task leaves
  const under  = tree.filter(r =>
    (r.resumen || r.isConsolidated) && r.edt.startsWith(prefix)
  );
  // Keep only direct children: no intermediate resumen/consolidated node above them
  return under.filter(r =>
    !under.some(o => o !== r && r.edt.startsWith(o.edt + '.'))
  );
}

function buildScurveCascadeFilters() {
  const wrap = document.getElementById('scurveCascadeFilters');
  if (!wrap || !D) return;
  _scurveFilterEdt = '';
  wrap.innerHTML = '';
  _addScurveCascadeSelect(wrap, null);
}

function _addScurveCascadeSelect(wrap, parentEdt) {
  const items = _getScurveCascadeChildren(parentEdt);   // ← scope-aware, includes consolidated tasks
  if (!items.length) return;

  const sel = document.createElement('select');
  sel.className = 'cascade-sel';
  sel.innerHTML = `<option value="">— Todas —</option>`
    + items.map(r => `<option value="${r.edt}">${r.tarea.trim()}</option>`).join('');

  wrap.appendChild(sel);

  sel.addEventListener('change', () => {
    const all = [...wrap.querySelectorAll('select')];
    const idx = all.indexOf(sel);
    // Remove todos os selects filhos deste
    all.slice(idx + 1).forEach(s => s.remove());

    if (sel.value) {
      // Selecionou um item: desce um nível
      _scurveFilterEdt = sel.value;
      _addScurveCascadeSelect(wrap, _scurveFilterEdt);
    } else {
      // "— Todas —": volta ao nível pai (select anterior), não ao root
      const prevSel = idx > 0 ? all[idx - 1] : null;
      _scurveFilterEdt = prevSel ? prevSel.value : '';
    }
    renderScurveFiltered();
  });
}

// ── Areas Bar + Table ─────────────────────────────────────────────────────────
function renderAreasBar() {
  const n = parseInt(document.getElementById('areaLvlBox')?.value || '3');
  const rows = D.areas.filter(a => a.nivel === n);

  destroyChart('areasBarChart');
  charts['areasBarChart'] = new Chart(document.getElementById('areasBarChart'), {
    type: 'bar',
    data: {
      labels: rows.map(a => a.tarea.trim().slice(0,35)),
      datasets: [
        { label:'% Plan Pond.', data: rows.map(a => +(a.pctPlan*100).toFixed(2)), backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label:'% Real Pond.', data: rows.map(a => +(a.pctReal*100).toFixed(2)), backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top' } },
      scales: { x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderAreasTable() {
  const n = parseInt(document.getElementById('areaLvlBox')?.value || '3');
  const rows = D.areas.filter(a => a.nivel === n);

  document.getElementById('areasTable').innerHTML = tableWrap(
    `<tr>
      <th class="left">Área / Grupo</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
      <th>H-H</th><th>Incidencia</th><th>% Plan Pond.</th><th>% Real Pond.</th>
      <th>Desvío Pond.</th><th>% Comp. Plan</th><th>% Comp. Real</th><th>Status</th>
    </tr>`,
    rows.map(r => `<tr>
      <td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td>
      <td>${pct(r.incidencia,3)}</td><td>${pct(r.pctPlan)}</td><td>${pct(r.pctReal)}</td>
      <td class="${devClass(r.desvPond)}">${signPct(r.desvPond)}</td>
      <td>${pbarDuo(r.pctCompPlan,r.pctCompReal)}</td>
      <td>${pct(r.pctCompReal)}</td><td>${statusBadge(r)}</td>
    </tr>`).join('')
  );
}

// ── Top Desvios Bar + Table ───────────────────────────────────────────────────
function renderDesviosBar(rows) {
  const top  = rows.slice(0,15);
  const vals = top.map(r => +(r.desviacion*100).toFixed(2));

  destroyChart('desviosBarChart');
  charts['desviosBarChart'] = new Chart(document.getElementById('desviosBarChart'), {
    type:'bar',
    data: {
      labels: top.map(r => r.edt),
      datasets:[{ label:'Desvío (%)', data:vals,
        backgroundColor: vals.map(v => v<0 ? 'rgba(192,0,0,0.7)' : 'rgba(0,132,74,0.7)'),
        borderRadius:3 }]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => {
          const r = top[ctx.dataIndex];
          return [`Desvío: ${ctx.parsed.x.toFixed(2)}%`, r.tarea.trim().slice(0,50)];
        }}}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderDesviosTable(rows) {
  document.getElementById('desviosTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th><th>Desvío</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="${devClass(r.desviacion)}">${signPct(r.desviacion)}</td>
    </tr>`).join('')
  );
}

// ── Críticas Bar + Table ──────────────────────────────────────────────────────
function renderCriticasBar(rows) {
  const top = rows.slice(0,15);
  destroyChart('criticasBarChart');
  charts['criticasBarChart'] = new Chart(document.getElementById('criticasBarChart'), {
    type:'bar',
    data: {
      labels: top.map(r => r.edt),
      datasets:[
        { label:'% Plan', data: top.map(r => +(r.pctCompPlan*100).toFixed(1)), backgroundColor:'rgba(0,84,166,0.6)', borderRadius:3 },
        { label:'% Real', data: top.map(r => +(r.pctCompReal*100).toFixed(1)), backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{position:'top'},
        tooltip:{ callbacks:{ afterTitle: items => top[items[0].dataIndex]?.tarea.trim().slice(0,55) }}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' }, max:100 } }
    }
  });
}

function renderCriticasTable(rows) {
  document.getElementById('criticasTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th><th>Desvío</th><th>Impacto Pond.</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="dev-neg">${signPct(r.desviacion)}</td><td class="dev-neg">${signPct(r.desvPond)}</td>
    </tr>`).join('')
  );
}

// ── Sin Avance Charts + Table ─────────────────────────────────────────────────
function renderSinAvanceCharts(rows) {
  const areaNames = {};
  D.areas.filter(a => a.nivel===3).forEach(a => { areaNames[a.edt] = a.tarea.trim().slice(0,22); });

  const byCount = {}, byHH = {};
  rows.forEach(r => {
    const k = r.edt.split('.').slice(0,3).join('.');
    byCount[k] = (byCount[k]||0) + 1;
    byHH[k]    = (byHH[k]   ||0) + r.hh;
  });

  const keys   = Object.keys(byCount).sort();
  const labels = keys.map(k => areaNames[k] || k);
  const colors = ['#c00000','#e68a00','#0054a6','#6a0dad','#00844a','#008080','#888'];

  destroyChart('sinAvanceDonut');
  charts['sinAvanceDonut'] = new Chart(document.getElementById('sinAvanceDonut'), {
    type:'doughnut',
    data:{ labels, datasets:[{ data: keys.map(k=>byCount[k]),
      backgroundColor: colors.slice(0,keys.length), borderWidth:2, borderColor:'#fff' }] },
    options:{ responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'right', labels:{ font:{size:11}, padding:8 } } } }
  });

  const hhEntries = Object.entries(byHH).sort((a,b) => b[1]-a[1]);
  destroyChart('sinAvanceHHBar');
  charts['sinAvanceHHBar'] = new Chart(document.getElementById('sinAvanceHHBar'), {
    type:'bar',
    data:{ labels: hhEntries.map(([k]) => areaNames[k]||k),
      datasets:[{ label:'H-H sin avance', data: hhEntries.map(([,v]) => Math.round(v)),
        backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{ x:{ ticks:{ font:{size:11} } } } }
  });
}

function renderSinAvanceTable(rows) {
  document.getElementById('sinAvanceTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Inicio LB</th><th>Fin LB</th>
     <th>H-H</th><th>Incidencia</th><th>% Plan</th><th>% Real</th></tr>`,
    rows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td class="dev-neg">0.0%</td>
    </tr>`).join('')
  );
}

// ── Ranking Bar + Table ───────────────────────────────────────────────────────
function renderRankingBar(rows) {
  const vals = rows.map(r => +(Math.abs(r.desvPond)*100).toFixed(3));
  destroyChart('rankingBarChart');
  charts['rankingBarChart'] = new Chart(document.getElementById('rankingBarChart'), {
    type:'bar',
    data:{ labels: rows.map(r => r.edt),
      datasets:[{ label:'Impacto Ponderado (%)', data:vals,
        backgroundColor: vals.map(v => v>0.5?'rgba(192,0,0,0.7)':v>0.2?'rgba(230,138,0,0.7)':'rgba(0,84,166,0.6)'),
        borderRadius:3 }]
    },
    options:{
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false},
        tooltip:{ callbacks:{ label: ctx => {
          const r = rows[ctx.dataIndex];
          return [`Impacto: ${ctx.parsed.x.toFixed(3)}%`, r.tarea.trim().slice(0,50)];
        }}}
      },
      scales:{ x:{ ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderRankingTable(rows) {
  document.getElementById('rankingTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>H-H</th>
     <th>Incidencia</th><th>% Plan Pond.</th><th>% Real Pond.</th><th>Impacto Pond.</th>
     <th>% Plan</th><th>% Real</th><th>Clasif.</th></tr>`,
    rows.map((r,i) => {
      const imp = Math.abs(r.desvPond);
      const cls = imp>0.005?'badge badge-crit':imp>0.002?'badge badge-late':imp>0.0005?'badge badge-warn':'badge badge-ok';
      const lbl = imp>0.005?'CRITICO':imp>0.002?'ALTO':imp>0.0005?'MEDIO':'BAJO';
      return `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
        <td>${pct(r.pctPlan)}</td><td>${pct(r.pctReal)}</td>
        <td class="${devClass(r.desvPond)}">${signPct(r.desvPond)}</td>
        <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
        <td><span class="${cls}">${lbl}</span></td>
      </tr>`;
    }).join('')
  );
}

// ── Consolidado (Power Blocks) ────────────────────────────────────────────────
//
// Zonas de consolidación: { prefix, pbIdx, discIdx }
//
// Jerarquía real (ambas áreas idéntica):
//   4.5.4.2.{PB 1-35}.{disc}.{act…}  → BESS  (PB en índice 4, disciplina en 5)
//   4.5.5.2.{PB 1-26}.{disc}.{act…}  → PV    (PB en índice 4, disciplina en 5)
//   4.5.5.1.*                          → PV Trabajos Generales — excluido (prefijo distinto)
//
// Estructura real: 4.5.4.2.{disc}.{PB}.{task}  → discIdx=4 (antes del PB), pbIdx=5
const CONS_ZONES = [
  // BESS: 4.5.4.2.{disc}.{PB}.{task}  → disc-first (discIdx < pbIdx)
  { prefix: '4.5.4.2', pbIdx: 5, discIdx: 4, label: 'BESS' },
  // PV:   4.5.5.2.{PB}.{disc}.{task}  → PB-first  (pbIdx < discIdx)
  { prefix: '4.5.5.2', pbIdx: 4, discIdx: 5, label: 'PV'   },
];

// Etiqueta por segmento de disciplina (1=Civil, 2=Mecánica, 3=Eléctrica)
const DISC_LABELS = { '1': 'Civil', '2': 'Mecánica', '3': 'Eléctrica' };

// Comparación numérica de EDT (e.g. '4.5.10' > '4.5.2')
function edtCmp(a, b) {
  const ap = a.split('.').map(Number);
  const bp = b.split('.').map(Number);
  for (let i = 0; i < Math.max(ap.length, bp.length); i++) {
    const d = (ap[i] ?? 0) - (bp[i] ?? 0);
    if (d) return d;
  }
  return 0;
}

// Prefijo común de una lista de nombres → "Hincado PB01"…"Hincado PB26" → "Hincado"
// Paso 1: reducción char a char hasta prefijo común
// Paso 2: elimina fragmentos residuales de número-de-PB al final
//   "Hincado PB"  → "Hincado"   (PB sin dígito quedó del recorte)
//   "Hincado PB0" → "Hincado"   (dígito parcial quedó del recorte)
//   "Fundación de Hormigón" → sin cambio (no termina en PB)
function namePrefix(names) {
  if (!names.length) return '';
  if (names.length === 1) return names[0].replace(/[\s\-–_.]*PB\s*\d*\s*$/i, '').replace(/[\s\-–_.]+$/, '').trim() || names[0];
  let p = names[0];
  for (let i = 1; i < names.length && p.length > 0; i++) {
    while (names[i].indexOf(p) !== 0) p = p.slice(0, -1);
  }
  p = p.replace(/[\s\-–_.]*PB\s*\d*\s*$/i, '')  // quita fragmento "PB", "PB0", "PB01", etc.
       .replace(/[\s\-–_.]+$/, '')                 // quita separadores finales residuales
       .trim();
  return p || names[0];
}

function buildConsolidated() {
  if (!D) return [];

  const groups = new Map();

  for (const leaf of D.allLeaves) {
    const parts = leaf.edt.split('.');

    // ── 1. Determinar zona (solo EDT, sin texto) ──────────────────
    let zone = null;
    for (const z of CONS_ZONES) {
      const zp = z.prefix.split('.');
      if (parts.length > zp.length && zp.every((s, i) => parts[i] === s)) {
        zone = z; break;
      }
    }
    if (!zone || parts.length <= zone.pbIdx) continue;

    // ── 2. Construir clave de grupo: reemplazar segmento PB por '*' ─
    const pbVal  = parts[zone.pbIdx];
    const kParts = [...parts];
    kParts[zone.pbIdx] = '*';
    const key = kParts.join('.');

    // Detectar disciplina del grupo (segmento después del PB)
    const discSeg = parts[zone.discIdx] || '';
    // Solo consolidar disciplinas conocidas (1=Civil, 2=Mecánica, 3=Eléctrica)
    if (!DISC_LABELS[discSeg]) continue;
    const discLabel = DISC_LABELS[discSeg];
    const zoneLabel = `${zone.label} — ${discLabel}`;

    if (!groups.has(key)) {
      groups.set(key, {
        key,
        zonePrefix: zone.prefix,
        zoneLabel,
        discSeg,
        minEdt:     leaf.edt,
        leaves:     [],
        pbSet:      new Set(),   // todos los PB con esta actividad
        pbPlanSet:  new Set(),   // PB donde plan > 0
        pbAvSet:    new Set(),   // PB donde real > 0
        pbDevSet:   new Set(),   // PB donde real < plan
        hhSum:      0,
        incSum:     0,
        planSum:    0,   // Σ incidencia × pctCompPlan
        realSum:    0,   // Σ incidencia × pctCompReal
      });
    }

    const g = groups.get(key);

    // Mantener EDT mínima del grupo (orden correcto en tabla)
    if (edtCmp(leaf.edt, g.minEdt) < 0) g.minEdt = leaf.edt;

    g.leaves.push(leaf);
    g.pbSet.add(pbVal);
    if (leaf.pctCompPlan > 0)                        g.pbPlanSet.add(pbVal);
    if (leaf.pctCompReal > 0)                        g.pbAvSet.add(pbVal);
    if (leaf.pctCompReal < leaf.pctCompPlan - 0.001) g.pbDevSet.add(pbVal);

    g.hhSum   += leaf.hh;
    g.incSum  += leaf.incidencia;
    g.planSum += leaf.incidencia * leaf.pctCompPlan;
    g.realSum += leaf.incidencia * leaf.pctCompReal;
  }

  // ── 3. Finalizar grupos ───────────────────────────────────────────
  const result = [];
  groups.forEach(g => {
    const inc = g.incSum;
    result.push({
      key:        g.key,
      minEdt:     g.minEdt,
      zonePrefix: g.zonePrefix,
      zoneLabel:  g.zoneLabel,
      discSeg:    g.discSeg,
      // Nombre consolidado: prefijo común de todos los nombres (sin "PBxx")
      tarea:      namePrefix(g.leaves.map(l => l.tarea.trim())),
      count:      g.leaves.length,
      pbTotal:    g.pbSet.size,
      pbPlan:     g.pbPlanSet.size,
      pbAv:       g.pbAvSet.size,
      pbDev:      g.pbDevSet.size,
      hhTotal:    g.hhSum,
      hhPlan:     g.leaves.reduce((s, l) => s + l.hh * l.pctCompPlan, 0),
      hhReal:     g.leaves.reduce((s, l) => s + l.hh * l.pctCompReal, 0),
      incTotal:   inc,
      pesoPlan:   g.planSum,
      pesoReal:   g.realSum,
      planConsol: inc > 0 ? g.planSum / inc : 0,
      realConsol: inc > 0 ? g.realSum / inc : 0,
      gap:        inc > 0 ? (g.realSum - g.planSum) / inc : 0,
    });
  });

  // Ordenar por EDT mínima (numérico, no alfabético)
  result.sort((a, b) => edtCmp(a.minEdt, b.minEdt));
  return result;
}

function renderConsolidado() {
  if (!D) return;
  _consCache = buildConsolidated();
  applyConsolidadoFilters();
}

function applyConsolidadoFilters() {
  if (!D || !_consCache) return;
  const zone = document.getElementById('consZoneBox')?.value || '';
  const disc = document.getElementById('consDiscBox')?.value || '';
  const q    = (document.getElementById('consSearch')?.value || '').toLowerCase().trim();

  const rows = _consCache.filter(g =>
    (!zone || g.zonePrefix === zone) &&
    (!disc || g.discSeg    === disc)  &&
    (!q    || g.tarea.toLowerCase().includes(q))
  );

  // ── Stats bar ────────────────────────────────────────────────────
  const statsEl = document.getElementById('consStats');
  if (statsEl) {
    if (rows.length) {
      const totInc  = rows.reduce((s, g) => s + g.incTotal, 0);
      const totPlan = rows.reduce((s, g) => s + g.pesoPlan, 0);
      const totReal = rows.reduce((s, g) => s + g.pesoReal, 0);
      const totHH   = rows.reduce((s, g) => s + g.hhTotal, 0);
      const pC  = totInc > 0 ? totPlan / totInc : 0;
      const rC  = totInc > 0 ? totReal / totInc : 0;
      const gap = rC - pC;
      statsEl.innerHTML = `
        <span class="cs-pill">${rows.length} tipos de actividad</span>
        <span class="cs-pill">HH total <strong>${Math.round(totHH).toLocaleString()}</strong></span>
        <span class="cs-pill">Plan consol. <strong>${pct(pC)}</strong></span>
        <span class="cs-pill">Real consol. <strong>${pct(rC)}</strong></span>
        <span class="cs-pill ${gap < 0 ? 'cs-neg' : 'cs-pos'}">Gap <strong>${signPct(gap)}</strong></span>`;
    } else {
      statsEl.innerHTML = '';
    }
  }

  renderConsolidadoTable(rows);
}

function renderConsolidadoTable(rows) {
  const el = document.getElementById('consolidadoTable');
  if (!el) return;

  if (!rows.length) {
    el.innerHTML = `<p class="cons-empty">
      Sin actividades en las zonas de consolidación.<br>
      <small>Verifique que el archivo contenga datos bajo los EDT: 4.5.4.2 (BESS Power Blocks) · 4.5.5.2 (PV Power Blocks)</small>
    </p>`;
    return;
  }

  el.innerHTML = tableWrap(
    `<tr>
      <th class="left" style="min-width:220px">Actividad</th>
      <th class="left" style="min-width:150px">Zona / Disciplina</th>
      <th title="Total de Power Blocks que tienen esta actividad">PB total</th>
      <th title="PB donde el avance planificado al corte es &gt; 0">PB planif.</th>
      <th title="PB donde el avance real es &gt; 0">PB c/avance</th>
      <th title="PB donde el avance real es menor al planificado">PB c/desvío</th>
      <th title="Suma de incidencias de todos los PB del grupo">Peso total</th>
      <th title="Σ incidencia × plan — peso del avance planificado">Peso plan.</th>
      <th title="Σ incidencia × real — peso del avance real">Peso real</th>
      <th title="Promedio ponderado del avance planificado">Plan consol.</th>
      <th title="Promedio ponderado del avance real">Real consol.</th>
      <th title="Real consolidado − Plan consolidado">Gap</th>
    </tr>`,
    rows.map(g => {
      const gapCls = devClass(g.gap);
      const slug   = g.zonePrefix.replace(/\./g, '-');
      // Barra de progreso PB: avance / total
      const pctAv  = g.pbTotal > 0 ? (g.pbAv / g.pbTotal * 100).toFixed(0) : 0;
      const miniBar = `<div class="pb-mini" title="${g.pbAv}/${g.pbTotal} PB con avance real">
        <div class="pb-mini-fill" style="width:${pctAv}%"></div></div>`;
      return `<tr>
        <td class="left"><span class="cons-tarea">${g.tarea}</span></td>
        <td class="left"><span class="zone-tag zone-${slug}">${g.zoneLabel}</span></td>
        <td class="cons-pb-num"><strong>${g.pbTotal}</strong></td>
        <td>${g.pbPlan}</td>
        <td>${g.pbAv}${miniBar}</td>
        <td class="${g.pbDev > 0 ? 'dev-neg' : 'dev-neutral'}">${g.pbDev > 0 ? g.pbDev : '—'}</td>
        <td>${pct(g.incTotal, 3)}</td>
        <td>${pct(g.pesoPlan, 3)}</td>
        <td>${pct(g.pesoReal, 3)}</td>
        <td>${pct(g.planConsol)}</td>
        <td>${pct(g.realConsol)}</td>
        <td class="${gapCls}"><strong>${signPct(g.gap)}</strong></td>
      </tr>`;
    }).join('')
  );
}

function setupConsolidado() {
  on('consZoneBox', 'change', () => { if (D) applyConsolidadoFilters(); });
  on('consDiscBox', 'change', () => { if (D) applyConsolidadoFilters(); });
  on('consSearch',  'input',  () => { if (D) applyConsolidadoFilters(); });
}

// ── Virtual Consolidation Tree ────────────────────────────────────────────────
// BESS 4.5.4.2.{disc}.{PB}.{task}  — disc-first (discIdx=4 < pbIdx=5)
//   Real discipline nodes (4.5.4.2.1/2/3) kept; PB+task nodes skipped;
//   consolidated task leaves injected after each discipline node.
//
// PV   4.5.5.2.{PB}.{disc}.{task}  — PB-first (pbIdx=4 < discIdx=5)
//   All children skipped; from zone root, virtual discipline nodes
//   (using PB-1 real records as templates) + consolidated task leaves injected.
function buildConsolTree() {
  if (!D || !_consCache || !_consCache.length) return D?.allRecords || [];

  // Helper: aggregate PB metrics from a list of consolidated groups
  function aggPB(groups) {
    let pbTotal = 0, pbPlan = 0, pbAv = 0, pbDev = 0, pesoPlan = 0, pesoReal = 0;
    for (const g of groups) {
      pbTotal  = Math.max(pbTotal, g.pbTotal);
      pbPlan   = Math.max(pbPlan,  g.pbPlan);
      pbAv     = Math.max(pbAv,    g.pbAv);
      pbDev    = Math.max(pbDev,   g.pbDev);
      pesoPlan += g.pesoPlan;
      pesoReal += g.pesoReal;
    }
    return { pbTotal, pbPlan, pbAv, pbDev, pesoPlan, pesoReal };
  }

  // Helper: build one consolidated task leaf object
  function makeLeaf(g, taskEdt, nivel) {
    return {
      edt:            taskEdt,
      tarea:          g.tarea,
      nivel,
      resumen:        false,
      isConsolidated: true,
      hh:             g.hhTotal,
      incidencia:     g.incTotal,
      pctCompPlan:    g.planConsol,
      pctCompReal:    g.realConsol,
      desviacion:     g.gap,
      status:         calcStatus({ pctCompPlan: g.planConsol, pctCompReal: g.realConsol }),
      pbTotal:  g.pbTotal,
      pbPlan:   g.pbPlan,
      pbAv:     g.pbAv,
      pbDev:    g.pbDev,
      pesoPlan: g.pesoPlan,
      pesoReal: g.pesoReal,
    };
  }

  // Group _consCache by zonePrefix → discSeg → { items[] }
  const byZoneDisc = new Map();
  for (const g of _consCache) {
    if (!byZoneDisc.has(g.zonePrefix)) byZoneDisc.set(g.zonePrefix, new Map());
    const dm = byZoneDisc.get(g.zonePrefix);
    if (!dm.has(g.discSeg)) dm.set(g.discSeg, { items: [] });
    dm.get(g.discSeg).items.push(g);
  }

  // Fast EDT lookup — used to find PB-1 discipline template records for PV
  const recByEdt = new Map(D.allRecords.map(r => [r.edt, r]));

  const result = [];

  for (const rec of D.allRecords) {
    let handled = false;

    for (const z of CONS_ZONES) {
      const pbFirst = z.pbIdx < z.discIdx; // true for PV (pbIdx=4, discIdx=5)

      // ── Zone root node (e.g. "4.5.4.2" / "4.5.5.2") ──────────────────────────
      if (rec.edt === z.prefix && rec.resumen) {
        const zoneMap = byZoneDisc.get(z.prefix);
        if (zoneMap?.size) {
          const allItems = [...zoneMap.values()].flatMap(d => d.items);
          result.push({ ...rec, ...aggPB(allItems) });

          if (pbFirst) {
            // PV: inject virtual disc nodes + consolidated tasks under zone root
            const discEntries = [...zoneMap.entries()].sort((a, b) => +a[0] - +b[0]);
            for (const [discSeg, discData] of discEntries) {
              const dAgg    = aggPB(discData.items);
              const discInc = discData.items.reduce((s, g) => s + g.incTotal, 0);
              const discHH  = discData.items.reduce((s, g) => s + g.hhTotal, 0);
              // Use PB-1 real record as display template (tarea, nivel)
              const tmplEdt = `${z.prefix}.1.${discSeg}`;
              const tmpl    = recByEdt.get(tmplEdt);
              const discNivel    = tmpl?.nivel ?? ((rec.nivel || 1) + 2);
              const discPlanFrac = discInc > 0 ? dAgg.pesoPlan / discInc : 0;
              const discRealFrac = discInc > 0 ? dAgg.pesoReal / discInc : 0;
              result.push({
                edt:         tmplEdt,
                tarea:       tmpl?.tarea ?? (DISC_LABELS[discSeg] || `Disc. ${discSeg}`),
                nivel:       discNivel,
                resumen:     true,
                isVirtual:   true,
                hh:          discHH,
                incidencia:  discInc,
                pctCompPlan: discPlanFrac,
                pctCompReal: discRealFrac,
                desviacion:  discRealFrac - discPlanFrac,
                status:      calcStatus({ pctCompPlan: discPlanFrac, pctCompReal: discRealFrac }),
                ...dAgg,
              });
              // Consolidated task leaves (EDT from PB-1 template)
              const sorted = [...discData.items].sort((a, b) => edtCmp(a.minEdt, b.minEdt));
              for (const g of sorted) {
                const taskSeg = g.key.split('.').pop();
                result.push(makeLeaf(g, `${z.prefix}.1.${discSeg}.${taskSeg}`, discNivel + 1));
              }
            }
          }
        } else {
          result.push(rec);
        }
        handled = true; break;
      }

      if (!rec.edt.startsWith(z.prefix + '.')) continue;

      const zDepth   = z.prefix.split('.').length;
      const recDepth = rec.edt.split('.').length;

      if (pbFirst) {
        // PV: skip ALL children — already injected from the zone root handler
        handled = true; break;
      }

      // ── BESS disc-first: discipline nodes (zDepth+1) ──────────────────────────
      if (recDepth === zDepth + 1 && rec.resumen) {
        const discSeg  = rec.edt.split('.')[z.discIdx];
        const discData = byZoneDisc.get(z.prefix)?.get(discSeg);
        if (discData?.items.length) {
          result.push({ ...rec, ...aggPB(discData.items) });
          const sorted = [...discData.items].sort((a, b) => edtCmp(a.minEdt, b.minEdt));
          for (const g of sorted) {
            result.push(makeLeaf(g, g.key, rec.nivel + 1));
          }
        } else {
          result.push(rec);
        }
        handled = true; break;
      }

      // PB nodes and their children (depth > zDepth+1): skip entirely
      if (recDepth > zDepth + 1) { handled = true; break; }
    }

    if (!handled) result.push(rec);
  }

  return result;
}

// ── WBS Árbol ─────────────────────────────────────────────────────────────────
function initArbol() {
  collapsedNodes.clear();
  if (!D) return;
  // Default: show up to nivel 3, collapse nivel 4+ summaries
  // Use _consolTree (with virtual nodes) if available, otherwise raw records
  const tree = _consolTree || D.allRecords;
  tree.filter(r => r.resumen && r.nivel >= 4)
      .forEach(r => collapsedNodes.add(r.edt));
}

function isArbolHidden(edt) {
  const parts = edt.split('.');
  for (let i = 1; i < parts.length; i++) {
    if (collapsedNodes.has(parts.slice(0, i).join('.'))) return true;
  }
  return false;
}

function buildArbolRow(r, edtsWithChildren, hidden) {
  const lvl     = r.nivel || 1;
  const indent  = Math.max(0, lvl - 1) * 16;
  const hasKids = edtsWithChildren.has(r.edt);
  const isColl  = collapsedNodes.has(r.edt);

  const lvlCls = `tree-lvl${Math.min(lvl, 6)}`;
  const hidCls = hidden ? ' tree-hidden' : '';

  const toggleBtn = hasKids
    ? `<button class="tree-toggle" data-edt="${r.edt}" title="${isColl ? 'Expandir' : 'Colapsar'}">${isColl ? '▶' : '▼'}</button>`
    : `<span class="tree-no-toggle"></span>`;

  // Columns: EDT | Actividad | H-H | PBs | PB Plan. | PB Av. | PB Dev. | Incid. | INCD.PLAN | INCD.REAL | % Plan | % Real | Desvío | Status

  // ── Consolidated leaf ────────────────────────────────────────────────────────
  if (r.isConsolidated) {
    const devCls   = devClass(r.desviacion);
    const pbDevCls = r.pbDev > 0 ? 'dev-neg' : 'dev-neutral';
    const hhVal    = r.hh > 0 ? Math.round(r.hh).toLocaleString() : '—';
    return `<tr class="tree-row tree-leaf tree-consol ${lvlCls}${hidCls}" data-edt="${r.edt}" data-rowtype="consolidada">
      <td class="left">
        <div class="tree-edt-cell">
          <span class="tree-indent" style="width:${indent}px"></span>
          <span class="tree-no-toggle"></span>
          <span class="tree-consol-dot">●</span>
        </div>
      </td>
      <td class="left"><span class="cons-tarea-tree">${r.tarea}</span></td>
      <td>${hhVal}</td>
      <td class="cons-pb-num"><strong>${r.pbTotal}</strong></td>
      <td>${r.pbPlan}</td>
      <td>${r.pbAv}</td>
      <td class="${pbDevCls}">${r.pbDev != null ? r.pbDev : '—'}</td>
      <td>${pct(r.incidencia,3)}</td>
      <td class="incd-plan">${pct(r.pesoPlan,3)}</td>
      <td class="incd-real">${pct(r.pesoReal,3)}</td>
      <td>${pct(r.pctCompPlan)}</td>
      <td>${pct(r.pctCompReal)}</td>
      <td class="${devCls}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
    </tr>`;
  }

  // ── Virtual discipline node (PV consolidated discipline — isVirtual) ──────────
  if (r.isVirtual) {
    const devCls   = devClass(r.desviacion);
    const pbDevCls = r.pbDev > 0 ? 'dev-neg' : 'dev-neutral';
    const hhVal    = r.hh > 0 ? Math.round(r.hh).toLocaleString() : '—';
    return `<tr class="tree-row tree-summary tree-virtual ${lvlCls}${hidCls}" data-edt="${r.edt}" data-rowtype="resumen">
      <td class="left">
        <div class="tree-edt-cell">
          <span class="tree-indent" style="width:${indent}px"></span>
          ${toggleBtn}
          <span class="tree-virtual-tag">${r.tarea.slice(0,3).toUpperCase()}</span>
        </div>
      </td>
      <td class="left"><strong class="tree-virtual-label">${r.tarea}</strong></td>
      <td>${hhVal}</td>
      <td class="cons-pb-num"><strong>${r.pbTotal}</strong></td>
      <td>${r.pbPlan}</td>
      <td>${r.pbAv}</td>
      <td class="${pbDevCls}">${r.pbDev != null ? r.pbDev : '—'}</td>
      <td>${pct(r.incidencia,3)}</td>
      <td class="incd-plan">${pct(r.pesoPlan,3)}</td>
      <td class="incd-real">${pct(r.pesoReal,3)}</td>
      <td>${pct(r.pctCompPlan)}</td>
      <td>${pct(r.pctCompReal)}</td>
      <td class="${devCls}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
    </tr>`;
  }

  // ── Regular record ───────────────────────────────────────────────────────────
  const isSum   = r.resumen;
  const typeCls = isSum ? 'tree-summary' : 'tree-leaf';
  const devCls  = devClass(r.desviacion);
  const hh      = r.hh > 0 ? Math.round(r.hh).toLocaleString() : '—';
  const incid   = r.incidencia > 0 ? pct(r.incidencia, 3) : '—';
  const devVal  = r.incidencia > 0.0001 ? signPct(r.desviacion) : '—';

  const hasPB      = r.pbTotal != null;
  const pbDevCls2  = hasPB && r.pbDev > 0 ? 'dev-neg' : '';
  const incdPlan   = r.incidencia > 0 ? pct(r.incidencia * r.pctCompPlan, 3) : '—';
  const incdReal   = r.incidencia > 0 ? pct(r.incidencia * r.pctCompReal, 3) : '—';
  return `<tr class="tree-row ${typeCls} ${lvlCls}${hidCls}" data-edt="${r.edt}" data-rowtype="${isSum ? 'resumen' : 'actividad'}">
    <td class="left">
      <div class="tree-edt-cell">
        <span class="tree-indent" style="width:${indent}px"></span>${toggleBtn}<span class="tree-edt-code">${r.edt}</span>
      </div>
    </td>
    <td class="left">${r.tarea.trim()}</td>
    <td>${hh}</td>
    <td class="cons-pb-num">${hasPB ? `<strong>${r.pbTotal}</strong>` : '—'}</td>
    <td>${hasPB ? r.pbPlan : '—'}</td>
    <td>${hasPB ? r.pbAv   : '—'}</td>
    <td class="${pbDevCls2}">${hasPB ? r.pbDev : '—'}</td>
    <td>${incid}</td>
    <td class="incd-plan">${incdPlan}</td>
    <td class="incd-real">${incdReal}</td>
    <td>${pct(r.pctCompPlan)}</td>
    <td>${pct(r.pctCompReal)}</td>
    <td class="${devCls}">${devVal}</td>
    <td>${statusBadge(r)}</td>
  </tr>`;
}

function renderArbol(filter) {
  const q = (filter || '').toLowerCase().trim();
  const tbody = document.getElementById('arbolBody');
  if (!tbody || !D) return;

  const allRecs = (_consolTree || D.allRecords).filter(r => r.edt);

  // Pre-compute which EDT codes have children (O(n) instead of O(n²))
  const edtsWithChildren = new Set();
  allRecs.forEach(r => {
    const parts = r.edt.split('.');
    for (let i = 1; i < parts.length; i++) {
      edtsWithChildren.add(parts.slice(0, i).join('.'));
    }
  });

  if (q) {
    // Search mode: show flat matches, no hiding
    const matches = allRecs.filter(r =>
      r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)
    );
    tbody.innerHTML = matches.map(r => buildArbolRow(r, edtsWithChildren, false)).join('');
    return;
  }

  tbody.innerHTML = allRecs.map(r => {
    const wbsHidden = _wbsFilterEdt !== ''
      && r.edt !== _wbsFilterEdt
      && !r.edt.startsWith(_wbsFilterEdt + '.');
    return buildArbolRow(r, edtsWithChildren,
      wbsHidden || (!wbsHidden && isArbolHidden(r.edt)));
  }).join('');
}

// ── Cascade combobox filter ────────────────────────────────────────────────────
// Retorna os filhos resumen diretos de parentEdt na _consolTree.
// "Direto" = nenhum outro nó resumen intermediário entre parent e filho.
function _getCascadeChildren(parentEdt) {
  if (!parentEdt) {
    // Nível raiz: nós de profundidade 3 do D.allRecords
    return D.allRecords.filter(r => r.resumen && r.edt.split('.').length === 3);
  }
  const tree   = _consolTree || D.allRecords;
  const prefix = parentEdt + '.';
  // Todos os nós resumen descendentes
  const under  = tree.filter(r => r.resumen && r.edt.startsWith(prefix));
  // Manter só os diretos: sem nó resumen intermediário acima deles
  return under.filter(r =>
    !under.some(o => o !== r && r.edt.startsWith(o.edt + '.'))
  );
}

function buildCascadeFilters() {
  const wrap = document.getElementById('arbolCascadeFilters');
  if (!wrap || !D) return;
  _wbsFilterEdt = '';
  wrap.innerHTML = '';
  _addCascadeSelect(wrap, null);
}

function _addCascadeSelect(wrap, parentEdt) {
  const items = _getCascadeChildren(parentEdt);
  if (!items.length) return;

  const sel = document.createElement('select');
  sel.className = 'cascade-sel';
  sel.innerHTML = `<option value="">— Todas —</option>`
    + items.map(r => `<option value="${r.edt}">${r.tarea.trim()}</option>`).join('');

  wrap.appendChild(sel);

  sel.addEventListener('change', () => {
    // Remove selects posteriores a este
    const all = [...wrap.querySelectorAll('select')];
    all.slice(all.indexOf(sel) + 1).forEach(s => s.remove());

    _wbsFilterEdt = sel.value;

    if (_wbsFilterEdt) {
      // Expandir o caminho selecionado na árvore
      const parts = _wbsFilterEdt.split('.');
      for (let i = 1; i <= parts.length; i++)
        collapsedNodes.delete(parts.slice(0, i).join('.'));
      _addCascadeSelect(wrap, _wbsFilterEdt);
    }

    renderArbol(document.getElementById('arbolSearch')?.value || '');
  });
}


function setupArbol() {
  on('arbolSearch', 'input', e => {
    if (!D) return;
    renderArbol(e.target.value);
  });
  on('arbolExpandAll', 'click', () => {
    if (!D) return;
    collapsedNodes.clear();
    renderArbol(document.getElementById('arbolSearch')?.value || '');
  });
  on('arbolCollapseAll', 'click', () => {
    if (!D) return;
    (_consolTree || D.allRecords).filter(r => r.resumen).forEach(r => collapsedNodes.add(r.edt));
    renderArbol(document.getElementById('arbolSearch')?.value || '');
  });
  // Event delegation — survives innerHTML rebuilds
  document.getElementById('arbolBody')?.addEventListener('click', e => {
    const btn = e.target.closest('.tree-toggle');
    if (!btn || !D) return;
    const edt = btn.dataset.edt;
    if (collapsedNodes.has(edt)) collapsedNodes.delete(edt);
    else                          collapsedNodes.add(edt);
    renderArbol(document.getElementById('arbolSearch')?.value || '');
  });
}

// ── Future ────────────────────────────────────────────────────────────────────
function renderFuture() {
  document.getElementById('futureTable').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th>
     <th>H-H</th><th>Incidencia</th><th>% Real actual</th><th>% Recuperable</th><th>Fin LB</th></tr>`,
    D.future.map((r,i) => {
      const rec = r.incidencia*(1-r.pctCompReal);
      return `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
        <td>${pct(r.pctCompReal)}</td><td class="dev-pos">${pct(rec,3)}</td>
        <td>${fmtDate(r.fin)}</td>
      </tr>`;
    }).join('')
  );
}

// ── Tabla ─────────────────────────────────────────────────────────────────────
let tablaRows = [];
function renderTabla(filtered) {
  tablaRows = filtered || D.allLeaves;

  const areaBox = document.getElementById('areaBox');
  if (areaBox.options.length === 1) {
    [...new Set(D.allLeaves.map(r => r.edt.split('.').slice(0,2).join('.')))].sort()
      .forEach(g => areaBox.appendChild(new Option(g, g)));
  }

  document.getElementById('tablaWrap').innerHTML = tableWrap(
    `<tr><th>#</th><th class="left">Actividad</th><th>EDT</th><th>Dur.</th>
     <th>Inicio LB</th><th>Fin LB</th><th>H-H</th><th>Incidencia</th>
     <th>% Plan</th><th>% Real</th><th>Desvío</th><th>Status</th></tr>`,
    tablaRows.map((r,i) => `<tr>
      <td>${i+1}</td><td class="left">${r.tarea.trim()}</td>
      <td>${r.edt}</td><td>${r.duracion}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td><td>${pct(r.incidencia,4)}</td>
      <td>${pct(r.pctCompPlan)}</td><td>${pct(r.pctCompReal)}</td>
      <td class="${devClass(r.desviacion)}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
    </tr>`).join('')
  );
}

// ── Populate area dropdowns ───────────────────────────────────────────────────
function populateAreaDropdowns() {
  const areaNames = {};
  D.areas.filter(a => a.nivel===3).forEach(a => { areaNames[a.edt] = a.tarea.trim().slice(0,28); });
  const opts = Object.entries(areaNames).sort()
    .map(([k,v]) => `<option value="${k}">${k} — ${v}</option>`).join('');
  ['desvAreaBox','critAreaBox','sinAreaBox','rankAreaBox'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = '<option value="">— Todas las áreas —</option>' + opts;
  });
}

// ── Tab filters ───────────────────────────────────────────────────────────────
function setupTabFilters() {
  on('areaLvlBox',  'change', () => { renderAreasBar(); renderAreasTable(); });

  function updateDesv() {
    if (!D) return;
    const q = document.getElementById('desvSearch').value.toLowerCase();
    const area = document.getElementById('desvAreaBox').value;
    const rows = D.topDesvios.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderDesviosBar(rows); renderDesviosTable(rows);
  }
  on('desvSearch',  'input',  updateDesv);
  on('desvAreaBox', 'change', updateDesv);

  function updateCrit() {
    if (!D) return;
    const q = document.getElementById('critSearch').value.toLowerCase();
    const area = document.getElementById('critAreaBox').value;
    const rows = D.critical.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderCriticasBar(rows); renderCriticasTable(rows);
  }
  on('critSearch',  'input',  updateCrit);
  on('critAreaBox', 'change', updateCrit);

  function updateSin() {
    if (!D) return;
    const q = document.getElementById('sinSearch').value.toLowerCase();
    const area = document.getElementById('sinAreaBox').value;
    const rows = D.sinAvance.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderSinAvanceCharts(rows); renderSinAvanceTable(rows);
  }
  on('sinSearch',  'input',  updateSin);
  on('sinAreaBox', 'change', updateSin);

  function updateRank() {
    if (!D) return;
    const q      = document.getElementById('rankSearch').value.toLowerCase();
    const area   = document.getElementById('rankAreaBox').value;
    const clasif = document.getElementById('rankClasif').value;
    const rows = D.ranking.filter(r => {
      const imp = Math.abs(r.desvPond);
      const c = imp>0.005?'CRITICO':imp>0.002?'ALTO':imp>0.0005?'MEDIO':'BAJO';
      return (!q      || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
             (!area   || r.edt.startsWith(area)) &&
             (!clasif || c === clasif);
    });
    renderRankingBar(rows.slice(0,20)); renderRankingTable(rows);
  }
  on('rankSearch',  'input',  updateRank);
  on('rankAreaBox', 'change', updateRank);
  on('rankClasif',  'change', updateRank);
}

// ── Table filters (Tabla tab) ─────────────────────────────────────────────────
function setupTableFilters() {
  on('searchBox', 'input',  applyFilters);
  on('statusBox', 'change', applyFilters);
  on('areaBox',   'change', applyFilters);
  on('exportBtn', 'click',  exportCsv);
}

function applyFilters() {
  if (!D) return;
  const q      = document.getElementById('searchBox').value.toLowerCase();
  const status = document.getElementById('statusBox').value;
  const area   = document.getElementById('areaBox').value;
  renderTabla(D.allLeaves.filter(r =>
    (!q      || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
    (!status || r.status === status) &&
    (!area   || r.edt.startsWith(area))
  ));
}

function exportCsv() {
  const hdr = ['EDT','Actividad','Duracion','Inicio LB','Fin LB','HH','Incidencia','Plan%','Real%','Desvio','Status'];
  const lines = [hdr.join(','), ...tablaRows.map(r =>
    [r.edt, `"${r.tarea.trim()}"`, r.duracion, r.inicio, r.fin,
     r.hh.toFixed(0), r.incidencia.toFixed(6),
     (r.pctCompPlan*100).toFixed(2), (r.pctCompReal*100).toFixed(2),
     (r.desviacion*100).toFixed(2), r.status].join(',')
  )];
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([lines.join('\n')], {type:'text/csv;charset=utf-8;'}));
  a.download = 'actividades_construccion.csv';
  a.click();
}

// ── Simulador Tab ─────────────────────────────────────────────────────────────

/** "p.p." formatter — percentage points with sign */
function _ppFmt(v, d=2) {
  if (v == null) return '—';
  return (v >= 0 ? '+' : '') + (v * 100).toFixed(d) + ' p.p.';
}

/** Stable DOM prefix for a row's result cells */
function _simTabSafeId(edt) { return 'stt-' + edt.replace(/\./g, '-'); }

/** Leaves available in the Simulador tab, optionally filtered by scope EDT */
function _simTabGetLeaves() {
  if (!D || !_consolTree) return [];
  const scope    = document.getElementById('simtabScope')?.value || '';
  const ZONE_PFXS = CONS_ZONES.map(z => z.prefix + '.');
  const all = [
    ..._consolTree.filter(r => r.isConsolidated),
    ...D.allLeaves.filter(r => !ZONE_PFXS.some(p => r.edt.startsWith(p))),
  ];
  if (!scope) return all;
  return all.filter(r => r.edt === scope || r.edt.startsWith(scope + '.'));
}

/** Calculate recovery for one _simTabRows entry */
function _simTabCalcRow(row) {
  const leaf = _simTabGetLeaves().find(r => r.edt === row.edt);
  if (!leaf) return null;
  const isPB  = !!leaf.isConsolidated;
  const delta = row.delta || 0;
  if (isPB && row.mode === 'pb') {
    const maxD    = Math.max(0, leaf.pbTotal - leaf.pbAv);
    const d       = Math.max(0, Math.min(delta, maxD));
    const newReal = Math.min(1, leaf.pctCompReal + (leaf.pbTotal > 0 ? d / leaf.pbTotal : 0));
    const dReal   = newReal - leaf.pctCompReal;
    return { leaf, isPB, newReal, deltaReal: dReal, recovery: leaf.incidencia * dReal };
  } else {
    const maxD    = (1 - leaf.pctCompReal) * 100;
    const d       = Math.max(0, Math.min(delta, maxD));
    const newReal = Math.min(1, leaf.pctCompReal + d / 100);
    const dReal   = newReal - leaf.pctCompReal;
    return { leaf, isPB, newReal, deltaReal: dReal, recovery: leaf.incidencia * dReal };
  }
}

function _simTabTotalRecovery() {
  return _simTabRows.reduce((s, r) => { const res = _simTabCalcRow(r); return s + (res?.recovery || 0); }, 0);
}

function setupSimTab() {
  on('simtabReset',   'click',  () => { _simTabRows = []; renderSimTab(); });
  on('simtabCalc',    'click',  () => renderSimTab());
  on('simtabAddBtn',  'click',  _simTabAdd);
  on('simtabModePB',  'click',  () => _simTabSetMode('pb'));
  on('simtabModePct', 'click',  () => _simTabSetMode('pct'));
  on('simtabScope',   'change', () => _simTabPopulateActSelect());
}

function _simTabSetMode(mode) {
  _simTabMode = mode;
  document.getElementById('simtabModePB')?.classList.toggle('simtab-mode-active',  mode === 'pb');
  document.getElementById('simtabModePct')?.classList.toggle('simtab-mode-active', mode === 'pct');
  const lbl  = document.getElementById('simtabDeltaLabel');
  const unit = document.getElementById('simtabDeltaUnit');
  const inp  = document.getElementById('simtabDelta');
  if (lbl)  lbl.textContent  = mode === 'pb' ? 'Quantidade de PB adicionais' : 'Porcentagem adicional';
  if (unit) unit.textContent = mode === 'pb' ? 'PB' : '%';
  if (inp)  inp.step         = mode === 'pb' ? '1' : '0.5';
}

function _simTabPopulateActSelect() {
  const sel = document.getElementById('simtabActSelect');
  if (!sel || !D) return;
  const added = new Set(_simTabRows.map(r => r.edt));
  const leaves = _simTabGetLeaves().filter(r => !added.has(r.edt));
  sel.innerHTML = `<option value="">Selecione uma atividade</option>`
    + leaves.map(r => {
      const tag = r.isConsolidated ? ` [PB×${r.pbTotal}]` : ' [Reg]';
      return `<option value="${r.edt}">${r.tarea.trim()}${tag}</option>`;
    }).join('');
}

function _simTabAdd() {
  const edt = document.getElementById('simtabActSelect')?.value;
  if (!edt || _simTabRows.find(r => r.edt === edt)) return;
  const leaf = _simTabGetLeaves().find(r => r.edt === edt);
  if (!leaf) return;
  const isPB = !!leaf.isConsolidated;
  const mode = isPB ? _simTabMode : 'pct'; // non-PB always uses pct mode
  const delta = parseFloat(document.getElementById('simtabDelta')?.value) || 0;
  _simTabRows.push({ edt, delta, mode });
  const deltaEl = document.getElementById('simtabDelta');
  if (deltaEl) deltaEl.value = '0';
  renderSimTab();
}

function initSimTab() {
  if (!D) return;
  _simTabRows = [];
  // Date badge
  const dateEl = document.getElementById('simtabDate');
  if (dateEl && D.meta.dataDate) dateEl.textContent = fmtDate(D.meta.dataDate);
  else if (dateEl && D.meta.dataWeek) dateEl.textContent = D.meta.dataWeek;
  // Scope selector
  const scopeSel = document.getElementById('simtabScope');
  if (scopeSel) {
    const tops = D.allRecords.filter(r => r.resumen && r.edt.split('.').length === 3);
    scopeSel.innerHTML = `<option value="">— Todo el proyecto —</option>`
      + tops.map(r => `<option value="${r.edt}">${r.tarea.trim()}</option>`).join('');
  }
  _simTabPopulateActSelect();
  _simTabUpdateKPIs();
}

function _simTabColorEl(id, v) {
  const el = document.getElementById(id);
  if (!el) return;
  el.classList.remove('simtab-pos', 'simtab-neg');
  if (v < -0.0001) el.classList.add('simtab-neg');
  else if (v > 0.0001) el.classList.add('simtab-pos');
}

function _simTabUpdateKPIs() {
  if (!D) return;
  const dev      = D.meta.desvio;
  const pctPlan  = D.meta.pctPlan;
  const pctReal  = D.meta.pctReal;
  const totalRec = _simTabTotalRecovery();
  const simReal  = pctReal + totalRec;
  const newDev   = dev + totalRec;

  // Left panel — current situation
  set('simtabPctPlan',   pct(pctPlan));
  set('simtabPctReal',   pct(pctReal));
  set('simtabDevActual', _ppFmt(dev));
  // Right panel — simulated result
  set('simtabPctRealSim', pct(simReal));
  set('simtabRecTotal',   _ppFmt(totalRec));
  set('simtabNewDev',     _ppFmt(newDev));
  // Sidebar summary
  set('simtabSumDevActual',   _ppFmt(dev));
  set('simtabSumRecTotal',    _ppFmt(totalRec));
  set('simtabSumNewDev',      _ppFmt(newDev));
  set('simtabSumPctReal',     pct(pctReal));
  set('simtabSumPctRealSim',  pct(simReal));
  // Sidebar detail
  set('simtabDetDevActual',   _ppFmt(dev));
  set('simtabDetRecTotal',    _ppFmt(totalRec));
  set('simtabDetNewDev',      _ppFmt(newDev));

  // Progress bars
  const setBar = (id, val) => {
    const el = document.getElementById(id);
    if (el) el.style.width = (Math.min(1, Math.max(0, val)) * 100).toFixed(1) + '%';
  };
  setBar('simtabPctPlanBar',    pctPlan);
  setBar('simtabPctRealBar',    pctReal);
  setBar('simtabPctRealSimBar', simReal);

  // Color classes
  ['simtabDevActual','simtabSumDevActual','simtabDetDevActual'].forEach(id => _simTabColorEl(id, dev));
  ['simtabNewDev',   'simtabSumNewDev',   'simtabDetNewDev'  ].forEach(id => _simTabColorEl(id, newDev));
  ['simtabRecTotal', 'simtabSumRecTotal', 'simtabDetRecTotal' ].forEach(id => _simTabColorEl(id, totalRec));
}

/** Update result cells of one row WITHOUT re-rendering the entire table */
function _simTabUpdateRow(idx) {
  const row  = _simTabRows[idx];
  if (!row) return;
  const res  = _simTabCalcRow(row);
  if (!res)  return;
  const sid  = _simTabSafeId(row.edt);
  const recCls = res.recovery >= 0 ? 'simtab-pos simtab-res-cell' : 'simtab-neg simtab-res-cell';

  const nrEl  = document.getElementById(sid + '-nr');
  const recEl = document.getElementById(sid + '-rec');
  const impEl = document.getElementById(sid + '-imp');
  if (nrEl)  nrEl.textContent  = pct(res.newReal);
  if (recEl) { recEl.textContent = _ppFmt(res.recovery); recEl.className = recCls; }
  if (impEl) { impEl.textContent = _ppFmt(res.recovery); impEl.className = recCls; }

  // Total row
  const totalRec = _simTabTotalRecovery();
  const totEl = document.getElementById('simtab-total-rec');
  if (totEl) {
    totEl.textContent = _ppFmt(totalRec);
    totEl.className   = 'simtab-res-cell ' + (totalRec >= 0 ? 'simtab-pos' : 'simtab-neg');
  }
}

function _renderSimTabTable() {
  const wrap = document.getElementById('simtabTable');
  if (!wrap) return;

  if (_simTabRows.length === 0) {
    wrap.innerHTML = '<p class="subtitle" style="padding:12px 0">Nenhuma atividade adicionada. Use o formulário acima para adicionar.</p>';
    return;
  }

  const totalRec = _simTabTotalRecovery();
  const totCls   = totalRec >= 0 ? 'simtab-pos' : 'simtab-neg';

  const rowsHtml = _simTabRows.map((row, idx) => {
    const res = _simTabCalcRow(row);
    if (!res) return '';
    const { leaf, isPB } = res;
    const sid    = _simTabSafeId(row.edt);
    const maxD   = isPB && row.mode === 'pb'
      ? Math.max(0, leaf.pbTotal - leaf.pbAv)
      : ((1 - leaf.pctCompReal) * 100).toFixed(1);
    const step   = (isPB && row.mode === 'pb') ? '1' : '0.5';
    const recCls = res.recovery >= 0 ? 'simtab-pos simtab-res-cell' : 'simtab-neg simtab-res-cell';
    const devCls = devClass(leaf.desviacion);

    return `<tr>
      <td class="left">${leaf.tarea.trim()}</td>
      <td><span class="sim-row-badge ${isPB ? 'pb' : 'reg'}">${isPB ? 'PB' : 'Reg'}</span></td>
      <td>${isPB ? leaf.pbTotal : '—'}</td>
      <td>${isPB ? leaf.pbAv : '—'}</td>
      <td>${isPB ? (leaf.pbPlan ?? '—') : '—'}</td>
      <td><input type="number" class="simtab-row-delta" data-idx="${idx}"
           min="0" max="${maxD}" step="${step}" value="${row.delta}"></td>
      <td>${pct(leaf.pctCompReal)}</td>
      <td id="${sid}-nr">${pct(res.newReal)}</td>
      <td class="${recCls}" id="${sid}-rec">${_ppFmt(res.recovery)}</td>
      <td class="${recCls}" id="${sid}-imp">${_ppFmt(res.recovery)}</td>
      <td><button class="simtab-rm" data-idx="${idx}">
        <svg width="12" height="12" viewBox="0 0 16 16" fill="currentColor" style="flex-shrink:0">
          <path d="M5.5 5.5A.5.5 0 0 1 6 6v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m2.5 0a.5.5 0 0 1 .5.5v6a.5.5 0 0 1-1 0V6a.5.5 0 0 1 .5-.5m3 .5a.5.5 0 0 0-1 0v6a.5.5 0 0 0 1 0z"/>
          <path d="M14.5 3a1 1 0 0 1-1 1H13v9a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V4h-.5a1 1 0 0 1-1-1V2a1 1 0 0 1 1-1H6a1 1 0 0 1 1-1h2a1 1 0 0 1 1 1h3.5a1 1 0 0 1 1 1zM4.118 4 4 4.059V13a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V4.059L11.882 4zM2.5 3h11V2h-11z"/>
        </svg>
        Remover
      </button></td>
    </tr>`;
  }).join('');

  wrap.innerHTML = tableWrap(
    `<tr>
       <th class="left">Atividade</th>
       <th>Tipo</th>
       <th title="Total Power Blocks">PB Total</th>
       <th title="PBs com avanço real">PB Executados Atual</th>
       <th title="PBs planejados até o corte">PB Planejados até o corte</th>
       <th>PB Simulados (Adicionais)</th>
       <th>% Real Atual</th>
       <th>% Real Simulado</th>
       <th>Recuperação (p.p.)</th>
       <th>Impacto no Desvio (p.p.)</th>
       <th>Ação</th>
     </tr>`,
    rowsHtml +
    `<tr class="simtab-total-row">
       <td colspan="8" class="right" style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">
         TOTAL DE RECUPERAÇÃO DO CENÁRIO (p.p.)</td>
       <td colspan="2" class="${totCls} simtab-res-cell" id="simtab-total-rec"
           style="font-size:15px;font-weight:800">${_ppFmt(totalRec)}</td>
       <td></td>
     </tr>`
  );

  // Wire up inline delta inputs — only update result cells (no full re-render = no focus loss)
  [...wrap.querySelectorAll('.simtab-row-delta')].forEach(inp => {
    inp.addEventListener('input', () => {
      const idx = parseInt(inp.dataset.idx);
      _simTabRows[idx].delta = parseFloat(inp.value) || 0;
      _simTabUpdateRow(idx);
      _simTabUpdateKPIs();
      _renderSimTabChart();
    });
  });
  // Remove buttons — full re-render OK (focus isn't on the table after delete)
  [...wrap.querySelectorAll('.simtab-rm')].forEach(btn => {
    btn.addEventListener('click', () => {
      _simTabRows.splice(parseInt(btn.dataset.idx), 1);
      renderSimTab();
    });
  });
}

function _renderSimTabChart() {
  if (!D) return;
  const totalRec  = _simTabTotalRecovery();
  const sc        = D.scurve;
  const currIdx   = sc.findIndex(s => s.isCurrent);
  const from      = Math.max(0, currIdx - 16);
  const slice     = sc.slice(from, Math.min(sc.length, currIdx + 17));
  const currInSlice = currIdx - from;

  const planData  = slice.map(s => +(s.plan * 100).toFixed(3));
  const realData  = slice.map(s => s.real != null ? +(s.real * 100).toFixed(3) : null);
  // Simulated = same as real but the last known real point is bumped by recovery
  const lastRealI = realData.reduce((last, v, i) => v != null ? i : last, -1);
  const simData   = realData.map((v, i) =>
    v != null && i === lastRealI ? +(v + totalRec * 100).toFixed(3) : v
  );

  destroyChart('simtabChart');
  charts['simtabChart'] = new Chart(document.getElementById('simtabChart'), {
    type: 'line',
    data: {
      labels: slice.map(s => s.week),
      datasets: [
        { label: 'Planejado',
          data: planData,
          borderColor: '#2563eb', backgroundColor: 'transparent',
          pointRadius: 0, fill: false, tension: 0.3, borderWidth: 2 },
        { label: 'Real',
          data: realData,
          borderColor: '#16a34a', backgroundColor: 'transparent',
          pointRadius: ctx => ctx.dataIndex === lastRealI ? 5 : 0,
          fill: false, tension: 0.3, borderWidth: 2.5, spanGaps: false },
        { label: 'Simulado',
          data: simData,
          borderColor: '#16a34a', backgroundColor: 'transparent',
          borderDash: [6, 4],
          pointRadius: ctx => ctx.dataIndex === lastRealI ? 7 : 0,
          pointStyle: 'rectRot',
          fill: false, tension: 0.3, borderWidth: 2, spanGaps: false },
      ]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      interaction: { mode: 'index', intersect: false },
      plugins: {
        legend: { position: 'top', labels: { boxWidth: 10, pointStyle: false } },
        tooltip: { callbacks: { label: ctx => ctx.dataset.label + ': ' + (ctx.parsed.y ?? 0).toFixed(2) + '%' } }
      },
      scales: { y: { ticks: { callback: v => v + '%' }, min: 0 } }
    }
  });
}

function renderSimTab() {
  if (!D) return;
  _simTabUpdateKPIs();
  _renderSimTabTable();
  _renderSimTabChart();
  _simTabPopulateActSelect();
}

// ── Deviation Recovery Simulator ─────────────────────────────────────────────

/** Returns all leaf-level records available for simulation.
 *  Consolidated PB leaves come from _consolTree; regular leaves from D.allLeaves. */
function _getSimLeaves() {
  if (!D || !_consolTree) return [];
  const ZONE_PFXS = CONS_ZONES.map(z => z.prefix + '.');
  return [
    ..._consolTree.filter(r => r.isConsolidated),
    ...D.allLeaves.filter(r => !ZONE_PFXS.some(p => r.edt.startsWith(p))),
  ];
}

function _findSimLeaf(edt) {
  return _getSimLeaves().find(r => r.edt === edt) || null;
}

/** Stable DOM id for the result panel of a sim row */
function _simSafeId(edt) { return 'simres-' + edt.replace(/\./g, '-'); }

/** Calculate simulation result for one leaf + delta.
 *  isPB  → delta = whole PBs to add
 *  !isPB → delta = percentage points to add  */
function _calcSimResult(leaf, rawDelta) {
  if (!leaf) return null;
  const isPB  = !!leaf.isConsolidated;
  const delta = parseFloat(rawDelta) || 0;

  if (isPB) {
    const maxD    = Math.max(0, leaf.pbTotal - leaf.pbAv);
    const d       = Math.max(0, Math.min(delta, maxD));
    const newReal = Math.min(1, leaf.pctCompReal + (leaf.pbTotal > 0 ? d / leaf.pbTotal : 0));
    const dReal   = newReal - leaf.pctCompReal;
    return { isPB, newReal, deltaReal: dReal, recovery: leaf.incidencia * dReal, maxDelta: maxD };
  } else {
    const maxD    = (1 - leaf.pctCompReal) * 100;
    const d       = Math.max(0, Math.min(delta, maxD));
    const newReal = Math.min(1, leaf.pctCompReal + d / 100);
    const dReal   = newReal - leaf.pctCompReal;
    return { isPB, newReal, deltaReal: dReal, recovery: leaf.incidencia * dReal, maxDelta: maxD };
  }
}

function _buildSimResultHtml(res, leaf) {
  if (!res || res.deltaReal < 0.000001) {
    return `<span class="sim-res-lbl" style="opacity:.7">Ingrese un valor para ver la simulación</span>`;
  }
  return `
    <div class="sim-res-item">
      <span class="sim-res-lbl">% Real actual</span>
      <span class="sim-res-val">${pct(leaf.pctCompReal)}</span>
    </div>
    <span class="sim-res-sep">→</span>
    <div class="sim-res-item">
      <span class="sim-res-lbl">Nuevo % Real</span>
      <span class="sim-res-val up">${pct(res.newReal)}</span>
    </div>
    <span class="sim-res-sep">·</span>
    <div class="sim-res-item">
      <span class="sim-res-lbl">Δ Real actividad</span>
      <span class="sim-res-val up">+${pct(res.deltaReal)}</span>
    </div>
    <span class="sim-res-sep">·</span>
    <div class="sim-res-item">
      <span class="sim-res-lbl">Recuperación proyecto</span>
      <span class="sim-res-val up">+${pct(res.recovery, 3)}</span>
    </div>`;
}

function setupSim() {
  const inp = document.getElementById('simSearch');
  const sug = document.getElementById('simSuggest');
  if (!inp || !sug) return;

  inp.addEventListener('input', () => {
    if (!D || !_consolTree) { sug.classList.add('hidden'); return; }
    const q = inp.value.trim().toLowerCase();
    if (q.length < 2) { sug.classList.add('hidden'); return; }

    const hits = _getSimLeaves().filter(r =>
      !_simRows.has(r.edt) && (
        r.edt.toLowerCase().includes(q) ||
        r.tarea.toLowerCase().includes(q)
      )
    ).slice(0, 12);

    if (!hits.length) { sug.classList.add('hidden'); return; }

    sug.innerHTML = hits.map(r => {
      const isPB = !!r.isConsolidated;
      return `<div class="sim-sug-item" data-edt="${r.edt}">
        <span class="sim-sug-name">${r.tarea.trim()}</span>
        <span class="sim-sug-edt">${r.edt}</span>
        <span class="sim-sug-badge ${isPB ? 'pb' : 'reg'}">${isPB ? `${r.pbTotal} PBs` : 'Regular'}</span>
      </div>`;
    }).join('');
    sug.classList.remove('hidden');

    [...sug.querySelectorAll('.sim-sug-item')].forEach(el => {
      el.addEventListener('mousedown', e => {
        e.preventDefault();
        addSimRow(el.dataset.edt);
        inp.value = '';
        sug.classList.add('hidden');
      });
    });
  });

  inp.addEventListener('blur', () => setTimeout(() => sug.classList.add('hidden'), 150));
}

function initSim() {
  _simRows.clear();
  // Set current deviation KPI immediately
  const el = document.getElementById('simCurDev');
  if (el && D) {
    const dev = D.meta.desvio;
    el.textContent = signPct(dev);
    el.className   = 'sim-kpi-val ' + (dev < -0.0001 ? 'neg' : dev > 0.0001 ? 'pos' : '');
  }
  renderSim();
}

function addSimRow(edt) {
  if (_simRows.has(edt)) return;
  _simRows.set(edt, 0);
  renderSim();
}

function removeSimRow(edt) {
  _simRows.delete(edt);
  renderSim();
}

/** Update only the result panel of one row (called on input change — avoids focus loss) */
function _updateSimRow(edt) {
  const leaf = _findSimLeaf(edt);
  if (!leaf) return;
  const res   = _calcSimResult(leaf, _simRows.get(edt) || 0);
  const resEl = document.getElementById(_simSafeId(edt));
  if (resEl) resEl.innerHTML = _buildSimResultHtml(res, leaf);
}

/** Update summary KPI bar */
function _updateSimSummary() {
  if (!D) return;
  const curDev = D.meta.desvio;
  let totalRec = 0;
  for (const [edt, delta] of _simRows.entries()) {
    const leaf = _findSimLeaf(edt);
    if (!leaf) continue;
    const res = _calcSimResult(leaf, delta);
    if (res) totalRec += res.recovery;
  }
  const newDev = curDev + totalRec;

  const recEl    = document.getElementById('simRecovery');
  const newDevEl = document.getElementById('simNewDev');
  if (recEl) {
    recEl.textContent = totalRec > 0.000001
      ? '+' + pct(totalRec, 3)
      : totalRec < -0.000001 ? pct(totalRec, 3) : '—';
    recEl.className = 'sim-kpi-val ' + (totalRec > 0.000001 ? 'pos' : totalRec < -0.000001 ? 'neg' : '');
  }
  if (newDevEl) {
    newDevEl.textContent = signPct(newDev);
    newDevEl.className   = 'sim-kpi-val ' + (newDev < -0.0001 ? 'neg' : newDev > 0.0001 ? 'pos' : '');
  }
}

/** Full re-render of sim rows (called on add/remove) */
function renderSim() {
  if (!D) return;
  const rowsEl = document.getElementById('simRows');
  if (!rowsEl) return;

  rowsEl.innerHTML = '';

  if (_simRows.size === 0) {
    rowsEl.innerHTML = '<p class="subtitle" style="margin:12px 0 4px">Use el buscador para agregar actividades al escenario de simulación.</p>';
    _updateSimSummary();
    return;
  }

  for (const [edt, delta] of _simRows.entries()) {
    const leaf = _findSimLeaf(edt);
    if (!leaf) continue;

    const isPB   = !!leaf.isConsolidated;
    const safeId = _simSafeId(edt);
    const res    = _calcSimResult(leaf, delta);
    const devCls = devClass(leaf.desviacion);
    const maxD   = res
      ? (isPB ? String(res.maxDelta) : res.maxDelta.toFixed(1))
      : '0';

    const statsHtml = [
      `<div class="sim-stat"><span class="sim-stat-lbl">Incidencia</span><span class="sim-stat-val">${pct(leaf.incidencia, 3)}</span></div>`,
      isPB ? `<div class="sim-stat"><span class="sim-stat-lbl">PBs totales</span><span class="sim-stat-val">${leaf.pbTotal}</span></div>` : '',
      isPB ? `<div class="sim-stat"><span class="sim-stat-lbl">PBs c/ avance</span><span class="sim-stat-val">${leaf.pbAv} / ${leaf.pbTotal}</span></div>` : '',
      `<div class="sim-stat"><span class="sim-stat-lbl">% Plan</span><span class="sim-stat-val">${pct(leaf.pctCompPlan)}</span></div>`,
      `<div class="sim-stat"><span class="sim-stat-lbl">% Real</span><span class="sim-stat-val">${pct(leaf.pctCompReal)}</span></div>`,
      `<div class="sim-stat"><span class="sim-stat-lbl">Desvío</span><span class="sim-stat-val ${devCls}">${signPct(leaf.desviacion)}</span></div>`,
    ].join('');

    const inputHtml = isPB
      ? `<span class="sim-input-lbl">PBs adicionales <small style="opacity:.65">(máx. ${maxD})</small>:</span>
         <input type="number" class="sim-delta" data-edt="${edt}" min="0" max="${maxD}" step="1" value="${delta}">`
      : `<span class="sim-input-lbl">% adicional <small style="opacity:.65">(máx. ${maxD}%)</small>:</span>
         <input type="number" class="sim-delta" data-edt="${edt}" min="0" max="${maxD}" step="0.5" value="${delta}">
         <span class="sim-delta-unit">%</span>`;

    const div = document.createElement('div');
    div.className = 'sim-row';
    div.dataset.edt = edt;
    div.innerHTML = `
      <div class="sim-row-hdr">
        <button class="sim-remove" data-edt="${edt}" title="Quitar actividad">✕</button>
        <span class="sim-row-name">${leaf.tarea.trim()}</span>
        <span class="sim-row-edt">${edt}</span>
        <span class="sim-row-badge ${isPB ? 'pb' : 'reg'}">${isPB ? `${leaf.pbTotal} PBs` : 'Regular'}</span>
      </div>
      <div class="sim-stats">${statsHtml}</div>
      <div class="sim-input-row">${inputHtml}</div>
      <div class="sim-result-row" id="${safeId}">${_buildSimResultHtml(res, leaf)}</div>`;

    rowsEl.appendChild(div);

    div.querySelector('.sim-remove').addEventListener('click', () => removeSimRow(edt));
    div.querySelector('.sim-delta').addEventListener('input', e => {
      _simRows.set(edt, parseFloat(e.target.value) || 0);
      _updateSimRow(edt);
      _updateSimSummary();
    });
  }

  _updateSimSummary();
}

// ── Scenarios ─────────────────────────────────────────────────────────────────
function setupScenarios() {
  on('scCalcA', 'click', () => {
    const edt     = document.getElementById('scEdtA').value.trim();
    const newReal = parseFloat(document.getElementById('scRealA').value) / 100;
    const act     = D?.allLeaves?.find(r => r.edt === edt);
    const el      = document.getElementById('scResultA');
    if (!act) { el.innerHTML = 'Actividad no encontrada. Verifique el EDT.'; el.classList.remove('hidden'); return; }
    const gain     = act.incidencia*(newReal - act.pctCompReal);
    const newTotal = D.meta.pctReal + gain;
    el.innerHTML = `
      <strong>Actividad:</strong> ${act.tarea.trim()}<br>
      <strong>Incidencia:</strong> ${pct(act.incidencia,4)}<br>
      <strong>% Real actual:</strong> ${pct(act.pctCompReal)}<br>
      <strong>% Real propuesto:</strong> ${pct(newReal)}<br>
      <strong>Ganancia en % ponderado:</strong> ${signPct(gain)}<br>
      <strong>Nuevo % Real total estimado:</strong> ${pct(newTotal)}
      ${gain >= 0 ? ' <span class="ok">▲ Mejora</span>' : ' <span class="alert">▼ Sin mejora</span>'}
    `;
    el.classList.remove('hidden');
  });

  on('scCalcB', 'click', () => {
    const target = parseFloat(document.getElementById('scDevB').value) / 100;
    const avg    = D.meta.avgIncidencia;
    const needed = Math.ceil(target / avg);
    const el     = document.getElementById('scResultB');
    el.innerHTML = `
      <strong>Desvío a recuperar:</strong> ${pct(target)}<br>
      <strong>Incidencia media por actividad:</strong> ${pct(avg,4)}<br>
      <strong>Actividades necesarias (estimado):</strong>
      <span style="font-size:20px;color:#002f6c;font-weight:700">${needed}</span><br>
      <small>Asumiendo completar cada actividad de 0% → 100%</small>
    `;
    el.classList.remove('hidden');
  });
}

// ── Tabs ──────────────────────────────────────────────────────────────────────
function setupTabs() {
  document.querySelectorAll('.tab').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.tab').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab)?.classList.add('active');
    });
  });
}

// ── Helpers ───────────────────────────────────────────────────────────────────
function on(id, event, fn) { document.getElementById(id)?.addEventListener(event, fn); }
function destroyChart(id) { if (charts[id]) { charts[id].destroy(); delete charts[id]; } }
function set(id, v) { const el = document.getElementById(id); if (el) el.textContent = v; }
function pct(v, d=2)     { return v == null ? '—' : (v*100).toFixed(d)+'%'; }
function signPct(v, d=2) { return v == null ? '—' : (v>=0?'+':'')+(v*100).toFixed(d)+'%'; }
function fmtDate(iso) {
  if (!iso) return '—';
  const [y,m,d] = iso.split('-');
  return `${d}/${m}/${y}`;
}
function devClass(v) { return v==null?'dev-neutral':v<-0.001?'dev-neg':v>0.001?'dev-pos':'dev-neutral'; }
function statusBadge(r) {
  const map = { completed:['badge-completed','Completada'], inProgress:['badge-inProgress','En progreso'],
                late:['badge-late','Atrasada'], notStarted:['badge-notStarted','No iniciada'] };
  const [cls,lbl] = map[r.status] || ['badge-notStarted','—'];
  return `<span class="badge ${cls}">${lbl}</span>`;
}
function pbarDuo(plan, real) {
  return `<div class="pbar-wrap">
    <div class="pbar"><div class="pbar-fill pbar-plan" style="width:${(plan*100).toFixed(1)}%"></div></div>
    <div class="pbar"><div class="pbar-fill pbar-real" style="width:${(real*100).toFixed(1)}%"></div></div>
  </div>`;
}
function tableWrap(thead, tbody) {
  return `<div class="grid-wrap"><table><thead>${thead}</thead><tbody>${tbody}</tbody></table></div>`;
}

let toastTimer;
function showToast(msg, isError = false, ms = 3500) {
  const el = document.getElementById('toast');
  el.textContent = msg;
  el.style.background = isError ? 'var(--danger)' : 'var(--text)';
  el.classList.remove('hidden');
  clearTimeout(toastTimer);
  if (ms < 60000) toastTimer = setTimeout(() => el.classList.add('hidden'), ms);
}
