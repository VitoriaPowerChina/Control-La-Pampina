/* Dashboard Construccion — La Pampina | PowerChina | SheetJS */
'use strict';

// ── State ────────────────────────────────────────────────────────────────────
let D = null;
let charts = {};
const collapsedNodes    = new Set();
const expandedPbNodes   = new Set(); // consolidated leaves with PBs expanded in WBS
let _wbsFilterEdt     = ''; // EDT selecionado na cascata WBS; '' = sem filtro
let _scurveFilterEdt  = ''; // EDT selecionado no filtro da Curva S filtrada
let _sinAvanceRows    = []; // last rendered Sin Avance rows (for PDF export)
let _criticasRows     = []; // last rendered Críticas rows (for PDF export)
let _consCache = null;
let _consolTree = null;
let _simRows    = new Map(); // edt → delta (number: PBs or percentage points)  — Escenarios tab
let _simTabRows = [];        // { edt, delta, mode:'pb'|'pct' }                   — Simulador tab
let _simTabMode = 'pb';      // current add-form mode in Simulador tab
let _recTargetWeeks    = 7;     // Recovery analysis — target weeks for the analysis
let _recRateWeeks      = 3;     // Recovery analysis — weeks to average for recent rate
let _top5SortAsc = true;        // Resumen top5: true = ascending deviation (most negative first)
let _rsTargetWeeks     = 7;     // Resumen tab S-curve — recovery target weeks (independent)
let _rsRateWeeks       = 3;     // Resumen tab analysis panel — recent rate window (independent)

// ── Boot ─────────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', () => {
  setupTheme();
  setupTabs();
  setupFileUpload();
  setupSimTab();
  setupArbol();
  setupConsolidado();
  setupTabFilters();
  setupPdfExports();
  setupRecovery();
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
  showToast(t('toast.reading'), false, 60000);
  try {
    const buf = await file.arrayBuffer();
    // Yield ao browser para mostrar o toast antes do parsing pesado
    await new Promise(r => setTimeout(r, 30));
    showToast(t('toast.processing'), false, 60000);
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
    showToast(t('toast.loaded') + D.meta.dataWeek + ' · ' + D.meta.actTotal + t('toast.activities'));
  } catch(err) {
    showToast(t('toast.error') + err.message, true);
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

// ── Recovery Analysis ─────────────────────────────────────────────────────────

/** Setup the target-weeks dropdown */
/** Toggle sort direction of the Resumen Top-5 panel */
function toggleTop5Sort() {
  _top5SortAsc = !_top5SortAsc;
  const btn = document.getElementById('rsTop5SortBtn');
  if (btn) btn.textContent = _top5SortAsc ? '↑ Mayor primero' : '↓ Menor primero';
  if (D) _tryRender(renderResumen);
}

async function exportResumenPDF() {
  const el  = document.getElementById('resumenContent');
  const btn = document.getElementById('rsExportBtn');
  if (!el) return;

  const origTxt = btn ? btn.innerHTML : '';
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Generando...'; }

  // ── 1. Force landscape-friendly render width ──────────────────────
  // Content height is ~900px regardless of width (fixed chart heights + grids).
  // To get aspect > 1.43 (A4 landscape), width must be > 900 * 1.43 ≈ 1290px.
  // We use 1500px to be safe, giving aspect ≈ 1.67 — well above 1.43.
  const CAPTURE_W = 1500;
  const savedWidth    = el.style.width;
  const savedMaxWidth = el.style.maxWidth;
  const savedMinWidth = el.style.minWidth;
  const savedOverflow = el.style.overflow;
  const savedHeight   = el.style.height;

  el.style.width    = CAPTURE_W + 'px';
  el.style.maxWidth = CAPTURE_W + 'px';
  el.style.minWidth = CAPTURE_W + 'px';
  el.style.overflow = 'visible';
  el.style.height   = 'auto';
  if (btn) btn.style.display = 'none';   // hide export btn from snapshot

  // Wait two frames so CSS grid + charts re-layout
  await new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)));

  try {
    const bgColor = getComputedStyle(document.documentElement)
                      .getPropertyValue('--bg').trim() || '#f8fafc';

    const canvas = await html2canvas(el, {
      scale: 2,
      useCORS: true,
      allowTaint: true,
      backgroundColor: bgColor,
      logging: false,
      windowWidth: CAPTURE_W,
      width:  CAPTURE_W,
      height: el.scrollHeight,
      scrollX: 0,
      scrollY: 0
    });

    const { jsPDF } = window.jspdf;
    // A4 landscape: 297 × 210 mm — ONE page
    const PDF_W = 297, PDF_H = 210, MARGIN = 5;
    const pdf = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });

    const usableW = PDF_W - MARGIN * 2;  // 287 mm
    const usableH = PDF_H - MARGIN * 2;  // 200 mm
    const aspect  = canvas.width / canvas.height;

    let imgW, imgH;
    if (aspect > usableW / usableH) {
      imgW = usableW; imgH = usableW / aspect;   // constrain by width
    } else {
      imgH = usableH; imgW = usableH * aspect;   // constrain by height
    }
    const x = MARGIN + (usableW - imgW) / 2;
    const y = MARGIN + (usableH - imgH) / 2;

    pdf.addImage(canvas.toDataURL('image/jpeg', 0.93), 'JPEG', x, y, imgW, imgH);

    const sem   = document.getElementById('rsMetaSemana')?.textContent?.trim() || '';
    const fecha = document.getElementById('rsMetaFecha')?.textContent?.trim()?.replace(/\//g,'-') || '';
    pdf.save(`Resumen_LaPampina_${sem || fecha || 'export'}.pdf`);

  } catch (e) {
    alert('Error al exportar PDF: ' + e.message);
  } finally {
    // ── Restore original styles ─────────────────────────────────────
    el.style.width    = savedWidth;
    el.style.maxWidth = savedMaxWidth;
    el.style.minWidth = savedMinWidth;
    el.style.overflow = savedOverflow;
    el.style.height   = savedHeight;
    if (btn) { btn.style.display = ''; btn.disabled = false; btn.innerHTML = origTxt; }
  }
}

function setupRecovery() {
  const sel = document.getElementById('recTargetWeeks');
  if (sel) {
    sel.addEventListener('change', () => {
      _recTargetWeeks = parseInt(sel.value, 10) || 7;
      // Keep modal select in sync
      const modalSel = document.getElementById('recTargetWeeksModal');
      if (modalSel) modalSel.value = sel.value;
      if (D) _tryRender(renderRecovery);
    });
  }

  // Modal target-weeks select — mirrors main select and re-renders modal chart
  const selModal = document.getElementById('recTargetWeeksModal');
  if (selModal) {
    selModal.addEventListener('change', () => {
      _recTargetWeeks = parseInt(selModal.value, 10) || 7;
      // Keep main select in sync
      if (sel) sel.value = selModal.value;
      if (D) {
        _tryRender(renderRecovery);   // updates recTargetDate + all panels
        // After renderRecovery updates recTargetDate, sync to modal label
        requestAnimationFrame(_syncModalDate);
        // Re-render modal chart with new target
        requestAnimationFrame(() => renderRecScurve('recScurveModalChart'));
      }
    });
  }

  const selR = document.getElementById('recRateWeeks');
  if (selR) {
    selR.addEventListener('change', () => {
      _recRateWeeks = parseInt(selR.value, 10) || 3;
      if (D) {
        _tryRender(renderRecovery);
        _tryRender(renderResumen);
        _tryRender(renderResMiniScurve);
      }
    });
  }

  // ── Resumen analysis panel — recent rate window ──────────────────────────────
  const rsRateSel = document.getElementById('rsRateWeeks');
  if (rsRateSel) {
    rsRateSel.addEventListener('change', () => {
      _rsRateWeeks = parseInt(rsRateSel.value, 10) || 3;
      if (D) _tryRender(renderResumen);
    });
  }

  // ── Resumen S-curve week selector (card) ──────────────────────────────────
  const rsScurveSel = document.getElementById('rsScurveTargetWeeks');
  if (rsScurveSel) {
    rsScurveSel.addEventListener('change', () => {
      _rsTargetWeeks = parseInt(rsScurveSel.value, 10) || 7;
      const modalSel = document.getElementById('rsScurveTargetWeeksModal');
      if (modalSel) modalSel.value = rsScurveSel.value;
      if (D) { _tryRender(renderResMiniScurve); _syncRsTargetDate(); }
    });
  }

  // ── Resumen S-curve week selector (modal) ─────────────────────────────────
  const rsScurveSelModal = document.getElementById('rsScurveTargetWeeksModal');
  if (rsScurveSelModal) {
    rsScurveSelModal.addEventListener('change', () => {
      _rsTargetWeeks = parseInt(rsScurveSelModal.value, 10) || 7;
      const mainSel = document.getElementById('rsScurveTargetWeeks');
      if (mainSel) mainSel.value = rsScurveSelModal.value;
      if (D) {
        _tryRender(renderResMiniScurve);
        _syncRsTargetDate();
        requestAnimationFrame(() => renderResMiniScurve('rsScurveModalChart'));
      }
    });
  }
}

/**
 * Simulate forward from currIdx at a fixed weeklyRate; return weeks needed to reach 100%.
 * Uses S-curve planned increments in Phase 1 (but only the rate matters here, not
 * the plan shape). Returns null if rate is zero/invalid.
 */
function _weeksToRecover(sc, currIdx, pctReal, weeklyRate) {
  if (!weeklyRate || weeklyRate <= 0) return null;
  let proj = pctReal;
  // Phase 1: advance week by week inside the S-curve
  for (let i = currIdx + 1; i < sc.length; i++) {
    proj += weeklyRate;
    if (proj >= 1.0) return i - currIdx;
  }
  // Phase 2: beyond plan end
  const remaining = 1 - proj;
  if (remaining <= 0) return sc.length - 1 - currIdx;
  return (sc.length - 1 - currIdx) + Math.ceil(remaining / weeklyRate);
}

/** Average real weekly increment over the last n data points */
function _recRecentRate(sc, currIdx, n) {
  n = n || 4;
  const pts = [];
  for (let i = Math.max(0, currIdx - n); i <= currIdx; i++) {
    if (sc[i] && sc[i].real != null) pts.push(sc[i].real);
  }
  if (pts.length < 2) return 0;
  return (pts[pts.length - 1] - pts[0]) / (pts.length - 1);
}

/** Master render for the Recovery tab */
function renderRecovery() {
  if (!D) return;
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  if (currIdx < 0) return;

  const pctPlan = sc[currIdx].plan || 0;
  const pctReal = sc[currIdx].real || 0;

  // Target week in S-curve
  const targetIdx      = Math.min(sc.length - 1, currIdx + _recTargetWeeks);
  const planAtTarget   = sc[targetIdx]?.plan || 0;
  const recNeeded      = Math.max(0, planAtTarget - pctReal);   // gap to close
  const reqRate        = _recTargetWeeks > 0 ? recNeeded / _recTargetWeeks : 0;

  // Target date label
  const targetDateRaw  = sc[targetIdx]?.date;
  const _loc = _lang === 'zh' ? 'zh-CN' : _lang === 'en' ? 'en-US' : 'es-CL';
  const targetDateLbl  = targetDateRaw
    ? new Date(targetDateRaw + 'T12:00:00').toLocaleDateString(_loc, { day:'2-digit', month:'2-digit', year:'numeric' })
    : '—';

  // Recent rate: last 3 weeks
  const recentRate     = _recRecentRate(sc, currIdx, _recRateWeeks);
  const accelAdd       = reqRate - recentRate;
  const accelFactor    = recentRate > 0 ? reqRate / recentRate : null;

  // Average planned rate for remaining weeks (for scenarios)
  const weeksLeft      = sc.length - 1 - currIdx;
  const planRemaining  = (sc[sc.length - 1].plan || 1) - pctPlan;
  const planRate       = weeksLeft > 0 ? planRemaining / weeksLeft : 0;

  const set = (id, v) => { const el = document.getElementById(id); if (el) el.textContent = v; };
  const ppR = v => v != null ? (v * 100).toFixed(2) + ' p.p./sem' : '—';
  const ppV = v => v != null ? (v * 100).toFixed(2) + ' p.p.' : '—';

  // ── Target date label ────────────────────────────────────────────────────
  set('recTargetDate', targetDateLbl);

  // ── Executive header ─────────────────────────────────────────────────────
  set('recExecFecha',  fmtDate(D.meta.dataDate) || '—');
  set('recExecSemana', D.meta.dataWeek || '—');

  // ── KPI cards ────────────────────────────────────────────────────────────
  const devPP    = (pctReal - pctPlan) * 100;
  const optimRate  = reqRate * 1.2;
  const optimWeeks = optimRate > 0 ? Math.ceil(recNeeded / optimRate) : null;

  set('recKpiPlan',    (pctPlan * 100).toFixed(2) + '%');
  set('recKpiReal',    (pctReal * 100).toFixed(2) + '%');
  set('recKpiDesv',    (devPP >= 0 ? '+' : '') + devPP.toFixed(2) + ' p.p.');
  set('recKpiReq',     (reqRate * 100).toFixed(2) + ' p.p./sem');
  set('recKpiReqSem',  _recTargetWeeks);
  set('recKpiAccel',   accelFactor != null ? accelFactor.toFixed(2) + 'x' : '—');
  set('recKpiOptRate',  (optimRate * 100).toFixed(2) + ' p.p./sem');
  set('recKpiOptWeeks', optimWeeks != null ? optimWeeks + ' sem' : '—');

  // Color the accel KPI based on severity
  const accelCard = document.querySelector('.rec-exec-kpi-card:nth-child(5)');
  if (accelCard && accelFactor != null) {
    accelCard.style.borderLeft = accelFactor > 1.5 ? '3px solid var(--danger)'
      : accelFactor > 1.1 ? '3px solid var(--warning)' : '3px solid var(--success)';
    const valEl = document.getElementById('recKpiAccel');
    if (valEl) valEl.style.color = accelFactor > 1.5 ? 'var(--danger)'
      : accelFactor > 1.1 ? 'var(--warning)' : 'var(--success)';
  }

  // ── Projection table title ────────────────────────────────────────────────
  const projTitleEl = document.getElementById('recProjTitleEl');
  if (projTitleEl) projTitleEl.textContent = t('rec.projTitle').replace('{n}', _recTargetWeeks);

  // ── Footer legend ─────────────────────────────────────────────────────────
  const targetWkLbl = sc[targetIdx]?.week || '';
  set('recFooterLegend', `${targetWkLbl}: semana objetivo del proyecto`);

  // ── Donut + area table ───────────────────────────────────────────────────
  // Pass actual S-curve deviation so the total row matches DESVÍO ACUM.
  _renderRecDesvDonut(pctReal - pctPlan);

  // ── Top activities ───────────────────────────────────────────────────────
  _renderRecTopActivities(devPP);

  // ── Levers ───────────────────────────────────────────────────────────────
  _renderRecLevers(reqRate, recNeeded);

  // ── Executive message ────────────────────────────────────────────────────
  _renderRecMessage(pctReal - pctPlan, pctReal, recNeeded, reqRate, accelFactor);

  // ── Analysis rows panel ──────────────────────────────────────────────────
  const rowsEl = document.getElementById('recAnalysisRows');
  if (rowsEl) {
    const mkRow = (lbl, val, valCls = '') => `
      <div class="rec-panel-row">
        <span class="rec-panel-row-lbl">${lbl}</span>
        <span class="rec-panel-row-val ${valCls}">${val}</span>
      </div>`;
    const reqRateHtml = `
      <div class="rec-panel-row rec-panel-row-highlight">
        <span class="rec-panel-row-lbl">${t('rec.row.reqRate')}</span>
        <span class="rec-panel-row-val" style="color:var(--primary)">${ppR(reqRate)}</span>
      </div>`;
    const accelCls  = accelAdd > 0 ? 'rec-val-danger' : 'rec-val-ok';
    const factorCls = accelFactor != null && accelFactor > 1.5 ? 'rec-val-danger'
                    : accelFactor != null && accelFactor > 1.1 ? 'rec-val-warn' : 'rec-val-ok';

    rowsEl.innerHTML =
      mkRow(t('rec.row.planAtTarget'), (planAtTarget * 100).toFixed(2) + '%') +
      mkRow(t('rec.row.realNow'),      (pctReal * 100).toFixed(2) + '%') +
      mkRow(t('rec.row.recNeeded'),    ppV(recNeeded)) +
      reqRateHtml +
      mkRow(t('rec.row.realRate').replace('{n}', _recRateWeeks), ppR(recentRate), 'rec-val-warn') +
      mkRow(t('rec.row.accelAdd'),     ppR(accelAdd),   accelCls) +
      mkRow(t('rec.row.accelFactor'),  accelFactor != null ? accelFactor.toFixed(2) + '×' : '—', factorCls);
  }

  // ── Status badge ─────────────────────────────────────────────────────────
  const badge = document.getElementById('recStatusBadge');
  const note  = document.getElementById('recStatusNote');
  if (badge) {
    if (accelFactor == null || recNeeded <= 0) {
      badge.textContent = t('rec.feasible');
      badge.className = 'rec-status-badge rec-status-ok';
      if (note) note.textContent = t('rec.sum.feasible').replace(/<[^>]+>/g, '');
    } else if (accelFactor > 1.5) {
      badge.textContent = t('rec.statusCrit');
      badge.className = 'rec-status-badge rec-status-danger';
      if (note) note.textContent = t('rec.notesCrit');
    } else if (accelFactor > 1.1) {
      badge.textContent = t('rec.statusWarn');
      badge.className = 'rec-status-badge rec-status-warn';
      if (note) note.textContent = t('rec.notesWarn');
    } else {
      badge.textContent = t('rec.statusOk');
      badge.className = 'rec-status-badge rec-status-ok';
      if (note) note.textContent = t('rec.notesOk');
    }
  }

  // ── Scenarios table ──────────────────────────────────────────────────────
  const scenEl = document.getElementById('recScenariosTable');
  if (scenEl) {
    const scenarios = [
      { name: t('rec.scen.current'),   rate: recentRate },
      { name: t('rec.scen.moderate'),  rate: (recentRate + reqRate) / 2 },
      { name: t('rec.scen.required'),  rate: reqRate,     hl: true },
      { name: t('rec.scen.accel'),     rate: reqRate * 1.2 },
      { name: t('rec.scen.intensive'), rate: reqRate * 1.5 },
    ];
    const scenRows = scenarios.map(s => {
      const wks     = _weeksToRecover(sc, currIdx, pctReal, s.rate);
      const netGain = s.rate - planRate;
      const dot     = s.rate <= 0 ? 'danger' : s.rate >= reqRate ? 'ok' : 'warn';
      const netCls  = netGain >= 0 ? 'rec-val-ok' : 'rec-val-danger';
      return `<tr ${s.hl ? 'class="rec-scen-hl"' : ''}>
        <td><strong>${s.name}</strong></td>
        <td>${(s.rate * 100).toFixed(2)}</td>
        <td class="${netCls}">${netGain >= 0 ? '+' : ''}${(netGain * 100).toFixed(2)}</td>
        <td>${wks != null ? wks : '—'}</td>
        <td><span class="rec-dot rec-dot-${dot}"></span></td>
      </tr>`;
    }).join('');
    scenEl.innerHTML = `<table class="rec-scen-table">
      <thead><tr>
        <th>${t('rec.scen.hdrName')}</th>
        <th>${t('rec.scen.hdrRate')}<br><small>p.p./sem</small></th>
        <th>${t('rec.scen.hdrNet')}<br><small>p.p./sem</small></th>
        <th>${t('rec.scen.hdrWeeks')}</th>
        <th>${t('rec.scen.hdrFeas')}</th>
      </tr></thead>
      <tbody>${scenRows}</tbody>
    </table>`;
  }

  // ── Projection table ─────────────────────────────────────────────────────
  const projTitle = document.getElementById('recProjTitleEl');
  if (projTitle) projTitle.textContent = t('rec.projTitle').replace('{n}', _recTargetWeeks);

  const projBody = document.getElementById('recProjBody');
  if (projBody) {
    let projReal = pctReal, projReq = pctReal, rows = '';
    for (let i = 1; i <= _recTargetWeeks; i++) {
      const scIdx   = currIdx + i;
      projReal     += recentRate;
      projReq      += reqRate;
      const planAcum = scIdx < sc.length ? sc[scIdx].plan : null;
      const devEsp   = planAcum != null ? projReal - planAcum : null;
      const addReal  = projReq - projReal;
      const fPct = v => (Math.min(1, Math.max(0, v)) * 100).toFixed(2) + '%';
      const fDev = v => v == null ? '—'
        : `<span style="color:${v >= 0 ? 'var(--success)' : 'var(--danger)'}">${v > 0 ? '+' : ''}${(v * 100).toFixed(2)}%</span>`;
      const fAdd = v =>
        `<span style="color:${v > 0.0002 ? 'var(--danger)' : v < -0.0002 ? 'var(--success)' : 'var(--text-muted)'}">${v > 0 ? '+' : ''}${(v * 100).toFixed(2)}%</span>`;
      const weekLbl = scIdx < sc.length ? (sc[scIdx].week || `+${i}`) : `+${i}`;
      rows += `<tr>
        <td>${weekLbl}</td>
        <td>${planAcum != null ? fPct(planAcum) : '—'}</td>
        <td>${fPct(projReal)}</td>
        <td>${fPct(projReq)}</td>
        <td>${fDev(devEsp)}</td>
        <td>${fAdd(addReal)}</td>
      </tr>`;
    }
    projBody.innerHTML = rows;
  }

  // ── Charts ───────────────────────────────────────────────────────────────
  _tryRender(renderRecScurve);
  _tryRender(renderRecWeekly);
}

function _renderRecDesvDonut(scDeviation) {
  const canvas = document.getElementById('recDesvAreaDonut');
  const tableEl = document.getElementById('recDesvAreaTable');
  if (!canvas || !D) return;

  const areas = (D.areas || []).filter(a => a.nivel === 3 && a.desvPond < 0);
  if (!areas.length) return;

  const totalDesvPond = areas.reduce((s, a) => s + Math.abs(a.desvPond), 0);
  const COLORS = ['#c00000','#e68a00','#2563eb','#7c3aed','#059669','#0891b2','#64748b'];

  const labels = areas.map(a => a.tarea.trim().slice(0, 22));
  const vals   = areas.map(a => +(Math.abs(a.desvPond) * 100).toFixed(3));

  destroyChart('recDesvAreaDonut');
  charts['recDesvAreaDonut'] = new Chart(canvas, {
    type: 'doughnut',
    data: { labels, datasets: [{ data: vals, backgroundColor: COLORS, borderWidth: 1 }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.parsed.toFixed(3)} p.p.` } }
      },
      cutout: '62%'
    }
  });

  // Table
  if (tableEl) {
    const rows = areas.map((a, i) => {
      const pct = totalDesvPond > 0 ? (Math.abs(a.desvPond) / totalDesvPond * 100).toFixed(1) : '—';
      return `<tr>
        <td style="text-align:left"><span style="display:inline-block;width:10px;height:10px;border-radius:2px;
          background:${COLORS[i % COLORS.length]};margin-right:5px;vertical-align:middle"></span>
          ${a.tarea.trim().slice(0,24)}</td>
        <td style="text-align:right;color:var(--danger);font-weight:600">${(a.desvPond * 100).toFixed(2)}</td>
        <td style="text-align:right;color:var(--text-muted)">${pct}%</td>
      </tr>`;
    }).join('');
    tableEl.innerHTML = `<table class="rec-desv-area-tbl">
      <thead><tr><th>Área</th><th style="text-align:right">Desvío (p.p.)</th><th style="text-align:right">% del desvío</th></tr></thead>
      <tbody>${rows}</tbody>
    </table>`;
  }

  const totalEl = document.getElementById('recDesvTotalRow');
  if (totalEl) {
    // Use the S-curve deviation (pctReal - pctPlan) when available so it matches
    // the DESVÍO ACUM. KPI; fall back to summing area desvPond otherwise.
    const tot = scDeviation != null
      ? scDeviation * 100
      : areas.reduce((s, a) => s + a.desvPond * 100, 0);
    totalEl.innerHTML = `<span>DESVÍO TOTAL DEL PROYECTO</span><span style="color:var(--danger);font-weight:800">${tot.toFixed(2)} p.p.</span>`;
  }
}

function _renderRecTopActivities(devTotalPP) {
  const el = document.getElementById('recTopActivities');
  if (!el || !D) return;

  const top5 = D.ranking.filter(r => r.desvPond < 0)
    .sort((a, b) => a.desvPond - b.desvPond)
    .slice(0, 5);

  const areaMap = _buildAreaMap();
  const totalNeg = top5.reduce((s, r) => s + Math.abs(r.desvPond), 0);
  const maxAbs   = top5.length ? Math.abs(top5[0].desvPond) : 1;

  const rows = top5.map((r, i) => {
    const area    = _areaOfEdt(r.edt, areaMap);
    const desvPP  = (r.desviacion * 100).toFixed(2);
    const pctDev  = devTotalPP !== 0 ? Math.abs(r.desvPond * 100 / (devTotalPP / 100)).toFixed(1) : '—';
    const barW    = Math.round(Math.abs(r.desvPond) / maxAbs * 100);
    const imp     = Math.abs(r.desvPond) * 100;
    const badge   = imp > 0.005 ? 'badge badge-crit' : imp > 0.002 ? 'badge badge-late' : 'badge badge-warn';
    const lbl     = imp > 0.005 ? 'Crítico' : imp > 0.002 ? 'Alto' : 'Medio';
    return `<tr>
      <td style="font-weight:700;color:var(--text-muted)">${i+1}</td>
      <td style="font-weight:600;text-align:left">${r.tarea.trim()}</td>
      <td style="font-size:10px;color:var(--primary);text-align:left">${area}</td>
      <td style="color:var(--danger);font-weight:700;text-align:right">${desvPP}</td>
      <td style="text-align:right">${pctDev}%</td>
      <td><div class="rec-act-bar-wrap"><div class="rec-act-bar" style="width:${barW}%"></div></div></td>
      <td><span class="${badge}">${lbl}</span></td>
    </tr>`;
  }).join('');

  el.innerHTML = `<table class="rec-act-table">
    <thead><tr>
      <th>#</th><th>Actividad</th><th>Área</th>
      <th style="text-align:right">Desvío (p.p.)</th>
      <th style="text-align:right">% del desvío</th>
      <th>Impacto</th><th>Estado</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;

  // subtotal note
  const noteEl = document.getElementById('recActivitiesNote');
  if (noteEl) {
    const subtotal = top5.reduce((s, r) => s + r.desvPond * 100, 0);
    const pctTot   = devTotalPP !== 0 ? Math.abs(subtotal / devTotalPP * 100).toFixed(1) : '—';
    noteEl.innerHTML = `<i class="bi bi-info-circle"></i> Estas 5 actividades representan el <b>${pctTot}%</b> del desvío total del proyecto (${devTotalPP.toFixed(2)} p.p.)`;
  }
}

function _renderRecLevers(reqRate, recNeeded) {
  const el = document.getElementById('recLeversTable');
  if (!el || !D) return;

  const levers = D.ranking.filter(r => r.desvPond < 0)
    .sort((a, b) => a.desvPond - b.desvPond)
    .slice(0, 5);

  const areaMap  = _buildAreaMap();
  const totalNeg = levers.reduce((s, r) => s + Math.abs(r.desvPond), 0) || 1;

  const rows = levers.map((r, i) => {
    const area   = _areaOfEdt(r.edt, areaMap);
    const pctImp = (Math.abs(r.desvPond) / totalNeg * 100).toFixed(0);
    const recPot = totalNeg > 0
      ? (Math.abs(r.desvPond) / totalNeg * reqRate * 100).toFixed(2)
      : '—';
    // Facilidad de ejecución: 1-5 based on desviacion severity and whether started
    const absDev = Math.abs(r.desviacion);
    let ease = r.pctCompReal > 0 ? 3 : 2;
    if (absDev < 0.15) ease += 1;
    else if (absDev > 0.6) ease -= 1;
    ease = Math.max(1, Math.min(5, ease));
    const dotsHtml = '●'.repeat(ease) + '<span style="opacity:.25">' + '●'.repeat(5 - ease) + '</span>';
    return `<tr>
      <td style="font-weight:700;color:var(--text-muted)">${i+1}</td>
      <td style="font-weight:600;text-align:left">${r.tarea.trim()}</td>
      <td style="font-size:10px;color:var(--primary);text-align:left">${area}</td>
      <td style="text-align:right;color:var(--success);font-weight:700">+${recPot}</td>
      <td style="text-align:right">${pctImp}%</td>
      <td class="rec-lev-dots">${dotsHtml}</td>
    </tr>`;
  }).join('');

  el.innerHTML = `<table class="rec-lev-table">
    <thead><tr>
      <th>#</th><th>Actividad</th><th>Área</th>
      <th style="text-align:right">Recup. potencial<br><small>(p.p./sem)</small></th>
      <th style="text-align:right">% Impacto</th>
      <th>Facilidad</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table>`;

  const noteEl = document.getElementById('recLeversNote');
  if (noteEl) {
    const totalRecovPot = levers.reduce((s, r) => s + Math.abs(r.desvPond) / totalNeg * reqRate * 100, 0);
    const pctCovered = recNeeded > 0 ? (totalRecovPot / (reqRate * 100) * 100).toFixed(0) : '—';
    noteEl.innerHTML = `<i class="bi bi-lightning-charge"></i> Con enfoque en estas actividades se puede recuperar hasta <b>+${totalRecovPot.toFixed(2)} p.p./sem</b> (${pctCovered}% de la meta).`;
  }
}

function _renderRecMessage(dev, pctReal, recNeeded, reqRate, accelFactor) {
  const el = document.getElementById('recMessagePanel');
  if (!el) return;

  const isCrit = accelFactor != null && accelFactor > 1.5;
  const isWarn = accelFactor != null && accelFactor > 1.1;
  const topActs = D.ranking.filter(r => r.desvPond < 0).slice(0, 3).map(r => r.tarea.trim());

  const bullets = [];

  // Bullet 1: deviation
  const devPP = Math.abs(dev * 100).toFixed(2);
  bullets.push({ cls: isCrit ? 'danger' : isWarn ? 'warn' : 'ok',
    text: `El desvío actual de <b>-${devPP} p.p.</b> requiere atención ${isCrit ? 'inmediata' : isWarn ? 'sostenida' : 'de seguimiento'}.` });

  // Bullet 2: main activities
  if (topActs.length) {
    bullets.push({ cls: 'warn',
      text: `Las actividades que más impactan el atraso son: <b>${topActs.join(', ')}</b>.` });
  }

  // Bullet 3: current rate
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  const rateRecent = _recRecentRate(sc, currIdx, _recRateWeeks);
  bullets.push({ cls: 'neutral',
    text: `Manteniendo el ritmo actual (<b>${(rateRecent * 100).toFixed(2)} p.p./sem</b>), el desvío aumentará.` });

  // Bullet 4: required rate
  bullets.push({ cls: isCrit ? 'danger' : 'warn',
    text: `Para recuperar en <b>${_recTargetWeeks} semanas</b>, se requiere avanzar <b>${(reqRate * 100).toFixed(2)} p.p./sem</b> (${accelFactor != null ? accelFactor.toFixed(2) + 'x' : '—'} el rendimiento actual).` });

  const colorMap = { danger: 'var(--danger)', warn: 'var(--warning)', ok: 'var(--success)', neutral: 'var(--text-muted)' };
  el.innerHTML = '<ul>' + bullets.map(b =>
    `<li style="color:${colorMap[b.cls]}">${b.text}</li>`
  ).join('') + '</ul>';
}

/** S-curve: reference design — plan / real / deviation fill / recovery line + annotations
 *  @param {string} [canvasId='recScurveChart'] — target canvas id (pass modal id to render there)
 *  @param {number} [tw]                         — override target weeks (default: _recTargetWeeks)
 */
function renderRecScurve(canvasId, tw) {
  canvasId = canvasId || 'recScurveChart';
  const targetWeeks = (tw != null ? tw : _recTargetWeeks);
  // Modal = either of the two modal canvases; callout always visible
  const isModal     = canvasId === 'recScurveModalChart' || canvasId === 'rsScurveModalChart';
  // Is this the Resumen mini chart (or its modal)?
  const isResumen   = canvasId === 'resMiniScurve' || canvasId === 'rsScurveModalChart';
  const canvas = document.getElementById(canvasId);
  if (!canvas || !D) return;

  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  if (currIdx < 0) return;

  const pctReal    = sc[currIdx].real  || 0;
  const planAtCurr = sc[currIdx].plan  || 0;
  const deviation  = pctReal - planAtCurr;           // negative = behind plan

  const targetIdx  = Math.min(sc.length - 1, currIdx + targetWeeks);
  const planAtTgt  = sc[targetIdx]?.plan || 0;
  const recNeeded  = Math.max(0, planAtTgt - pctReal);
  const reqRate    = targetWeeks > 0 ? recNeeded / targetWeeks : 0;
  const targetWeekLabel = sc[targetIdx]?.week || '';

  // ── Full view: start from week 0, end at target week ─────────────────────────
  const viewStart = 0;
  const viewEnd   = targetIdx;
  const scView    = sc.slice(viewStart, viewEnd + 1);
  // ci/ti are the same as currIdx/targetIdx since viewStart = 0
  const ci = currIdx;
  const ti = targetIdx;

  const weekLabels = scView.map(s => s.week);
  const planData   = scView.map(s => +(s.plan * 100).toFixed(2));

  // Real line: only up to cutoff
  const realData = scView.map((s, i) =>
    (i <= ci && s.real != null) ? +(s.real * 100).toFixed(2) : null);

  // Recovery line: cutoff → target
  const planRange = planAtTgt - planAtCurr;
  const recovData = scView.map((s, i) => {
    if (i < ci || i > ti) return null;
    const t = planRange !== 0 ? (s.plan - planAtCurr) / planRange : (i - ci) / targetWeeks;
    const v = pctReal + t * (planAtTgt - pctReal);
    return +(Math.min(1, Math.max(0, v)) * 100).toFixed(2);
  });

  // X-axis ticks: ~8 labels in history, every week in recovery window
  const histStep = Math.max(1, Math.round(ci / 8));
  const xTickLabels = scView.map((s, i) => {
    if (i >= ci) return s.week;                          // every week in recovery
    return (i % histStep === 0) ? s.week : null;         // sparse in history
  });

  // Update chart card title dynamically (only for the main recovery chart)
  if (canvasId === 'recScurveChart') {
    const titleEl = document.querySelector('.rec-curve-title');
    if (titleEl) titleEl.textContent = `CURVA S - PROYECTO  (Meta de recuperación: ${targetWeeks} semanas)`;
  }

  // ── Annotation plugin ────────────────────────────────────────────────────────
  const annotPlugin = {
    id: 'recAnnotations',
    afterDraw(chart) {
      const { ctx, scales: { x, y }, chartArea } = chart;
      const currX    = x.getPixelForValue(ci);
      const tgtX     = x.getPixelForValue(ti);
      const currRealY = y.getPixelForValue(pctReal * 100);
      const currPlanY = y.getPixelForValue(planAtCurr * 100);
      const tgtY     = y.getPixelForValue(planAtTgt * 100);

      ctx.save();

      // 1. Vertical cutoff line (dark, dashed)
      ctx.setLineDash([6, 4]);
      ctx.strokeStyle = 'rgba(30,41,59,.65)';
      ctx.lineWidth = 1.5;
      ctx.beginPath();
      ctx.moveTo(currX, chartArea.top);
      ctx.lineTo(currX, chartArea.bottom);
      ctx.stroke();
      ctx.setLineDash([]);

      // 2. "Fecha de corte" dark pill at bottom of cutoff line
      const fcText = t('rec.cutoffLbl');
      ctx.font = 'bold 10px sans-serif';
      const fcW = ctx.measureText(fcText).width + 18;
      const fcH = 22;
      const fcX = currX - fcW / 2;
      const fcY = chartArea.bottom - fcH - 4;  // inside chart area, just above x-axis
      ctx.fillStyle = '#1e293b';
      _rrect(ctx, fcX, fcY, fcW, fcH, 5);
      ctx.fill();
      ctx.fillStyle = '#fff';
      ctx.textAlign = 'center';
      ctx.fillText(fcText, currX, fcY + 14);

      // 3. Deviation annotation at cutoff: red bidirectional arrow + label
      if (deviation < 0 && currPlanY < currRealY - 6) {
        const midY = (currRealY + currPlanY) / 2;
        const arrX = currX + 10;

        // Arrow shaft
        ctx.strokeStyle = '#dc2626';
        ctx.lineWidth = 1.5;
        ctx.beginPath();
        ctx.moveTo(arrX, currPlanY + 5);
        ctx.lineTo(arrX, currRealY - 5);
        ctx.stroke();
        // Arrowhead up (toward plan)
        ctx.fillStyle = '#dc2626';
        ctx.beginPath();
        ctx.moveTo(arrX, currPlanY + 2);
        ctx.lineTo(arrX - 4, currPlanY + 10);
        ctx.lineTo(arrX + 4, currPlanY + 10);
        ctx.closePath(); ctx.fill();
        // Arrowhead down (toward real)
        ctx.beginPath();
        ctx.moveTo(arrX, currRealY - 2);
        ctx.lineTo(arrX - 4, currRealY - 10);
        ctx.lineTo(arrX + 4, currRealY - 10);
        ctx.closePath(); ctx.fill();

        // Label
        const devTxt = (deviation * 100).toFixed(2).replace('.', ',') + ' p.p.';
        ctx.font = 'bold 12px sans-serif';
        ctx.fillStyle = '#dc2626';
        ctx.textAlign = 'left';
        ctx.fillText(devTxt, arrX + 8, midY + 4);
      }

      // 4. Dashed vertical drop from endpoint to x-axis
      ctx.setLineDash([4, 3]);
      ctx.strokeStyle = 'rgba(37,99,235,.45)';
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(tgtX, tgtY);
      ctx.lineTo(tgtX, chartArea.bottom);
      ctx.stroke();
      ctx.setLineDash([]);


      ctx.restore();
    }
  };

  destroyChart(canvasId);
  charts[canvasId] = new Chart(canvas, {
    type: 'line',
    plugins: [annotPlugin],
    data: {
      labels: weekLabels,
      datasets: [
        // 0 — Plan
        { label: t('rec.lbl.plan'),
          data: planData, borderColor: '#2563eb',
          borderWidth: 2, pointRadius: 0, fill: false, tension: 0.3, order: 4 },
        // 1 — Real
        { label: t('rec.lbl.real'),
          data: realData, borderColor: '#166534',
          borderWidth: 2.5, pointRadius: 0, fill: false, tension: 0.3, spanGaps: false, order: 3 },
        // 2 — Recovery line (orange dashed, stops at targetIdx)
        { label: `${t('rec.lbl.recovery')} (meta ${targetWeeks} sem)`,
          data: recovData, borderColor: '#f97316',
          borderWidth: 2, borderDash: [8, 5], pointRadius: 0,
          fill: false, tension: 0.3, spanGaps: false, order: 2 },
        // 3 — Deviation fill: fills between real (index 1) and plan (index 0) → pink area
        { label: t('rec.lbl.desvFill'),
          data: realData, borderColor: 'transparent', borderWidth: 0,
          backgroundColor: 'rgba(220,38,38,.12)',
          pointRadius: 0, tension: 0.3, spanGaps: false,
          fill: { target: 0, above: 'transparent', below: 'rgba(220,38,38,.13)' },
          order: 5 },
        // 4 — Crossing dot: blue filled circle on Plan line where Recovery meets it
        { label: null,
          data: scView.map((_, i) => i === ti ? +(planAtTgt * 100).toFixed(2) : null),
          borderColor: '#2563eb', backgroundColor: '#2563eb',
          pointRadius: scView.map((_, i) => i === ti ? 7 : 0),
          pointHoverRadius: 9, borderWidth: 2, fill: false, tension: 0,
          showLine: false, spanGaps: false, order: 0 },
        // 5 — Current-week dot: green filled circle on Real line at fecha de corte
        { label: null,
          data: scView.map((_, i) => i === ci ? +(pctReal * 100).toFixed(2) : null),
          borderColor: '#166534', backgroundColor: '#166534',
          pointRadius: scView.map((_, i) => i === ci ? 5 : 0),
          pointHoverRadius: 7, borderWidth: 2, fill: false, tension: 0,
          showLine: false, spanGaps: false, order: 0 },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { top: 6, bottom: 4 } },
      plugins: {
        legend: {
          display: true, position: 'top',
          labels: {
            boxWidth: 28, boxHeight: 3, font: { size: 11 }, padding: 16,
            filter: item => item.text != null,
            generateLabels(chart) {
              const ds = chart.data.datasets;
              return [
                { text: ds[0].label, strokeStyle: ds[0].borderColor, lineWidth: 2,
                  fillStyle: 'transparent', hidden: false, datasetIndex: 0 },
                { text: ds[1].label, strokeStyle: ds[1].borderColor, lineWidth: 2.5,
                  fillStyle: 'transparent', hidden: false, datasetIndex: 1 },
                { text: ds[2].label, strokeStyle: ds[2].borderColor, lineWidth: 2,
                  lineDash: [8, 5], fillStyle: 'transparent', hidden: false, datasetIndex: 2 },
                { text: ds[3].label, strokeStyle: 'transparent', lineWidth: 0,
                  fillStyle: 'rgba(220,38,38,.2)', hidden: false, datasetIndex: 3 },
              ];
            }
          }
        },
        tooltip: {
          mode: 'index', intersect: false,
          callbacks: {
            title: items => items[0] ? weekLabels[items[0].dataIndex] : '',
            afterBody(items) {
              const i = items[0]?.dataIndex;
              if (i == null || i !== currIdx) return [];
              return [
                `── ${t('rec.cutoffLbl')} ──`,
                `Plan en sem.+${targetWeeks}: ${(planAtTgt * 100).toFixed(2)}%`,
                `Desvío actual: ${(deviation * 100).toFixed(2)} p.p.`,
              ];
            }
          }
        },
      },
      scales: {
        y: {
          min: 0,
          // Scale Y to the plan value at the target week + 20% headroom, rounded to next 5%
          max: Math.ceil(planAtTgt * 100 * 1.20 / 5) * 5,
          ticks: { callback: v => v + '%', font: { size: 11 } },
          grid: { color: 'rgba(0,0,0,.06)' }
        },
        x: {
          ticks: {
            callback: (_, i) => xTickLabels[i] ?? null,
            maxRotation: 90, minRotation: 45, font: { size: 9 },
            autoSkip: false,
          },
          grid: { display: false }
        },
      },
      interaction: { mode: 'index', intersect: false },
    },
  });

}

// ── S-curve expand modal ─────────────────────────────────────────────────────
function openScurveModal() {
  const modal = document.getElementById('scurveModal');
  if (!modal) return;

  // Sync filter values from main controls → modal controls
  const mainSel  = document.getElementById('recTargetWeeks');
  const modalSel = document.getElementById('recTargetWeeksModal');
  if (mainSel && modalSel) modalSel.value = mainSel.value;

  // Sync target date label
  _syncModalDate();

  modal.classList.add('open');
  // Two rAF so the modal canvas has correct dimensions before rendering
  requestAnimationFrame(() => requestAnimationFrame(() => renderRecScurve('recScurveModalChart')));
}

function closeScurveModal() {
  const modal = document.getElementById('scurveModal');
  if (!modal) return;
  modal.classList.remove('open');
  destroyChart('recScurveModalChart');
}

function _syncModalDate() {
  const src = document.getElementById('recTargetDate');
  const dst = document.getElementById('recTargetDateModal');
  if (src && dst) dst.textContent = src.textContent;
}

// ESC closes the modal
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') closeScurveModal();
});

/** Draw a rounded rectangle path (helper used by renderRecScurve annotations) */
function _rrect(ctx, x, y, w, h, r) {
  ctx.beginPath();
  ctx.moveTo(x + r, y);
  ctx.lineTo(x + w - r, y);
  ctx.quadraticCurveTo(x + w, y, x + w, y + r);
  ctx.lineTo(x + w, y + h - r);
  ctx.quadraticCurveTo(x + w, y + h, x + w - r, y + h);
  ctx.lineTo(x + r, y + h);
  ctx.quadraticCurveTo(x, y + h, x, y + h - r);
  ctx.lineTo(x, y + r);
  ctx.quadraticCurveTo(x, y, x + r, y);
  ctx.closePath();
}

/** Weekly rhythm chart: real increments (history) + required-rate reference line */
function renderRecWeekly() {
  const canvas = document.getElementById('recWeeklyChart');
  if (!canvas || !D) return;

  const sc      = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  if (currIdx < 0) return;

  const pctReal    = sc[currIdx].real  || 0;
  const planAtCurr = sc[currIdx].plan  || 0;
  const targetIdx  = Math.min(sc.length - 1, currIdx + _recTargetWeeks);
  const planAtTgt  = sc[targetIdx]?.plan || 0;
  const recNeeded  = Math.max(0, planAtTgt - pctReal);
  const reqRate    = _recTargetWeeks > 0 ? (recNeeded / _recTargetWeeks) * 100 : 0;  // % per week

  // Last _recRateWeeks real weekly increments
  const nWeeks     = Math.max(_recRateWeeks, 8);
  const sliceStart = Math.max(1, currIdx - nWeeks + 1);
  const labels     = [];
  const realInc    = [];
  const planInc    = [];

  for (let gi = sliceStart; gi <= currIdx; gi++) {
    labels.push(sc[gi].week);
    const rPrev = sc[gi - 1]?.real, rCurr = sc[gi]?.real;
    realInc.push(rPrev != null && rCurr != null
      ? +((rCurr - rPrev) * 100).toFixed(3) : null);
    const pPrev = sc[gi - 1]?.plan ?? 0;
    planInc.push(+((sc[gi].plan - pPrev) * 100).toFixed(3));
  }

  const validReal = realInc.filter(v => v != null);
  const avgReal   = validReal.length
    ? validReal.reduce((a, b) => a + b, 0) / validReal.length : 0;

  // Required-rate horizontal line (same length as bars)
  const reqLine = labels.map(() => +reqRate.toFixed(3));
  // Avg-real horizontal line
  const avgLine = labels.map(() => +avgReal.toFixed(3));

  // Footer
  const footer = document.getElementById('recWeeklyFooter');
  if (footer) {
    footer.innerHTML =
      `${t('rec.lbl.avgRate')}: <b>${avgReal.toFixed(2)}%</b>/sem` +
      ` &nbsp;|&nbsp; ` +
      `${t('rec.lbl.reqRate')}: <b style="color:#f97316">${reqRate.toFixed(2)}%</b>/sem`;
  }

  destroyChart('recWeeklyChart');
  charts['recWeeklyChart'] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [
        // Plan increment bars (background reference)
        { type: 'bar', label: t('rec.lbl.plan'), data: planInc,
          backgroundColor: 'rgba(37,99,235,.25)', borderColor: '#2563eb',
          borderWidth: 1, order: 3 },
        // Real increment bars
        { type: 'bar', label: t('rec.lbl.real'), data: realInc,
          backgroundColor: 'rgba(22,163,74,.55)', borderColor: '#16a34a',
          borderWidth: 1, order: 2 },
        // Required rate — orange horizontal line
        { type: 'line', label: t('rec.lbl.reqRate'),
          data: reqLine, borderColor: '#f97316', borderWidth: 2,
          borderDash: [6, 4], pointRadius: 0, fill: false,
          tension: 0, order: 0 },
        // Avg real rate — green dashed line
        { type: 'line', label: t('rec.lbl.avgRate'),
          data: avgLine, borderColor: '#16a34a', borderWidth: 1.5,
          borderDash: [3, 3], pointRadius: 0, fill: false,
          tension: 0, order: 1 },
      ],
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: {
          display: true, position: 'bottom',
          labels: { boxWidth: 22, boxHeight: 3, font: { size: 11 }, padding: 12 },
        },
        tooltip: {
          mode: 'index', intersect: false,
          callbacks: {
            label: item => `${item.dataset.label}: ${item.parsed.y != null ? item.parsed.y.toFixed(2) + '%' : '—'}`,
          },
        },
      },
      scales: {
        y: {
          min: 0,
          ticks: { callback: v => v.toFixed(2) + '%', font: { size: 10 } },
          grid: { color: 'rgba(0,0,0,.06)' },
        },
        x: { ticks: { maxRotation: 30, font: { size: 10 } } },
      },
    },
  });
}

// ── Render ────────────────────────────────────────────────────────────────────
function _tryRender(fn, ...args) {
  try { fn(...args); } catch(e) { console.error('[render] ' + fn.name, e); }
}

function render() {
  _tryRender(renderKPIs);
  _tryRender(renderResumen);
  _tryRender(renderResMiniScurve);
  _syncRsTargetDate();
  _tryRender(renderRsDesvAreaDonut);
  _tryRender(renderAreaChart);
  _tryRender(renderScurve);
  const _desvNeg = D.ranking.filter(r => r.desvPond < 0);
  _tryRender(renderDesviosBar, _desvNeg);
  _tryRender(renderDesviosTable, _desvNeg);
  const _critSorted = [...D.critical].sort((a, b) => a.desviacion - b.desviacion);
  _tryRender(renderCriticasBar, _critSorted);
  _tryRender(renderCriticasTable, _critSorted);
  _tryRender(renderSinAvanceCharts, D.sinAvance);
  _tryRender(renderSinAvanceTable, D.sinAvance);
  _tryRender(renderPlazos);
  const _rankNeg = D.ranking.filter(r => r.desvPond < 0);
  const _rankPos = D.ranking.filter(r => r.desvPond > 0).sort((a,b) => b.desvPond - a.desvPond);
  _tryRender(renderRankingBar,      _rankNeg.slice(0, 20));
  _tryRender(renderRankingTable,    _rankNeg);
  _tryRender(renderRankingBarPos,   _rankPos.slice(0, 15));
  _tryRender(renderRankingTablePos, _rankPos);
  _tryRender(renderConsolidado);
  _tryRender(renderRecovery);
  try { populateAreaDropdowns(); } catch(e) { console.error('[render] populateAreaDropdowns', e); }
  try { _consolTree = buildConsolTree(); } catch(e) { console.error('[render] buildConsolTree', e); }
  try { buildCascadeFilters(); } catch(e) { console.error('[render] buildCascadeFilters', e); }
  try { buildScurveCascadeFilters(); } catch(e) { console.error('[render] buildScurveCascadeFilters', e); }
  _tryRender(renderScurveFiltered);
  _tryRender(initSimTab);
  _tryRender(initArbol);
  _tryRender(renderArbol);
  _tryRender(renderFuture);
  _tryRender(renderDesviosPorAreas);
}

// ── Desvíos por Áreas ─────────────────────────────────────────────────────────
const DV_COLORS = ['#f59e0b','#22c55e','#3b82f6','#14b8a6','#8b5cf6','#06b6d4','#ec4899','#eab308','#ef4444','#a78bfa'];

function renderDesviosPorAreas() {
  if (!D) return;

  const areas = D.areas.filter(a => a.nivel === 3 && a.incidencia > 0.001);
  const totalRow = D.allRecords.find(r => r.edt === '4.5');
  const currIdx = D.scurve.findIndex(s => s.isCurrent);

  // ── KPIs ──────────────────────────────────────────────────────────────────
  const totalDesvPond = areas.reduce((s, a) => s + (a.desvPond || 0), 0);
  const totalPP = totalDesvPond * 100;
  const pctPlan = totalRow ? totalRow.pctCompPlan * 100 : 0;
  const pctReal = totalRow ? totalRow.pctCompReal * 100 : 0;
  const areasNeg = areas.filter(a => (a.pctCompReal - a.pctCompPlan) < 0);

  const dvKpiTotal = document.getElementById('dvKpiTotal');
  const dvKpiPlan  = document.getElementById('dvKpiPlan');
  const dvKpiReal  = document.getElementById('dvKpiReal');
  const dvKpiAreas = document.getElementById('dvKpiAreas');
  const dvDateBadge = document.getElementById('dvDateBadge');

  if (dvDateBadge && D.meta.dataDate) dvDateBadge.textContent = t('dv.cutDate') + ': ' + fmtDate(D.meta.dataDate);
  if (dvKpiTotal) {
    dvKpiTotal.textContent = (totalPP >= 0 ? '+' : '') + totalPP.toFixed(2) + ' p.p.';
    dvKpiTotal.className = 'dv-kpi-val ' + (totalPP < 0 ? 'neg' : 'pos');
  }
  if (dvKpiPlan) dvKpiPlan.textContent = pctPlan.toFixed(2) + '%';
  if (dvKpiReal) dvKpiReal.textContent = pctReal.toFixed(2) + '%';
  if (dvKpiAreas) dvKpiAreas.textContent = areasNeg.length + ' / ' + areas.length;

  // ── Build area data with desvio_pp and impacto ─────────────────────────────
  // Previous week's deviation: find last week before currIdx where real > 0 for total
  const prevIdx = (function() {
    for (let i = currIdx - 1; i >= 0; i--) {
      if (D.scurve[i].real != null && D.scurve[i].real > 0) return i;
    }
    return -1;
  })();

  const areaData = areas.map((a, i) => {
    const desvPP = (a.pctCompReal - a.pctCompPlan) * 100;
    const impacto = desvPP * a.incidencia; // p.p. × (fraction 0-1) = p.p. contribution

    // Trend: compare current desvio_pp vs prev week
    let prevDesvPP = desvPP;
    if (prevIdx >= 0 && a.planSeries && a.realSeries) {
      const prevPlan = a.planSeries[prevIdx] || 0;
      const prevReal = a.realSeries[prevIdx] || 0;
      // Only use prevReal if it had real data
      if (prevReal > 0) {
        prevDesvPP = (prevReal - prevPlan) * 100;
      }
    }
    const trend = desvPP > prevDesvPP ? 'up' : desvPP < prevDesvPP ? 'down' : 'neutral';

    return { area: a, desvPP, impacto, trend, color: DV_COLORS[i % DV_COLORS.length] };
  });

  // Sort by impacto (most negative first)
  areaData.sort((a, b) => a.impacto - b.impacto);

  // ── Ranking table ──────────────────────────────────────────────────────────
  const totalImpacto    = areaData.reduce((s, d) => s + d.impacto, 0);
  const totalDesvPP2    = totalRow ? (totalRow.pctCompReal - totalRow.pctCompPlan) * 100 : 0;
  const totalIncidPlan  = areaData.reduce((s, d) => s + d.area.incidencia * d.area.pctCompPlan, 0) * 100;
  const totalIncidReal  = areaData.reduce((s, d) => s + d.area.incidencia * d.area.pctCompReal, 0) * 100;

  const trendHtml = (trend) => {
    if (trend === 'up')   return '<span class="dv-trend-up">&#8599;</span>';
    if (trend === 'down') return '<span class="dv-trend-down">&#8600;</span>';
    return '<span class="dv-trend-neu">→</span>';
  };

  const tableRows = areaData.map(d => {
    const incidPlan    = d.area.incidencia * d.area.pctCompPlan * 100;
    const incidReal    = d.area.incidencia * d.area.pctCompReal * 100;
    const realCls      = devClass(d.area.pctCompReal - d.area.pctCompPlan);
    const incidRealCls = devClass(incidReal - incidPlan);
    const impCls       = devClass(d.impacto);
    const desvCls      = devClass(d.desvPP);
    return `<tr>
    <td class="left" style="max-width:160px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis" title="${d.area.tarea.trim()}">${d.area.tarea.trim()}</td>
    <td>${(d.area.pctCompPlan * 100).toFixed(1)}%</td>
    <td class="${realCls}">${(d.area.pctCompReal * 100).toFixed(1)}%</td>
    <td class="${desvCls}">${(d.desvPP >= 0 ? '+' : '') + d.desvPP.toFixed(2)}%</td>
    <td>${(d.area.incidencia * 100).toFixed(3)}%</td>
    <td>${incidPlan.toFixed(3)}%</td>
    <td class="${incidRealCls}">${incidReal.toFixed(3)}%</td>
    <td class="${impCls}">${(d.impacto >= 0 ? '+' : '') + d.impacto.toFixed(3)}%</td>
    <td>${trendHtml(d.trend)}</td>
  </tr>`;
  }).join('');

  const totalRow2Cls = totalDesvPP2 < 0 ? 'dev-neg' : totalDesvPP2 > 0 ? 'dev-pos' : 'dev-neutral';
  const totalImpCls  = totalImpacto < 0 ? 'dev-neg' : totalImpacto > 0 ? 'dev-pos' : 'dev-neutral';

  const totalRealCls     = totalRow ? devClass(totalRow.pctCompReal - totalRow.pctCompPlan) : 'dev-neutral';
  const totalIncidRealCls = devClass(totalIncidReal - totalIncidPlan);

  const totalHtml = `<tr class="dv-total-row">
    <td class="left"></td>
    <td><strong>${totalRow ? (totalRow.pctCompPlan * 100).toFixed(1) + '%' : '—'}</strong></td>
    <td class="${totalRealCls}"><strong>${totalRow ? (totalRow.pctCompReal * 100).toFixed(1) + '%' : '—'}</strong></td>
    <td class="${totalRow2Cls}"><strong>${(totalDesvPP2 >= 0 ? '+' : '') + totalDesvPP2.toFixed(2)}%</strong></td>
    <td>—</td>
    <td><strong>${totalIncidPlan.toFixed(3)}%</strong></td>
    <td class="${totalIncidRealCls}"><strong>${totalIncidReal.toFixed(3)}%</strong></td>
    <td class="${totalImpCls}"><strong>${(totalImpacto >= 0 ? '+' : '') + totalImpacto.toFixed(3)}%</strong></td>
    <td>—</td>
  </tr>`;

  const dvRankingTable = document.getElementById('dvRankingTable');
  if (dvRankingTable) {
    dvRankingTable.innerHTML = tableWrap(
      `<tr>
        <th class="left">${t('th.area')}</th>
        <th>${t('th.pctPlan')}</th>
        <th>${t('th.pctActual')}</th>
        <th>% Desvío</th>
        <th>Incid Total</th>
        <th>Incid Plan</th>
        <th>Incid Real</th>
        <th>Incid Desvío</th>
        <th>${t('th.trend')}</th>
      </tr>`,
      tableRows + totalHtml
    );
  }

  // ── Horizontal bar chart ───────────────────────────────────────────────────
  const chartLabels = areaData.map(d => d.area.tarea.trim().slice(0, 28));
  const chartValues = areaData.map(d => +d.impacto.toFixed(4));
  const chartColors = areaData.map(d => d.color);

  destroyChart('dvImpactChart');
  const dvImpactCanvas = document.getElementById('dvImpactChart');
  if (dvImpactCanvas) {
    const chartHeight = Math.max(180, areaData.length * 38);
    dvImpactCanvas.parentElement.style.height = chartHeight + 'px';

    charts['dvImpactChart'] = new Chart(dvImpactCanvas, {
      type: 'bar',
      data: {
        labels: chartLabels,
        datasets: [{
          label: 'Impacto (p.p.)',
          data: chartValues,
          backgroundColor: chartColors,
          borderRadius: 3,
        }]
      },
      options: {
        indexAxis: 'y',
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: { display: false },
          tooltip: {
            callbacks: {
              label: ctx => {
                const v = ctx.parsed.x;
                return 'Impacto: ' + (v >= 0 ? '+' : '') + v.toFixed(3) + ' p.p.';
              }
            }
          },
          datalabels: { display: false },
        },
        scales: {
          x: { ticks: { callback: v => (v >= 0 ? '+' : '') + v.toFixed(2) } },
          y: { ticks: { font: { size: 11 } } }
        }
      }
    });

    const footerEl = document.getElementById('dvImpactFooter');
    if (footerEl) {
      footerEl.textContent = 'Impacto total das áreas: ' + (totalImpacto >= 0 ? '+' : '') + totalImpacto.toFixed(3) + ' p.p.';
    }
  }


  // ── Status table ──────────────────────────────────────────────────────────
  const statusRows = areaData.map(d => {
    let label, cls;
    if (d.desvPP >= 0) {
      label = t('dv.onTime'); cls = 'dv-status-ok';
    } else if (d.trend === 'up') {
      label = t('dv.recovering'); cls = 'dv-status-rec';
    } else {
      label = t('dv.late'); cls = 'dv-status-bad';
    }
    return `<tr>
      <td class="left" title="${d.area.tarea.trim()}">${d.area.tarea.trim().length > 36 ? d.area.tarea.trim().slice(0, 36) + '…' : d.area.tarea.trim()}</td>
      <td><span class="${cls}">${label}</span></td>
    </tr>`;
  }).join('');

  const dvStatusTable = document.getElementById('dvStatusTable');
  if (dvStatusTable) {
    dvStatusTable.innerHTML = tableWrap(
      `<tr><th class="left">${t('th.area')}</th><th>${t('th.situation')}</th></tr>`,
      statusRows
    );
  }
}

// ── KPIs ──────────────────────────────────────────────────────────────────────
// ── Forecast computation (shared between KPIs and Resumo) ─────────────────────
function _computeForecast() {
  if (!D) return { spi: null, trend: '—', base: '—', wksOver: null, R: null, P: null, projAtEnd: null, extraRate: null };
  const _loc = _lang === 'zh' ? 'zh-CN' : _lang === 'en' ? 'en-US' : 'es-CL';
  const sc = D.scurve;
  const realPts = sc.filter(s => s.real != null && s.real > 0);
  let forecastTrend = '—', forecastSPI = null, forecastWksOver = null;
  let diagR = null, diagP = null, diagProjAtEnd = null, diagExtraRate = null;

  if (realPts.length >= 1) {
    const lastRealPt = realPts[realPts.length - 1];
    const R = lastRealPt.real;
    const currPtIdx = sc.indexOf(lastRealPt);
    const P = sc[currPtIdx]?.plan || 0;
    diagR = R; diagP = P;

    if (R > 0 && P > 0 && R < 1) {
      forecastSPI = R / P;                              // IDP = real ÷ planejado
      let projected = R;
      let forecastDate = null;

      // ── Fase 1: dentro do prazo base — IDP × incremento planejado ──────────
      for (let i = currPtIdx + 1; i < sc.length; i++) {
        const inc = (sc[i].plan || 0) - (sc[i - 1].plan || 0);
        projected += inc * forecastSPI;
        if (projected >= 1.0) {
          forecastDate = new Date((sc[i].date || sc[sc.length - 1].date) + 'T12:00:00');
          forecastWksOver = 0;
          break;
        }
      }
      diagProjAtEnd = projected;                        // % real projetado no fim do prazo base

      // ── Fase 2: além do prazo base ────────────────────────────────────────
      // Usa a mesma lógica da Fase 1 (IDP × taxa planejada), mas com a taxa
      // planejada MÉDIA de todo o trecho restante — sem janela arbitrária.
      // Equivale a: ritmo real acumulado ÷ semanas totais com dados.
      if (!forecastDate) {
        let weeklyRate = 0;

        // IDP × taxa planejada média de todo o período restante
        // (consistente com Fase 1 e não depende de nenhuma janela)
        const wksLeft  = sc.length - currPtIdx - 1;
        const workLeft = (sc[sc.length - 1].plan || 0) - P;
        if (wksLeft > 0 && workLeft > 0) {
          weeklyRate = (workLeft / wksLeft) * forecastSPI;
        }

        // Fallback: taxa real média de toda a história disponível
        if (weeklyRate <= 0 && realPts.length > 0) {
          weeklyRate = realPts[realPts.length - 1].real / realPts.length;
        }

        diagExtraRate = weeklyRate;
        if (weeklyRate > 0) {
          const weeksExtra = Math.ceil((1 - projected) / weeklyRate);
          forecastWksOver  = weeksExtra;
          const base = new Date(sc[sc.length - 1].date + 'T12:00:00');
          base.setDate(base.getDate() + weeksExtra * 7);
          forecastDate = base;
        }
      }

      if (forecastDate) {
        forecastTrend = forecastDate.toLocaleDateString(_loc, { month: 'short', year: 'numeric' });
      }
    }
  }

  const endLB = D.meta.endLB;
  let forecastBase = '—';
  if (endLB) {
    forecastBase = new Date(endLB + 'T12:00:00').toLocaleDateString(_loc, { month: 'short', year: 'numeric' });
  }
  return {
    spi: forecastSPI, trend: forecastTrend, base: forecastBase, wksOver: forecastWksOver,
    R: diagR, P: diagP, projAtEnd: diagProjAtEnd, extraRate: diagExtraRate,
  };
}

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
  if (devEl) {
    devEl.classList.toggle('kpi-alert', m.desvio < 0);
    devEl.classList.toggle('positive',  m.desvio >= 0);
  }

  // ── Dual forecast ─────────────────────────────────────────────────────────
  const _loc = _lang === 'zh' ? 'zh-CN' : _lang === 'en' ? 'en-US' : 'es-CL';
  const fc = _computeForecast();
  const { spi: forecastSPI, trend: forecastTrend, base: forecastBase, wksOver: forecastWksOver } = fc;

  set('kForecastTrend', forecastTrend);
  const baseEl = document.getElementById('kForecastBase');
  if (baseEl) baseEl.textContent = forecastBase;

  // Build balloon HTML — with full calculation trace
  const spiStr  = forecastSPI != null ? forecastSPI.toFixed(3) : '—';
  const wksStr  = forecastWksOver != null
    ? (forecastWksOver === 0
        ? `✔ ${t('kpi.withinPlan')}`
        : `+${forecastWksOver.toLocaleString()} ${t('kpi.trendWeeks')}`)
    : '—';
  const traceRows = fc.R != null ? `
    <div class="kpi-balloon-trace">
      <div>% Real: <b>${pct(fc.R)}</b> &nbsp;|&nbsp; % Plan: <b>${pct(fc.P)}</b></div>
      ${fc.projAtEnd != null ? `<div>Proj. no fim LB: <b>${pct(fc.projAtEnd)}</b></div>` : ''}
      ${fc.extraRate  != null && forecastWksOver > 0 ? `<div>Taxa extrapol.: <b>${(fc.extraRate * 100).toFixed(3)}%/sem</b></div>` : ''}
    </div>` : '';
  const balloonHtml = `
    <div class="kpi-balloon-section">
      <div class="kpi-balloon-title">📈 ${t('kpi.trendProj')}</div>
      <div class="kpi-balloon-val">${forecastTrend}</div>
      <div class="kpi-balloon-desc">
        <strong>${t('kpi.spiLabel')} = ${spiStr}</strong> <span style="opacity:.7">(${t('kpi.spiDesc')})</span><br>
        ${t('kpi.trendMethod')}<br>
        <strong>${wksStr}</strong>
      </div>
      ${traceRows}
    </div>
    <hr class="kpi-balloon-divider">
    <div class="kpi-balloon-section">
      <div class="kpi-balloon-title">📋 ${t('kpi.contractDate')}</div>
      <div class="kpi-balloon-val">${forecastBase}</div>
      <div class="kpi-balloon-desc">${t('kpi.baselineDesc')}</div>
    </div>`;
  const balloonEl = document.getElementById('kForecastBalloon');
  if (balloonEl) balloonEl.innerHTML = balloonHtml;

  // Bind hover/click once
  const kCard = document.getElementById('kForecastCard');
  if (kCard && !kCard._balloonBound) {
    kCard._balloonBound = true;
    let pinned = false;
    kCard.addEventListener('mouseenter', () => {
      document.getElementById('kForecastBalloon')?.classList.add('visible');
    });
    kCard.addEventListener('mouseleave', () => {
      if (!pinned) document.getElementById('kForecastBalloon')?.classList.remove('visible');
    });
    kCard.addEventListener('click', e => {
      e.stopPropagation();
      pinned = !pinned;
      const b = document.getElementById('kForecastBalloon');
      if (b) b.classList.toggle('visible', pinned);
    });
    document.addEventListener('click', () => {
      pinned = false;
      document.getElementById('kForecastBalloon')?.classList.remove('visible');
    });
  }
}

// ── Resumen ejecutivo ─────────────────────────────────────────────────────────
function renderResumen() {
  if (!document.getElementById('rsMetaFecha')) return;
  const m = D.meta;
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  const pctReal = currIdx >= 0 ? (sc[currIdx].real || 0) : (m.pctReal || 0);
  const pctPlan = m.pctPlan || 0;
  const dev = m.desvio; // already = pctReal - pctPlan

  // Recovery figures (same as Recovery tab)
  const targetIdx = Math.min(sc.length - 1, currIdx + _recTargetWeeks);
  const planAtTarget = sc[targetIdx]?.plan || 0;
  const recNeeded = Math.max(0, planAtTarget - pctReal);
  const reqRate   = _recTargetWeeks > 0 ? recNeeded / _recTargetWeeks : 0;
  const recentRate = _recRecentRate(sc, currIdx, _rsRateWeeks);
  const accelFactor = recentRate > 0 ? reqRate / recentRate : null;
  const optRate  = reqRate * 1.2;
  const optWeeks = optRate > 0 ? Math.ceil(recNeeded / optRate) : null;

  // Header metadata
  _rsSet('rsMetaFecha',  fmtDate(m.dataDate));
  _rsSet('rsMetaSemana', m.dataWeek || '—');

  // 6 KPI cards
  _rsSet('rsKpiPlan',   pct(pctPlan));
  _rsSet('rsKpiReal',   pct(pctReal));
  _rsSet('rsKpiDesv',   (dev >= 0 ? '+' : '') + pct(dev));
  _rsSet('rsKpiRecReq', (reqRate * 100).toFixed(3) + '%');
  _rsSet('rsKpiAccel',  accelFactor != null ? accelFactor.toFixed(2) + 'x' : '—');
  _rsSet('rsKpiOptim',  optWeeks != null ? optWeeks + ' sem' : '—');

  // Analysis panel
  _rsRenderAnalysis(pctPlan, pctReal, dev, planAtTarget, recNeeded, reqRate, recentRate, accelFactor);

  // Top 5 desvíos
  _rsRenderTop5();

  // Top 5 por impacto
  _rsRenderImpact();

  // Palancas de recuperación
  _rsRenderLevers();

  // Mensaje ejecutivo
  _rsRenderMessage(dev, pctReal, recNeeded, reqRate, accelFactor);
}

function _rsSet(id, val) {
  const el = document.getElementById(id);
  if (el) el.textContent = val;
}

function _rsRenderAnalysis(pctPlan, pctReal, dev, planAtTarget, recNeeded, reqRate, recentRate, accelFactor) {
  const el = document.getElementById('rsAnalysisPanel');
  if (!el) return;
  const rows = [
    { lbl: '% Plan acumulado',          val: pct(pctPlan) },
    { lbl: '% Real acumulado',          val: pct(pctReal) },
    { lbl: 'Desvío acumulado',          val: (dev >= 0?'+':'')+pct(dev),        cls: dev < 0 ? 'rs-val-neg' : 'rs-val-ok' },
    { lbl: `Plan en sem. +${_recTargetWeeks}`, val: pct(planAtTarget) },
    { lbl: 'Recuperación necesaria',    val: pct(recNeeded),                     highlight: true },
    { lbl: 'Tasa req. (% / sem)',       val: (reqRate*100).toFixed(2)+'%',       highlight: true },
    { lbl: `Tasa reciente (${_rsRateWeeks} sem)`, val: (recentRate*100).toFixed(2)+'%' },
    { lbl: 'Factor aceleración',        val: accelFactor!=null ? accelFactor.toFixed(2)+'x' : '—',
      cls: accelFactor != null ? (accelFactor > 2 ? 'rs-val-neg' : accelFactor > 1.2 ? 'rs-val-warn' : 'rs-val-ok') : '' },
  ];
  el.innerHTML = rows.map(r => `
    <div class="rs-analysis-row${r.highlight ? ' rs-analysis-highlight' : ''}">
      <span class="rs-analysis-lbl">${r.lbl}</span>
      <span class="rs-analysis-val${r.highlight ? ' rs-val-primary' : r.cls ? ' '+r.cls : ''}">${r.val}</span>
    </div>`).join('');
}

function _rsRenderTop5() {
  const el = document.getElementById('rsTop5Panel');
  if (!el) return;
  const areaMap = _buildAreaMap();
  // Use D.critical (same pool as Críticas tab: incidencia > 0.003 && desviacion < -0.05)
  // Sorted by desviacion ascending (most negative first) — matches Críticas grid
  const byDesv = [...(D.critical || [])].sort((a, b) => a.desviacion - b.desviacion);
  const top5 = byDesv.slice(0, 5)
    .sort((a, b) => _top5SortAsc
      ? a.desviacion - b.desviacion     // ↑ Mayor primero: most negative deviation first
      : b.desviacion - a.desviacion);   // ↓ Menor primero: least negative deviation first
  el.innerHTML = top5.map((r, i) => {
    const name = r.tarea.trim();
    const short = name.length > 40 ? name.slice(0, 40) + '…' : name;
    const area  = _areaOfEdt(r.edt, areaMap);
    return `
      <div class="rs-top5-item">
        <span class="rs-top5-rank">${i+1}</span>
        <div class="rs-top5-info">
          <div class="rs-top5-name" title="${name}">${short}</div>
          <div class="rs-top5-meta">${r.edt} · ${area} · Incid: ${pct(r.incidencia,3)}</div>
        </div>
        <span class="rs-top5-dev">${signPct(r.desviacion)}</span>
      </div>`;
  }).join('');
}

function _rsRenderImpact() {
  const el = document.getElementById('rsImpactPanel');
  if (!el) return;
  const areaMap = _buildAreaMap();
  // Fixed selection: top 5 by |desvPond| (most impactful negative activities)
  // D.ranking is already sorted by |desvPond| desc, so just filter neg and slice
  const top5 = D.ranking.filter(r => r.desvPond < 0).slice(0, 5);
  el.innerHTML = top5.map((r, i) => {
    const name   = r.tarea.trim();
    const short  = name.length > 38 ? name.slice(0, 38) + '…' : name;
    const area   = _areaOfEdt(r.edt, areaMap);
    const desvPP = (r.desvPond * 100).toFixed(3); // impact in p.p.
    return `
      <div class="rs-top5-item">
        <span class="rs-top5-rank">${i+1}</span>
        <div class="rs-top5-info">
          <div class="rs-top5-name" title="${name}">${short}</div>
          <div class="rs-top5-meta">${r.edt} · ${area} · Incid: ${pct(r.incidencia,3)}</div>
        </div>
        <div class="rs-imp-vals">
          <span class="rs-imp-dev">${signPct(r.desviacion)}</span>
          <span class="rs-imp-pond">Imp: <b>${desvPP} p.p.</b></span>
        </div>
      </div>`;
  }).join('');
}

function _rsRenderLevers() {
  const el = document.getElementById('rsLeversPanel');
  if (!el) return;
  const levers = D.ranking.filter(r => r.desvPond < 0).slice(0, 5);
  const areaMap = _buildAreaMap();
  el.innerHTML = levers.map((r, i) => {
    const name  = r.tarea.trim();
    const short = name.length > 38 ? name.slice(0, 38) + '…' : name;
    const upside = pct(Math.abs(r.desvPond || 0), 3);
    return `
      <div class="rs-lever-item">
        <span class="rs-lever-rank">${i+1}</span>
        <div class="rs-lever-info">
          <div class="rs-lever-name" title="${name}">${short}</div>
          <div class="rs-lever-meta">${r.edt} · Potencial: <b>${upside}</b></div>
        </div>
        <span class="rs-lever-upside">+${upside}</span>
      </div>`;
  }).join('');
}

function _rsRenderMessage(dev, pctReal, recNeeded, reqRate, accelFactor) {
  const el = document.getElementById('rsMessagePanel');
  if (!el) return;
  const isCrit = dev < -0.02;
  const isLate = dev < -0.005;

  const bullets = [];
  if (isCrit) {
    bullets.push({ type:'danger', text:`El proyecto acumula un <b>desvío crítico de ${pct(Math.abs(dev))}</b> respecto al plan.` });
  } else if (isLate) {
    bullets.push({ type:'warn', text:`El proyecto presenta un <b>desvío de ${pct(Math.abs(dev))}</b> respecto al plan acumulado.` });
  } else {
    bullets.push({ type:'ok', text:`El proyecto se encuentra <b>dentro del plan</b> con avance real de ${pct(pctReal)}.` });
  }

  if (recNeeded > 0.0001) {
    bullets.push({ type:'warn', text:`Para alcanzar la meta en <b>${_recTargetWeeks} semanas</b>, se requiere recuperar <b>${pct(recNeeded)}</b> (${(reqRate*100).toFixed(3)}% / sem).` });
  }

  if (accelFactor != null && accelFactor > 1.05) {
    const t2 = accelFactor > 2 ? 'danger' : 'warn';
    bullets.push({ type:t2, text:`Se debe acelerar el ritmo de trabajo <b>${accelFactor.toFixed(1)}x</b> respecto a la tasa reciente de las últimas ${_recRateWeeks} semanas.` });
  } else if (accelFactor != null && accelFactor <= 1.05) {
    bullets.push({ type:'ok', text:`La tasa reciente de avance es suficiente para alcanzar la meta (factor ≈ ${accelFactor != null ? accelFactor.toFixed(2) : '—'}x).` });
  }

  const worstArea = D.areas
    .filter(a => a.nivel === 3 && a.desvPond < -0.005)
    .sort((a, b) => a.desvPond - b.desvPond)[0];
  if (worstArea) {
    bullets.push({ type:'danger', text:`Área más crítica: <b>${worstArea.tarea.trim()}</b> — desvío ponderado ${signPct(worstArea.desvPond)}.` });
  }

  el.innerHTML = bullets.map(b => `
    <div class="rs-message-bullet">
      <div class="rs-msg-dot rs-msg-dot-${b.type}"></div>
      <div class="rs-message-text">${b.text}</div>
    </div>`).join('');
}

// ── Resumen S-curve: delegates to the shared renderRecScurve engine ────────────
function renderResMiniScurve(canvasId) {
  canvasId = canvasId || 'resMiniScurve';
  _syncRsTargetDate();
  renderRecScurve(canvasId, _rsTargetWeeks);
}

// ── Sync Resumen S-curve target date labels ───────────────────────────────────
function _syncRsTargetDate() {
  if (!D) return;
  const sc = D.scurve;
  const currIdx = sc.findIndex(s => s.isCurrent);
  if (currIdx < 0) return;
  const targetIdx = Math.min(sc.length - 1, currIdx + _rsTargetWeeks);
  const week = sc[targetIdx]?.week || '—';
  const el1 = document.getElementById('rsScurveTargetDate');
  const el2 = document.getElementById('rsScurveTargetDateModal');
  if (el1) el1.textContent = week;
  if (el2) el2.textContent = week;
}

// ── Resumen S-curve modal open/close ──────────────────────────────────────────
function openRsScurveModal() {
  const modal = document.getElementById('rsScurveModal');
  if (!modal) return;
  const mainSel  = document.getElementById('rsScurveTargetWeeks');
  const modalSel = document.getElementById('rsScurveTargetWeeksModal');
  if (mainSel && modalSel) modalSel.value = mainSel.value;
  _syncRsTargetDate();
  modal.classList.add('open');
  requestAnimationFrame(() => requestAnimationFrame(() => renderResMiniScurve('rsScurveModalChart')));
}
function closeRsScurveModal() {
  document.getElementById('rsScurveModal')?.classList.remove('open');
  destroyChart('rsScurveModalChart');
}

// ── Deviation-by-area donut for Resumen tab ───────────────────────────────────
function renderRsDesvAreaDonut() {
  const canvas = document.getElementById('rsDesvAreaDonut');
  if (!canvas || !D) return;

  const areas = D.areas
    .filter(a => a.nivel === 3 && a.incidencia > 0.001 && a.desvPond < -0.001)
    .sort((a, b) => a.desvPond - b.desvPond); // most negative first

  destroyChart('rsDesvAreaDonut');

  if (!areas.length) {
    const tableEl = document.getElementById('rsDesvAreaTable');
    if (tableEl) tableEl.innerHTML = '<div style="font-size:11px;color:var(--text-muted);padding:8px 0">Sin desvíos negativos por área.</div>';
    return;
  }

  const labels = areas.map(a => { const n = a.tarea.trim(); return n.length > 22 ? n.slice(0,22)+'…' : n; });
  const values = areas.map(a => +Math.abs(a.desvPond * 100).toFixed(3));
  const palette = ['#c00000','#e05555','#f87171','#fca5a5','#fecaca'];

  charts['rsDesvAreaDonut'] = new Chart(canvas, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{ data: values, backgroundColor: palette.slice(0, areas.length), borderWidth:2, borderColor:'#fff' }]
    },
    options: {
      responsive: true, maintainAspectRatio: false,
      cutout: '60%',
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => `${ctx.label}: ${ctx.parsed.toFixed(3)}%` } }
      }
    }
  });

  const tableEl = document.getElementById('rsDesvAreaTable');
  if (tableEl) {
    tableEl.innerHTML = areas.map(a => {
      const name  = a.tarea.trim();
      const short = name.length > 26 ? name.slice(0, 26) + '…' : name;
      return `<div class="rs-desv-row">
        <span class="rs-desv-name" title="${name}">${short}</span>
        <span class="rs-desv-val rs-desv-neg">${signPct(a.desvPond)}</span>
      </div>`;
    }).join('');
  }
}

// ── Area bar chart ────────────────────────────────────────────────────────────
function renderAreaChart() {
  const canvas = document.getElementById('areaChart');
  if (!canvas) return;
  const areas = D.areas.filter(a => a.nivel===3 && a.incidencia > 0.001);
  const plan  = areas.map(a => +(a.pctPlan*100).toFixed(2));
  const real  = areas.map(a => +(a.pctReal*100).toFixed(2));

  destroyChart('areaChart');
  charts['areaChart'] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels: areas.map(a => a.tarea.trim()),
      datasets: [
        { label: t('sc.plan'), data:plan, backgroundColor:'rgba(0,84,166,0.7)', borderRadius:3 },
        { label: t('sc.actual'), data:real, backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
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
          label: t('sc.plan'),
          data: planData,
          borderColor: '#0054a6',
          backgroundColor: 'rgba(0,84,166,.07)',
          pointRadius: 0, fill: true, tension: 0.3, borderWidth: 2,
          yAxisID: 'y'
        },
        {
          label: t('sc.actual'),
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
              return [t('th.deviation') + ': ' + (dev >= 0 ? '+' : '') + dev.toFixed(2) + '%'];
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
  if (!document.getElementById('scurveChart')) return;
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
  if (!document.getElementById('scurveFilteredChart')) return;

  const filtEdt  = _scurveFilterEdt;
  const leaves   = _scurveLeavesForEdt(filtEdt);
  const totalInc = leaves.reduce((s, r) => s + (r.incidencia || 0), 0);

  // Label in subtitle — prefer _consolTree name (shows virtual disc labels too)
  const filtLabel = filtEdt
    ? (_consolTree?.find(r => r.edt === filtEdt)?.tarea?.trim()
        || D.allRecords.find(r => r.edt === filtEdt)?.tarea?.trim()
        || filtEdt)
    : t('dv.total');
  const labelEl = document.getElementById('scurveFilterLabel');
  if (labelEl) labelEl.textContent = (filtEdt ? t('th.area') + ': ' : '') + filtLabel;

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
  sel.innerHTML = `<option value="">${t('cr.allAreas')}</option>`
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

  const areasBarCanvas = document.getElementById('areasBarChart');
  if (!areasBarCanvas) return;
  destroyChart('areasBarChart');
  charts['areasBarChart'] = new Chart(areasBarCanvas, {
    type: 'bar',
    data: {
      labels: rows.map(a => a.tarea.trim().slice(0,35)),
      datasets: [
        { label: t('th.pctPlan'),   data: rows.map(a => +(a.pctCompPlan*100).toFixed(2)), backgroundColor:'rgba(0,84,166,0.7)',  borderRadius:3 },
        { label: t('th.pctActual'), data: rows.map(a => +(a.pctCompReal*100).toFixed(2)), backgroundColor:'rgba(0,163,108,0.7)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins: { legend:{ position:'top' } },
      scales: { x:{ min:0, max:100, ticks:{ callback: v => v+'%' } } }
    }
  });
}

function renderAreasTable() {
  const areasTableEl = document.getElementById('areasTable');
  if (!areasTableEl) return;
  const n = parseInt(document.getElementById('areaLvlBox')?.value || '3');
  const rows = D.areas.filter(a => a.nivel === n);

  areasTableEl.innerHTML = tableWrap(
    `<tr>
      <th class="left">${t('th.areaGroup')}</th><th>${t('th.edt')}</th><th>${t('th.start')}</th><th>${t('th.end')}</th>
      <th>${t('th.hh')}</th><th>${t('th.incidence')}</th><th>${t('th.pctPlanPond')}</th><th>${t('th.pctRealPond')}</th>
      <th>${t('th.desvPond')}</th><th>${t('th.pctCompPlan')}</th><th>${t('th.pctCompReal')}</th><th>${t('th.status')}</th>
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
  const canvas = document.getElementById('desviosBarChart');
  if (!canvas) return;
  const top  = rows.slice(0,15);
  const vals = top.map(r => +(r.desviacion*100).toFixed(2));

  destroyChart('desviosBarChart');
  charts['desviosBarChart'] = new Chart(canvas, {
    type:'bar',
    data: {
      labels: top.map(r => r.edt),
      datasets:[{ label: t('th.deviation') + ' (%)', data:vals,
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
  const el = document.getElementById('desviosTable');
  if (!el) return;
  el.innerHTML = tableWrap(
    `<tr><th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th><th>${t('th.start')}</th><th>${t('th.end')}</th>
     <th>${t('th.hh')}</th><th>${t('th.incidence')}</th><th>${t('th.pctPlan')}</th><th>${t('th.pctActual')}</th><th>${t('th.deviation')}</th></tr>`,
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
/** Build a level-3 EDT → name lookup from D.areas */
function _buildAreaMap() {
  const map = {};
  (D?.areas || []).filter(a => a.nivel === 3).forEach(a => { map[a.edt] = a.tarea.trim(); });
  return map;
}
/** Return the nivel-3 parent name for an activity EDT (e.g. "4.5.4.1.1" → "4.5.4" name) */
function _areaOfEdt(edt, areaMap) {
  const key = edt.split('.').slice(0, 3).join('.');
  return areaMap[key] || key;
}

function renderCriticasBar(rows) {
  const criticasBarCanvas = document.getElementById('criticasBarChart');
  if (!criticasBarCanvas) return;
  const top     = rows.slice(0, 15);
  const areaMap = _buildAreaMap();
  const truncName = r => {
    const name = r.tarea.trim();
    return name.length > 40 ? name.slice(0, 38) + '…' : name;
  };
  destroyChart('criticasBarChart');
  charts['criticasBarChart'] = new Chart(criticasBarCanvas, {
    type:'bar',
    data: {
      labels: top.map(truncName),
      datasets:[
        { label: t('cr.plan'),   data: top.map(r => +(r.pctCompPlan*100).toFixed(1)), backgroundColor:'rgba(0,84,166,0.6)',  borderRadius:3 },
        { label: t('cr.actual'), data: top.map(r => +(r.pctCompReal*100).toFixed(1)), backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 },
      ]
    },
    options: {
      indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{ position:'top' },
        tooltip:{ callbacks:{
          afterTitle: items => {
            const r = top[items[0].dataIndex];
            return r ? `${r.edt}  ·  ${_areaOfEdt(r.edt, areaMap)}` : '';
          }
        }}
      },
      scales:{
        x:{ ticks:{ callback: v => v+'%' }, max:100 },
        y:{ ticks:{ font:{ size:11 } } }
      }
    }
  });
}

function renderCriticasTable(rows) {
  _criticasRows = rows;
  const criticasTableEl = document.getElementById('criticasTable');
  if (!criticasTableEl) return;
  const areaMap = _buildAreaMap();
  const totIncidPlan = rows.reduce((s, r) => s + r.incidencia * r.pctCompPlan, 0);
  const totIncidReal = rows.reduce((s, r) => s + r.incidencia * r.pctCompReal, 0);
  const totIncidDesv = rows.reduce((s, r) => s + r.incidencia * (r.desviacion || 0), 0);
  const totalRow = `<tr class="dv-total-row">
    <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
    <td></td>
    <td><strong>${pct(totIncidPlan, 4)}</strong></td>
    <td class="${devClass(totIncidReal - totIncidPlan)}"><strong>${pct(totIncidReal, 4)}</strong></td>
    <td class="${devClass(totIncidDesv)}"><strong>${signPct(totIncidDesv, 4)}</strong></td>
  </tr>`;
  criticasTableEl.innerHTML = tableWrap(
    `<tr>
      <th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th>
      <th class="left">${t('th.area')}</th>
      <th>${t('th.start')}</th><th>${t('th.end')}</th>
      <th>${t('th.hh')}</th>
      <th>${t('th.pctPlan')}</th><th>${t('th.pctActual')}</th><th>% Desvío</th>
      <th>Incid Total</th><th>Incid Plan</th><th>Incid Real</th><th>Incid Desvío</th>
    </tr>`,
    totalRow + rows.map((r, i) => {
      const incidPlan    = r.incidencia * r.pctCompPlan;
      const incidReal    = r.incidencia * r.pctCompReal;
      const realCls      = devClass(r.pctCompReal - r.pctCompPlan);
      const incidRealCls = devClass(incidReal - incidPlan);
      return `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td class="left" style="font-size:11px;color:var(--text-muted)">${_areaOfEdt(r.edt, areaMap)}</td>
        <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td>
        <td>${pct(r.pctCompPlan)}</td>
        <td class="${realCls}">${pct(r.pctCompReal)}</td>
        <td class="${devClass(r.desviacion)}">${signPct(r.desviacion)}</td>
        <td>${pct(r.incidencia,4)}</td>
        <td>${pct(incidPlan,4)}</td>
        <td class="${incidRealCls}">${pct(incidReal,4)}</td>
        <td class="${devClass(r.desvPond)}">${signPct(r.desvPond)}</td>
      </tr>`;
    }).join('')
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

  const sinAvanceDonutCanvas = document.getElementById('sinAvanceDonut');
  if (sinAvanceDonutCanvas) {
    destroyChart('sinAvanceDonut');
    charts['sinAvanceDonut'] = new Chart(sinAvanceDonutCanvas, {
      type:'doughnut',
      data:{ labels, datasets:[{ data: keys.map(k=>byCount[k]),
        backgroundColor: colors.slice(0,keys.length), borderWidth:2, borderColor:'#fff' }] },
      options:{ responsive:true, maintainAspectRatio:false,
        plugins:{ legend:{ position:'right', labels:{ font:{size:11}, padding:8 } } } }
    });
  }

  const hhEntries = Object.entries(byHH).sort((a,b) => b[1]-a[1]);
  const sinAvanceHHBarCanvas = document.getElementById('sinAvanceHHBar');
  if (!sinAvanceHHBarCanvas) return;
  destroyChart('sinAvanceHHBar');
  charts['sinAvanceHHBar'] = new Chart(sinAvanceHHBarCanvas, {
    type:'bar',
    data:{ labels: hhEntries.map(([k]) => areaNames[k]||k),
      datasets:[{ label: t('sin.hhDataset'), data: hhEntries.map(([,v]) => Math.round(v)),
        backgroundColor:'rgba(192,0,0,0.65)', borderRadius:3 }] },
    options:{ indexAxis:'y', responsive:true, maintainAspectRatio:false,
      plugins:{ legend:{display:false} },
      scales:{ x:{ ticks:{ font:{size:11} } } } }
  });
}

function renderSinAvanceTable(rows) {
  _sinAvanceRows = rows;
  const sinAvanceTableEl = document.getElementById('sinAvanceTable');
  if (!sinAvanceTableEl) return;
  const sort   = document.getElementById('sinSort')?.value || 'incidencia';
  const incHdr = `Incid Total${sort === 'incidencia' ? ' ▼' : ''}`;
  const plnHdr = `${t('th.pctPlan')}${sort === 'pctPlan' ? ' ▼' : ''}`;
  // pre-compute per-row values and totals
  const computed = rows.map(r => {
    const incidPlan = r.incidencia * r.pctCompPlan;
    const incidDesv = -incidPlan;                   // real=0 → desvio = 0 - plan
    return { r, incidPlan, incidDesv };
  });
  const totalIncidPlan = computed.reduce((s, c) => s + c.incidPlan, 0);
  const totalIncidDesv = computed.reduce((s, c) => s + c.incidDesv, 0);

  const bodyRows = computed.map(({ r, incidPlan, incidDesv }, i) => `<tr>
        <td>${i+1}</td><td class="left">${r.tarea.trim()}</td><td>${r.edt}</td>
        <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
        <td>${Math.round(r.hh).toLocaleString()}</td>
        <td>${pct(r.pctCompPlan)}</td>
        <td class="dev-neg">0.0%</td>
        <td class="dev-neg">${signPct(-r.pctCompPlan)}</td>
        <td>${pct(r.incidencia, 4)}</td>
        <td>${pct(incidPlan, 4)}</td>
        <td class="dev-neg">0.000%</td>
        <td class="dev-neg">${signPct(incidDesv, 4)}</td>
      </tr>`).join('');

  // total row: empty cells 1-9, value cols 10-12
  const totalRow = `<tr class="dv-total-row">
        <td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td><td></td>
        <td></td>
        <td><strong>${pct(totalIncidPlan, 4)}</strong></td>
        <td class="dev-neg"><strong>0.000%</strong></td>
        <td class="dev-neg"><strong>${signPct(totalIncidDesv, 4)}</strong></td>
      </tr>`;

  sinAvanceTableEl.innerHTML = tableWrap(
    `<tr>
      <th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th>
      <th>${t('th.start')}</th><th>${t('th.end')}</th><th>${t('th.hh')}</th>
      <th>${plnHdr}</th><th>${t('th.pctActual')}</th><th>% Desvío</th>
      <th>${incHdr}</th><th>Incid Plan</th><th>Incid Real</th><th>Incid Desvío</th>
    </tr>`,
    totalRow + bodyRows
  );
}

/** Export Sin Avance table to PDF (landscape A4, with green total row). */
function exportSinAvancePDF() {
  if (!window.jspdf) { showToast('jsPDF not available — check CDN connection', true); return; }
  if (!D || !_sinAvanceRows.length) { showToast(t('toast.error') + 'Sin datos', true); return; }
  const btn = document.getElementById('sinAvancePdfBtn');
  if (btn) btn.disabled = true;
  showToast(t('pdf.generating'), false, 60000);
  try {
    const { jsPDF } = window.jspdf;
    const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const ML    = 10;
    const title = t('sin.title');
    let startY  = _pdfHeader(doc, title, pageW, ML);
    startY      = _pdfMeta(doc, startY, ML, pageW);

    // Pre-compute
    const computed = _sinAvanceRows.map(r => ({
      r,
      incidPlan: r.incidencia * r.pctCompPlan,
      incidDesv: -(r.incidencia * r.pctCompPlan),
    }));
    const totIncidPlan = computed.reduce((s, c) => s + c.incidPlan, 0);
    const totIncidDesv = computed.reduce((s, c) => s + c.incidDesv, 0);

    const head = [['#', t('th.activity'), t('th.edt'),
      t('th.start'), t('th.end'), t('th.hh'),
      t('th.pctPlan'), '% Real', '% Desvío',
      'Incid Total', 'Incid Plan', 'Incid Real', 'Incid Desvío']];

    // Row 0 = total (green), rows 1..n = data
    const totalBodyRow = ['', '', '', '', '', '', '', '', '', '',
      (totIncidPlan * 100).toFixed(4) + '%',
      '0.000%',
      ((totIncidDesv * 100) >= 0 ? '+' : '') + (totIncidDesv * 100).toFixed(4) + '%',
    ];

    const dataRows = computed.map(({ r, incidPlan, incidDesv }, i) => [
      String(i + 1),
      r.tarea.trim(),
      r.edt,
      fmtDate(r.inicio),
      fmtDate(r.fin),
      Math.round(r.hh).toLocaleString('es-CL'),
      (r.pctCompPlan * 100).toFixed(2) + '%',
      '0.0%',
      '-' + (r.pctCompPlan * 100).toFixed(2) + '%',
      (r.incidencia * 100).toFixed(4) + '%',
      (incidPlan * 100).toFixed(4) + '%',
      '0.000%',
      ((incidDesv * 100) >= 0 ? '+' : '') + (incidDesv * 100).toFixed(4) + '%',
    ]);

    const body = [totalBodyRow, ...dataRows];

    doc.autoTable({
      head, body, startY,
      margin: { top: 23, left: ML, right: ML, bottom: 14 },
      styles: { fontSize: 6.5, cellPadding: { top: 1.4, right: 2, bottom: 1.4, left: 2 },
        overflow: 'linebreak', valign: 'middle', lineColor: [210, 218, 230], lineWidth: 0.2 },
      headStyles: { fillColor: [0, 57, 115], textColor: 255, fontStyle: 'bold', fontSize: 7, halign: 'center' },
      alternateRowStyles: { fillColor: [250, 252, 255] },
      columnStyles: {
        0:  { cellWidth:  8, halign: 'center' },
        1:  { cellWidth: 58, halign: 'left'   },
        2:  { cellWidth: 24, halign: 'left'   },
        3:  { cellWidth: 17, halign: 'center' },
        4:  { cellWidth: 17, halign: 'center' },
        5:  { cellWidth: 13, halign: 'right'  },
        6:  { cellWidth: 13, halign: 'right'  },
        7:  { cellWidth: 12, halign: 'right'  },
        8:  { cellWidth: 14, halign: 'right', fontStyle: 'bold' },
        9:  { cellWidth: 16, halign: 'right'  },
        10: { cellWidth: 16, halign: 'right'  },
        11: { cellWidth: 16, halign: 'right'  },
        12: { cellWidth: 18, halign: 'right', fontStyle: 'bold' },
      },
      didParseCell(data) {
        if (data.section !== 'body') return;
        if (data.row.index === 0) {
          // Total row — green background, bold black text
          data.cell.styles.fillColor  = [214, 240, 224];
          data.cell.styles.fontStyle  = 'bold';
          data.cell.styles.textColor  = [0, 0, 0];
          data.cell.styles.lineColor  = [22, 163, 74];
          data.cell.styles.lineWidth  = { top: 0.6, bottom: 0.6, left: 0.1, right: 0.1 };
          return;
        }
        // Data rows
        if (data.column.index === 8)  data.cell.styles.textColor = [200, 30, 30]; // % Desvío
        if (data.column.index === 12) data.cell.styles.textColor = [200, 30, 30]; // Incid Desvío
      },
      didDrawPage(data) {
        _pdfHeader(doc, title, pageW, ML);
        const pg = doc.internal.getCurrentPageInfo().pageNumber;
        doc.setFontSize(7); doc.setTextColor(160, 160, 160);
        doc.text('PowerChina · La Pampina', ML, pageH - 4);
        doc.text(`${t('pdf.page')} ${pg}`, pageW - ML, pageH - 4, { align: 'right' });
        doc.setTextColor(0, 0, 0);
      },
    });

    _pdfPageNumbers(doc, pageW, pageH, ML);
    doc.save('SinAvance_LaPampina.pdf');
    showToast(t('pdf.success'));
  } catch (err) {
    console.error('[exportSinAvancePDF]', err);
    showToast(t('toast.error') + err.message, true);
  } finally {
    if (btn) btn.disabled = false;
  }
}

/** Export Críticas table to PDF (landscape A4, with green total row). */
function exportCriticasPDF() {
  if (!window.jspdf) { showToast('jsPDF not available — check CDN connection', true); return; }
  if (!D || !_criticasRows.length) { showToast(t('toast.error') + 'Sin datos', true); return; }
  const btn = document.getElementById('criticasPdfBtn');
  if (btn) btn.disabled = true;
  showToast(t('pdf.generating'), false, 60000);
  try {
    const { jsPDF } = window.jspdf;
    const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const ML    = 10;
    const title = t('cr.title');
    let startY  = _pdfHeader(doc, title, pageW, ML);
    startY      = _pdfMeta(doc, startY, ML, pageW);
    const areaMap = _buildAreaMap();

    // Pre-compute totals
    const computed = _criticasRows.map(r => ({
      r,
      incidPlan: r.incidencia * r.pctCompPlan,
      incidReal: r.incidencia * r.pctCompReal,
      incidDesv: r.incidencia * (r.desviacion || 0),
    }));
    const totIncidPlan = computed.reduce((s, c) => s + c.incidPlan, 0);
    const totIncidReal = computed.reduce((s, c) => s + c.incidReal, 0);
    const totIncidDesv = computed.reduce((s, c) => s + c.incidDesv, 0);

    const head = [['#', t('th.activity'), t('th.edt'), t('th.area'),
      t('th.start'), t('th.end'), t('th.hh'),
      t('th.pctPlan'), '% Real', '% Desvío',
      'Incid Total', 'Incid Plan', 'Incid Real', 'Incid Desvío']];

    // Row 0 = total (green)
    const s4 = v => ((v * 100) >= 0 ? '+' : '') + (v * 100).toFixed(4) + '%';
    const totalBodyRow = ['', '', '', '', '', '', '', '', '', '', '',
      (totIncidPlan * 100).toFixed(4) + '%',
      (totIncidReal * 100).toFixed(4) + '%',
      s4(totIncidDesv),
    ];

    const dataRows = computed.map(({ r, incidPlan, incidReal, incidDesv }, i) => [
      String(i + 1),
      r.tarea.trim(),
      r.edt,
      _areaOfEdt(r.edt, areaMap),
      fmtDate(r.inicio),
      fmtDate(r.fin),
      Math.round(r.hh).toLocaleString('es-CL'),
      (r.pctCompPlan * 100).toFixed(2) + '%',
      (r.pctCompReal * 100).toFixed(2) + '%',
      ((r.desviacion || 0) >= 0 ? '+' : '') + ((r.desviacion || 0) * 100).toFixed(2) + '%',
      (r.incidencia * 100).toFixed(4) + '%',
      (incidPlan * 100).toFixed(4) + '%',
      (incidReal * 100).toFixed(4) + '%',
      s4(incidDesv),
    ]);

    const body = [totalBodyRow, ...dataRows];

    doc.autoTable({
      head, body, startY,
      margin: { top: 23, left: ML, right: ML, bottom: 14 },
      styles: { fontSize: 6.5, cellPadding: { top: 1.4, right: 2, bottom: 1.4, left: 2 },
        overflow: 'linebreak', valign: 'middle', lineColor: [210, 218, 230], lineWidth: 0.2 },
      headStyles: { fillColor: [0, 57, 115], textColor: 255, fontStyle: 'bold', fontSize: 7, halign: 'center' },
      alternateRowStyles: { fillColor: [250, 252, 255] },
      columnStyles: {
        0:  { cellWidth:  7, halign: 'center' },
        1:  { cellWidth: 50, halign: 'left'   },
        2:  { cellWidth: 22, halign: 'left'   },
        3:  { cellWidth: 28, halign: 'left'   },
        4:  { cellWidth: 16, halign: 'center' },
        5:  { cellWidth: 16, halign: 'center' },
        6:  { cellWidth: 12, halign: 'right'  },
        7:  { cellWidth: 13, halign: 'right'  },
        8:  { cellWidth: 13, halign: 'right'  },
        9:  { cellWidth: 14, halign: 'right', fontStyle: 'bold' },
        10: { cellWidth: 14, halign: 'right'  },
        11: { cellWidth: 14, halign: 'right'  },
        12: { cellWidth: 14, halign: 'right'  },
        13: { cellWidth: 17, halign: 'right', fontStyle: 'bold' },
      },
      didParseCell(data) {
        if (data.section !== 'body') return;
        if (data.row.index === 0) {
          // Total row — green background, bold black text
          data.cell.styles.fillColor = [214, 240, 224];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.textColor = [0, 0, 0];
          data.cell.styles.lineColor = [22, 163, 74];
          data.cell.styles.lineWidth = { top: 0.6, bottom: 0.6, left: 0.1, right: 0.1 };
          return;
        }
        const c = computed[data.row.index - 1];
        if (!c) return;
        if (data.column.index === 9) {  // % Desvío
          const dev = c.r.desviacion || 0;
          if (dev < -0.00001) data.cell.styles.textColor = [200, 30, 30];
          else if (dev > 0.00001) data.cell.styles.textColor = [22, 130, 60];
        }
        if (data.column.index === 13) {  // Incid Desvío
          const v = c.incidDesv;
          if (v < -0.00001) data.cell.styles.textColor = [200, 30, 30];
          else if (v > 0.00001) data.cell.styles.textColor = [22, 130, 60];
        }
      },
      didDrawPage(data) {
        _pdfHeader(doc, title, pageW, ML);
        const pg = doc.internal.getCurrentPageInfo().pageNumber;
        doc.setFontSize(7); doc.setTextColor(160, 160, 160);
        doc.text('PowerChina · La Pampina', ML, pageH - 4);
        doc.text(`${t('pdf.page')} ${pg}`, pageW - ML, pageH - 4, { align: 'right' });
        doc.setTextColor(0, 0, 0);
      },
    });

    _pdfPageNumbers(doc, pageW, pageH, ML);
    doc.save('Criticas_LaPampina.pdf');
    showToast(t('pdf.success'));
  } catch (err) {
    console.error('[exportCriticasPDF]', err);
    showToast(t('toast.error') + err.message, true);
  } finally {
    if (btn) btn.disabled = false;
  }
}

// ── Ranking Bar + Table ───────────────────────────────────────────────────────
function _renderRankingBar(rows, canvasId, isNeg) {
  const canvas = document.getElementById(canvasId);
  if (!canvas) return;

  const vals   = rows.map(r => +(Math.abs(r.desvPond) * 100).toFixed(3));
  const labels = rows.map(r => { const n = r.tarea.trim(); return n.length > 38 ? n.slice(0, 37) + '…' : n; });

  const chartHeight = Math.max(160, rows.length * 34);
  canvas.parentElement.style.height = chartHeight + 'px';

  destroyChart(canvasId);
  charts[canvasId] = new Chart(canvas, {
    type: 'bar',
    data: {
      labels,
      datasets: [{ data: vals, borderRadius: 3,
        backgroundColor: isNeg
          ? vals.map(v => v > 0.5 ? 'rgba(192,0,0,0.75)' : v > 0.2 ? 'rgba(230,138,0,0.75)' : 'rgba(0,84,166,0.65)')
          : vals.map(v => v > 0.5 ? 'rgba(0,132,74,0.8)'  : v > 0.2 ? 'rgba(0,132,74,0.55)' : 'rgba(0,132,74,0.35)'),
      }]
    },
    options: {
      indexAxis: 'y', responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: ctx => {
          const r = rows[ctx.dataIndex];
          return [`Impacto: ${isNeg ? '-' : '+'}${ctx.parsed.x.toFixed(3)}%`, `EDT: ${r.edt}`];
        }}}
      },
      scales: {
        x: { ticks: { callback: v => v + '%' } },
        y: { ticks: { font: { size: 11 } } }
      }
    }
  });
}

function _renderRankingTable(rows, tableId, isNeg) {
  const el = document.getElementById(tableId);
  if (!el) return;

  // ── totals ────────────────────────────────────────────────────────────────
  const totalHH       = rows.reduce((s, r) => s + (r.hh || 0), 0);
  const totalIncid    = rows.reduce((s, r) => s + r.incidencia, 0);
  const totalPctPlan  = rows.reduce((s, r) => s + r.pctPlan, 0);
  const totalPctReal  = rows.reduce((s, r) => s + r.pctReal, 0);
  const totalDesvPond = rows.reduce((s, r) => s + r.desvPond, 0);
  const avgPlan       = totalIncid > 0 ? totalPctPlan / totalIncid : 0;
  const avgReal       = totalIncid > 0 ? totalPctReal / totalIncid : 0;

  const rankTotalRow = `<tr class="dv-total-row">
    <td></td><td></td><td></td>
    <td><strong>${Math.round(totalHH).toLocaleString()}</strong></td>
    <td><strong>${pct(totalIncid, 4)}</strong></td>
    <td><strong>${pct(totalPctPlan, 4)}</strong></td>
    <td class="${devClass(avgReal - avgPlan)}"><strong>${pct(totalPctReal, 4)}</strong></td>
    <td class="${devClass(totalDesvPond)}"><strong>${signPct(totalDesvPond, 4)}</strong></td>
    <td><strong>${pct(avgPlan)}</strong></td>
    <td class="${devClass(avgReal - avgPlan)}"><strong>${pct(avgReal)}</strong></td>
    <td></td>
  </tr>`;

  el.innerHTML = tableWrap(
    `<tr><th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th><th>${t('th.hh')}</th>
     <th>${t('th.incidence')}</th><th>${t('th.pctPlanPond')}</th><th>${t('th.pctRealPond')}</th><th>${t('th.impactPond')}</th>
     <th>${t('th.pctPlan')}</th><th>${t('th.pctActual')}</th><th>${t('th.clasif')}</th></tr>`,
    rankTotalRow + rows.map((r, i) => {
      const imp = Math.abs(r.desvPond);
      const cls = imp>0.005?'badge badge-crit':imp>0.002?'badge badge-late':imp>0.0005?'badge badge-warn':'badge badge-ok';
      const lbl = imp>0.005?t('rank.crit'):imp>0.002?t('rank.high'):imp>0.0005?t('rank.mid'):t('rank.low');
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

/** Public wrappers — called by render() and updateRank() */
function renderRankingBar(rows)   { _renderRankingBar(rows, 'rankingBarNegChart', true); }
function renderRankingTable(rows) { _renderRankingTable(rows, 'rankingNegTable', true); }
function renderRankingBarPos(rows){ _renderRankingBar(rows, 'rankingBarPosChart', false); }
function renderRankingTablePos(rows){ _renderRankingTable(rows, 'rankingPosTable', false); }

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
        pbIdx:      zone.pbIdx,  // index of PB segment in EDT parts
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
      pbIdx:      g.pbIdx,
      leaves:     g.leaves.slice().sort((a, b) => edtCmp(a.edt, b.edt)),
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
        <span class="cs-pill">Incid. Plan <strong>${pct(pC)}</strong></span>
        <span class="cs-pill">Incid. Real <strong>${pct(rC)}</strong></span>
        <span class="cs-pill ${gap < 0 ? 'cs-neg' : 'cs-pos'}">Incid. Desv <strong>${signPct(gap)}</strong></span>`;
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

  const thead = `<tr>
    <th class="left" style="min-width:240px">Actividad</th>
    <th class="left" style="min-width:150px">Zona / Disciplina</th>
    <th title="Total de Power Blocks que tienen esta actividad">PB total</th>
    <th title="PB donde el avance planificado al corte es &gt; 0">PB planif.</th>
    <th title="PB donde el avance real es &gt; 0">PB c/avance</th>
    <th title="PB donde el avance real es menor al planificado">PB c/desvío</th>
    <th title="Suma de incidencias de todos los PB del grupo">Incid. total</th>
    <th title="Σ incidencia × plan — peso del avance planificado">Incid. plan.</th>
    <th title="Σ incidencia × real — peso del avance real">Incid. real</th>
    <th title="Incid. real − Incid. plan. (desvío ponderado)">Incid. desvío</th>
    <th title="Promedio ponderado del avance planificado">% Plan</th>
    <th title="Promedio ponderado del avance real">% Real</th>
    <th title="Barras de progreso plan vs real">Progreso</th>
    <th title="% Real − % Plan (desvío porcentual consolidado)">% Desv.</th>
  </tr>`;

  let tbody = '';
  rows.forEach((g, idx) => {
    const gapCls = devClass(g.gap);
    const slug   = g.zonePrefix.replace(/\./g, '-');
    const pctAv  = g.pbTotal > 0 ? (g.pbAv / g.pbTotal * 100).toFixed(0) : 0;
    const miniBar = `<div class="pb-mini" title="${g.pbAv}/${g.pbTotal} PB con avance real">
      <div class="pb-mini-fill" style="width:${pctAv}%"></div></div>`;

    // ── Group (summary) row ───────────────────────────────────────────────
    tbody += `<tr class="cons-group-row">
      <td class="left">
        <button class="cons-expand-btn" data-cidx="${idx}" title="Ver PB individuales">▶</button>
        <span class="cons-tarea">${g.tarea}</span>
      </td>
      <td class="left"><span class="zone-tag zone-${slug}">${g.zoneLabel}</span></td>
      <td class="cons-pb-num"><strong>${g.pbTotal}</strong></td>
      <td>${g.pbPlan}</td>
      <td>${g.pbAv}${miniBar}</td>
      <td class="${g.pbDev > 0 ? 'dev-neg' : 'dev-neutral'}">${g.pbDev > 0 ? g.pbDev : '—'}</td>
      <td>${pct(g.incTotal, 3)}</td>
      <td>${pct(g.pesoPlan, 3)}</td>
      <td>${pct(g.pesoReal, 3)}</td>
      <td class="${devClass(g.pesoReal - g.pesoPlan)}"><strong>${signPct(g.pesoReal - g.pesoPlan, 3)}</strong></td>
      <td>${pct(g.planConsol)}</td>
      <td>${pct(g.realConsol)}</td>
      <td><div class="cons-leaf-bars">
        <div class="cons-leaf-bar" style="width:${(g.planConsol*100).toFixed(1)}%;background:var(--primary)" title="Plan ${(g.planConsol*100).toFixed(1)}%"></div>
        <div class="cons-leaf-bar" style="width:${(g.realConsol*100).toFixed(1)}%;background:${g.gap < -0.001 ? 'var(--danger)' : 'var(--success)'}" title="Real ${(g.realConsol*100).toFixed(1)}%"></div>
      </div></td>
      <td class="${gapCls}"><strong>${signPct(g.gap)}</strong></td>
    </tr>`;

    // ── Detail rows (one per PB leaf, hidden by default) ─────────────────
    (g.leaves || []).forEach(leaf => {
      const pbNum    = g.pbIdx != null ? (leaf.edt.split('.')[g.pbIdx] || '?') : '?';
      const leafGap  = leaf.pctCompReal - leaf.pctCompPlan;
      const leafCls  = devClass(leafGap);
      const planW    = (leaf.pctCompPlan  * 100).toFixed(1);
      const realW    = (leaf.pctCompReal  * 100).toFixed(1);
      const realClr  = leafGap < -0.001 ? 'var(--danger)' : 'var(--success)';
      const bars     = `<div class="cons-leaf-bars">
        <div class="cons-leaf-bar" style="width:${planW}%;background:var(--primary)" title="Plan ${planW}%"></div>
        <div class="cons-leaf-bar" style="width:${realW}%;background:${realClr}" title="Real ${realW}%"></div>
      </div>`;
      tbody += `<tr class="cons-detail-row" data-cparent="${idx}">
        <td class="left cons-detail-cell">
          <span class="cons-pb-tag">PB ${pbNum}</span>
          <span class="cons-edt-lbl">${leaf.edt}</span>
        </td>
        <td></td>
        <td>—</td><td>—</td><td>—</td><td>—</td>
        <td>${pct(leaf.incidencia, 3)}</td>
        <td>${pct(leaf.incidencia * leaf.pctCompPlan, 3)}</td>
        <td>${pct(leaf.incidencia * leaf.pctCompReal, 3)}</td>
        <td class="${leafCls}"><strong>${signPct(leaf.incidencia * leafGap, 3)}</strong></td>
        <td>${planW}%</td>
        <td>${realW}%</td>
        <td>${bars}</td>
        <td class="${leafCls}"><strong>${signPct(leafGap)}</strong></td>
      </tr>`;
    });
  });

  // ── totals ────────────────────────────────────────────────────────────────
  const totPbTotal  = rows.reduce((s, g) => s + g.pbTotal, 0);
  const totPbPlan   = rows.reduce((s, g) => s + g.pbPlan,  0);
  const totPbAv     = rows.reduce((s, g) => s + g.pbAv,    0);
  const totPbDev    = rows.reduce((s, g) => s + g.pbDev,   0);
  const totIncTotal = rows.reduce((s, g) => s + g.incTotal, 0);
  const totPesoPlan = rows.reduce((s, g) => s + g.pesoPlan, 0);
  const totPesoReal = rows.reduce((s, g) => s + g.pesoReal, 0);
  const totIncDesv  = totPesoReal - totPesoPlan;
  const totPlanCons = totIncTotal > 0 ? totPesoPlan / totIncTotal : 0;
  const totRealCons = totIncTotal > 0 ? totPesoReal / totIncTotal : 0;
  const totGap      = totRealCons - totPlanCons;
  const totGapCls   = devClass(totGap);
  const totBar = `<div class="cons-leaf-bars">
    <div class="cons-leaf-bar" style="width:${(totPlanCons*100).toFixed(1)}%;background:var(--primary)" title="Plan ${(totPlanCons*100).toFixed(1)}%"></div>
    <div class="cons-leaf-bar" style="width:${(totRealCons*100).toFixed(1)}%;background:${totGap < -0.001 ? 'var(--danger)' : 'var(--success)'}" title="Real ${(totRealCons*100).toFixed(1)}%"></div>
  </div>`;
  const consTotalRow = `<tr class="dv-total-row">
    <td class="left"></td>
    <td></td>
    <td><strong>${totPbTotal}</strong></td>
    <td><strong>${totPbPlan}</strong></td>
    <td><strong>${totPbAv}</strong></td>
    <td class="${totPbDev > 0 ? 'dev-neg' : 'dev-neutral'}"><strong>${totPbDev > 0 ? totPbDev : '—'}</strong></td>
    <td><strong>${pct(totIncTotal, 3)}</strong></td>
    <td><strong>${pct(totPesoPlan, 3)}</strong></td>
    <td class="${devClass(totRealCons - totPlanCons)}"><strong>${pct(totPesoReal, 3)}</strong></td>
    <td class="${devClass(totIncDesv)}"><strong>${signPct(totIncDesv, 3)}</strong></td>
    <td><strong>${pct(totPlanCons)}</strong></td>
    <td class="${devClass(totRealCons - totPlanCons)}"><strong>${pct(totRealCons)}</strong></td>
    <td>${totBar}</td>
    <td class="${totGapCls}"><strong>${signPct(totGap)}</strong></td>
  </tr>`;

  el.innerHTML = `<div class="grid-wrap"><table><thead>${thead}</thead><tbody>${consTotalRow}${tbody}</tbody></table></div>`;

  // ── Event delegation: expand / collapse ───────────────────────────────
  const table = el.querySelector('table');
  if (table) {
    table.addEventListener('click', e => {
      const btn = e.target.closest('.cons-expand-btn');
      if (!btn) return;
      const cidx  = btn.dataset.cidx;
      const isOpen = btn.classList.toggle('open');
      el.querySelectorAll(`.cons-detail-row[data-cparent="${cidx}"]`)
        .forEach(r => r.classList.toggle('cons-detail-visible', isOpen));
    });
  }
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
      leaves:   g.leaves,
      pbIdx:    g.pbIdx,
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

/** Thin two-bar visual (plan=blue, real=green/red) matching Consolidado style */
function _progBars(plan, real, desvio) {
  const pw = (plan * 100).toFixed(1);
  const rw = (real * 100).toFixed(1);
  const rc = desvio < -0.00001 ? 'var(--danger)' : 'var(--success)';
  return `<div class="cons-leaf-bars">` +
    `<div class="cons-leaf-bar" style="width:${pw}%;background:var(--primary)" title="Plan ${pw}%"></div>` +
    `<div class="cons-leaf-bar" style="width:${rw}%;background:${rc}" title="Real ${rw}%"></div>` +
    `</div>`;
}

function buildArbolRow(r, edtsWithChildren, hidden) {
  const lvl     = r.nivel || 1;
  const indent  = Math.max(0, lvl - 1) * 16;
  const hasKids = edtsWithChildren.has(r.edt);
  const isColl  = collapsedNodes.has(r.edt);

  const lvlCls = `tree-lvl${Math.min(lvl, 6)}`;
  const hidCls = hidden ? ' tree-hidden' : '';

  const toggleBtn = hasKids
    ? `<button class="tree-toggle" data-edt="${r.edt}" title="${isColl ? t('wb.toggleExpand') : t('wb.toggleCollapse')}">${isColl ? '▶' : '▼'}</button>`
    : `<span class="tree-no-toggle"></span>`;

  // Columns: EDT | Actividad | H-H | PBs | PB Plan. | PB Av. | PB Dev. | Incid. | INCD.PLAN | INCD.REAL | % Plan | % Real | Desvío | Status

  // ── Consolidated leaf ────────────────────────────────────────────────────────
  if (r.isConsolidated) {
    const devCls    = devClass(r.desviacion);
    const pbDevCls  = r.pbDev > 0 ? 'dev-neg' : 'dev-neutral';
    const hhVal     = r.hh > 0 ? Math.round(r.hh).toLocaleString() : '—';
    const hasLeaves = r.leaves?.length > 0;
    const isExp     = hasLeaves && expandedPbNodes.has(r.edt);
    const expandBtn = hasLeaves
      ? `<button class="tree-expand-pb${isExp ? ' open' : ''}" data-edt="${r.edt}" title="Ver PBs individuales">▶</button>`
      : `<span class="tree-no-toggle"></span>`;
    const mainRow = `<tr class="tree-row tree-leaf tree-consol ${lvlCls}${hidCls}" data-edt="${r.edt}" data-rowtype="consolidada">
      <td class="left">
        <div class="tree-edt-cell">
          <span class="tree-indent" style="width:${indent}px"></span>
          ${expandBtn}
          <span class="tree-consol-dot">●</span>
        </div>
      </td>
      <td class="left"><span class="cons-tarea-tree">${r.tarea}</span></td>
      <td>${hhVal}</td>
      <td>${fmtDate(r.inicio)}</td>
      <td>${fmtDate(r.fin)}</td>
      <td class="cons-pb-num"><strong>${r.pbTotal}</strong></td>
      <td>${r.pbPlan}</td>
      <td>${r.pbAv}</td>
      <td class="${pbDevCls}">${r.pbDev != null ? r.pbDev : '—'}</td>
      <td>${pct(r.incidencia,3)}</td>
      <td class="incd-plan">${pct(r.pesoPlan,3)}</td>
      <td class="incd-real">${pct(r.pesoReal,3)}</td>
      <td class="${devClass(r.pesoReal - r.pesoPlan)}">${signPct(r.pesoReal - r.pesoPlan, 3)}</td>
      <td>${pct(r.pctCompPlan)}</td>
      <td>${pct(r.pctCompReal)}</td>
      <td class="${devCls}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
      <td>${_progBars(r.pctCompPlan, r.pctCompReal, r.desviacion)}</td>
    </tr>`;
    if (!hasLeaves) return mainRow;
    const detailRows = r.leaves.map(leaf => {
      const leafGap = leaf.pctCompReal - leaf.pctCompPlan;
      const leafCls = devClass(leafGap);
      const pbNum   = r.pbIdx != null ? (leaf.edt.split('.')[r.pbIdx] || '?') : '?';
      const planW   = (leaf.pctCompPlan * 100).toFixed(1);
      const realW   = (leaf.pctCompReal * 100).toFixed(1);
      const realClr = leafGap < -0.00001 ? 'var(--danger)' : 'var(--success)';
      const bars    = `<div class="cons-leaf-bars">
        <div class="cons-leaf-bar" style="width:${planW}%;background:var(--primary)" title="Plan ${planW}%"></div>
        <div class="cons-leaf-bar" style="width:${realW}%;background:${realClr}" title="Real ${realW}%"></div>
      </div>`;
      const visClass = isExp ? ' arbol-pb-detail-visible' : '';
      return `<tr class="arbol-pb-detail${visClass}${hidCls}" data-pb-parent="${r.edt}">
        <td class="left">
          <div class="tree-edt-cell">
            <span class="tree-indent" style="width:${indent + 20}px"></span>
            <span class="tree-no-toggle"></span>
            <span class="cons-pb-tag">PB ${pbNum}</span>
          </div>
        </td>
        <td class="left"><span class="cons-edt-lbl">${leaf.edt}</span></td>
        <td>—</td>
        <td>${fmtDate(leaf.inicio)}</td>
        <td>${fmtDate(leaf.fin)}</td>
        <td>—</td><td>—</td><td>—</td><td>—</td>
        <td>${pct(leaf.incidencia,3)}</td>
        <td>${pct(leaf.incidencia * leaf.pctCompPlan,3)}</td>
        <td>${pct(leaf.incidencia * leaf.pctCompReal,3)}</td>
        <td class="${leafCls}"><strong>${signPct(leaf.incidencia * leafGap, 3)}</strong></td>
        <td>${planW}%</td>
        <td>${realW}%</td>
        <td class="${leafCls}"><strong>${signPct(leafGap)}</strong></td>
        <td>${statusBadge(leaf)}</td>
        <td>${bars}</td>
      </tr>`;
    }).join('');
    return mainRow + detailRows;
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
      <td>${fmtDate(r.inicio)}</td>
      <td>${fmtDate(r.fin)}</td>
      <td class="cons-pb-num"><strong>${r.pbTotal}</strong></td>
      <td>${r.pbPlan}</td>
      <td>${r.pbAv}</td>
      <td class="${pbDevCls}">${r.pbDev != null ? r.pbDev : '—'}</td>
      <td>${pct(r.incidencia,3)}</td>
      <td class="incd-plan">${pct(r.pesoPlan,3)}</td>
      <td class="incd-real">${pct(r.pesoReal,3)}</td>
      <td class="${devClass(r.pesoReal - r.pesoPlan)}">${signPct(r.pesoReal - r.pesoPlan, 3)}</td>
      <td>${pct(r.pctCompPlan)}</td>
      <td>${pct(r.pctCompReal)}</td>
      <td class="${devCls}">${signPct(r.desviacion)}</td>
      <td>${statusBadge(r)}</td>
      <td>${_progBars(r.pctCompPlan, r.pctCompReal, r.desviacion)}</td>
    </tr>`;
  }

  // ── Regular record ───────────────────────────────────────────────────────────
  const isSum   = r.resumen;
  const typeCls = isSum ? 'tree-summary' : 'tree-leaf';
  const devCls      = devClass(r.desviacion);
  const hh          = r.hh > 0 ? Math.round(r.hh).toLocaleString() : '—';
  const incid       = r.incidencia > 0 ? pct(r.incidencia, 3) : '—';
  const devVal      = r.incidencia > 0.0001 ? signPct(r.desviacion) : '—';
  const incidDesv   = r.incidencia > 0.0001 ? r.incidencia * r.desviacion : null;
  const incidDesvCls = devClass(incidDesv);
  const incidDesvVal = incidDesv != null ? signPct(incidDesv, 3) : '—';

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
    <td>${fmtDate(r.inicio)}</td>
    <td>${fmtDate(r.fin)}</td>
    <td class="cons-pb-num">${hasPB ? `<strong>${r.pbTotal}</strong>` : '—'}</td>
    <td>${hasPB ? r.pbPlan : '—'}</td>
    <td>${hasPB ? r.pbAv   : '—'}</td>
    <td class="${pbDevCls2}">${hasPB ? r.pbDev : '—'}</td>
    <td>${incid}</td>
    <td class="incd-plan">${incdPlan}</td>
    <td class="incd-real">${incdReal}</td>
    <td class="${incidDesvCls}">${incidDesvVal}</td>
    <td>${pct(r.pctCompPlan)}</td>
    <td>${pct(r.pctCompReal)}</td>
    <td class="${devCls}">${devVal}</td>
    <td>${statusBadge(r)}</td>
    <td>${r.incidencia > 0.0001 ? _progBars(r.pctCompPlan, r.pctCompReal, r.desviacion) : '—'}</td>
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
    // Search mode: show flat matches, no hiding; no total row
    const matches = allRecs.filter(r =>
      r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)
    );
    tbody.innerHTML = matches.map(r => buildArbolRow(r, edtsWithChildren, false)).join('');
    return;
  }

  // ── WBS total row — uses D.meta (same source as KPI bar) ───────────────────
  const m        = D.meta;
  const devCls   = devClass(m.desvio);
  const realCls  = devClass(m.pctReal - m.pctPlan);
  const wbsTotalRow = `<tr class="dv-total-row">
    <td></td>
    <td></td>
    <td><strong>${Math.round(m.totalHH).toLocaleString()}</strong></td>
    <td>—</td><td>—</td>
    <td>—</td><td>—</td><td>—</td><td>—</td>
    <td><strong>100.000%</strong></td>
    <td class="incd-plan"><strong>${pct(m.pctPlan, 3)}</strong></td>
    <td class="incd-real"><strong>${pct(m.pctReal, 3)}</strong></td>
    <td class="${devCls}"><strong>${signPct(m.desvio, 3)}</strong></td>
    <td><strong>${pct(m.pctPlan)}</strong></td>
    <td class="${realCls}"><strong>${pct(m.pctReal)}</strong></td>
    <td class="${devCls}"><strong>${signPct(m.desvio)}</strong></td>
    <td>—</td>
    <td>${_progBars(m.pctPlan, m.pctReal, m.desvio)}</td>
  </tr>`;

  tbody.innerHTML = wbsTotalRow + allRecs.map(r => {
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
  sel.innerHTML = `<option value="">${t('cr.allAreas')}</option>`
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
    if (!D) return;
    // ── WBS hierarchy toggle ───────────────────────────────────────────────
    const treeBtn = e.target.closest('.tree-toggle');
    if (treeBtn) {
      const edt = treeBtn.dataset.edt;
      if (collapsedNodes.has(edt)) collapsedNodes.delete(edt);
      else                          collapsedNodes.add(edt);
      renderArbol(document.getElementById('arbolSearch')?.value || '');
      return;
    }
    // ── PB individual expand/collapse ──────────────────────────────────────
    const pbBtn = e.target.closest('.tree-expand-pb');
    if (pbBtn) {
      const edt    = pbBtn.dataset.edt;
      const isOpen = pbBtn.classList.toggle('open');
      if (isOpen) expandedPbNodes.add(edt); else expandedPbNodes.delete(edt);
      document.getElementById('arbolBody')
        ?.querySelectorAll(`.arbol-pb-detail[data-pb-parent="${edt}"]`)
        .forEach(tr => tr.classList.toggle('arbol-pb-detail-visible', isOpen));
    }
  });
}

// ── Future ────────────────────────────────────────────────────────────────────
function renderFuture() {
  const futureTableEl = document.getElementById('futureTable');
  if (!futureTableEl) return;
  futureTableEl.innerHTML = tableWrap(
    `<tr><th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th>
     <th>${t('th.hh')}</th><th>${t('th.incidence')}</th><th>${t('th.pctActualCur')}</th><th>${t('th.pctRecoverable')}</th><th>${t('th.end')}</th></tr>`,
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


// ── Populate area dropdowns ───────────────────────────────────────────────────
function populateAreaDropdowns() {
  const areaNames = {};
  D.areas.filter(a => a.nivel===3).forEach(a => { areaNames[a.edt] = a.tarea.trim().slice(0,28); });
  const opts = Object.entries(areaNames).sort()
    .map(([k,v]) => `<option value="${k}">${k} — ${v}</option>`).join('');
  ['desvAreaBox','critAreaBox','sinAreaBox','rankAreaBox'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<option value="">${t('cr.allAreas')}</option>` + opts;
  });
}

// ── Tab filters ───────────────────────────────────────────────────────────────
function setupTabFilters() {

  function updateDesv() {
    if (!D) return;
    const q = document.getElementById('desvSearch').value.toLowerCase();
    const area = document.getElementById('desvAreaBox').value;
    const rows = D.ranking.filter(r =>
      r.desvPond < 0 &&
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));
    renderDesviosBar(rows); renderDesviosTable(rows);
  }
  on('desvSearch',  'input',  updateDesv);
  on('desvAreaBox', 'change', updateDesv);

  function updateCrit() {
    if (!D) return;
    const q    = document.getElementById('critSearch').value.toLowerCase();
    const area = document.getElementById('critAreaBox').value;
    const sort = (document.getElementById('critSort')?.value) || 'desv_asc';

    let rows = D.critical.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));

    // Apply sort
    rows = [...rows].sort((a, b) => {
      if (sort === 'desv_asc')  return a.desviacion - b.desviacion;   // most negative first
      if (sort === 'desv_desc') return b.desviacion - a.desviacion;   // least negative first
      if (sort === 'imp_asc')   return a.desvPond   - b.desvPond;     // most negative impact first
      if (sort === 'imp_desc')  return b.desvPond   - a.desvPond;     // least negative impact first
      return 0;
    });

    renderCriticasBar(rows); renderCriticasTable(rows);
  }
  on('critSearch',  'input',  updateCrit);
  on('critAreaBox', 'change', updateCrit);
  on('critSort',    'change', updateCrit);

  function updateSin() {
    if (!D) return;
    const q    = document.getElementById('sinSearch').value.toLowerCase();
    const area = document.getElementById('sinAreaBox').value;
    const sort = document.getElementById('sinSort')?.value || 'incidencia';

    let rows = D.sinAvance.filter(r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area)));

    if (sort === 'pctPlan') {
      rows = [...rows].sort((a, b) => b.pctCompPlan - a.pctCompPlan);
    } else {
      rows = [...rows].sort((a, b) => b.incidencia - a.incidencia);
    }

    renderSinAvanceCharts(rows); renderSinAvanceTable(rows);
  }
  on('sinSearch',  'input',  updateSin);
  on('sinAreaBox', 'change', updateSin);
  on('sinSort',    'change', updateSin);

  function updateRank() {
    if (!D) return;
    const q         = document.getElementById('rankSearch').value.toLowerCase();
    const area      = document.getElementById('rankAreaBox').value;
    const negImpact = (document.getElementById('rankNegImpact')?.value) || '';
    const posImpact = (document.getElementById('rankPosImpact')?.value) || '';

    const impactMatch = (r, band) => {
      if (!band) return true;
      const pp = Math.abs(r.desvPond) * 100;
      if (band === 'ALTO')  return pp > 0.2;
      if (band === 'MEDIO') return pp > 0.05 && pp <= 0.2;
      if (band === 'BAJO')  return pp <= 0.05;
      return true;
    };

    const textArea = r =>
      (!q    || r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)) &&
      (!area || r.edt.startsWith(area));

    const negRows = D.ranking.filter(r => r.desvPond < 0 && textArea(r) && impactMatch(r, negImpact));
    const posRows = D.ranking.filter(r => r.desvPond > 0 && textArea(r) && impactMatch(r, posImpact))
                             .sort((a, b) => b.desvPond - a.desvPond);

    renderRankingBar(negRows.slice(0, 20));    renderRankingTable(negRows);
    renderRankingBarPos(posRows.slice(0, 15)); renderRankingTablePos(posRows);
  }
  on('rankSearch',    'input',  updateRank);
  on('rankAreaBox',   'change', updateRank);
  on('rankNegImpact', 'change', updateRank);
  on('rankPosImpact', 'change', updateRank);

  // Plazos accordion filters
  on('plSearch',       'input',  () => { if (D) renderPlazos(); });
  on('plStatusFilter', 'change', () => { if (D) renderPlazos(); });
  on('plUpWeeks',      'input',  () => { if (D) renderPlazos(); });
  on('plUpWeeks',      'change', () => { if (D) renderPlazos(); });
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
  sel.innerHTML = `<option value="">${t('sim.selectAct')}</option>`
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
    scopeSel.innerHTML = `<option value="">${t('sim.allProject')}</option>`
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
    wrap.innerHTML = `<p class="subtitle" style="padding:12px 0">${t('sim.noActs')}</p>`;
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
        ${t('sim.remove')}
      </button></td>
    </tr>`;
  }).join('');

  wrap.innerHTML = tableWrap(
    `<tr>
       <th class="left">${t('sim.th.activity')}</th>
       <th>${t('sim.th.type')}</th>
       <th>${t('sim.th.pbTotal')}</th>
       <th>${t('sim.th.pbExec')}</th>
       <th>${t('sim.th.pbPlan')}</th>
       <th>${t('sim.th.pbSim')}</th>
       <th>${t('sim.th.pctReal')}</th>
       <th>${t('sim.th.pctSimReal')}</th>
       <th>${t('sim.th.recovery')}</th>
       <th>${t('sim.th.impact')}</th>
       <th>${t('sim.th.action')}</th>
     </tr>`,
    rowsHtml +
    `<tr class="simtab-total-row">
       <td colspan="8" class="right" style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.05em">
         ${t('sim.totalRow')}</td>
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
  if (!document.getElementById('simtabChart')) return;
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
        { label: t('sim.planDs'),
          data: planData,
          borderColor: '#2563eb', backgroundColor: 'transparent',
          pointRadius: 0, fill: false, tension: 0.3, borderWidth: 2 },
        { label: t('sim.actualDs'),
          data: realData,
          borderColor: '#16a34a', backgroundColor: 'transparent',
          pointRadius: ctx => ctx.dataIndex === lastRealI ? 5 : 0,
          fill: false, tension: 0.3, borderWidth: 2.5, spanGaps: false },
        { label: t('sim.simDs'),
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
  document.querySelectorAll('.nav-item').forEach(btn => {
    btn.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-pane').forEach(p => p.classList.remove('active'));
      btn.classList.add('active');
      document.getElementById('tab-' + btn.dataset.tab)?.classList.add('active');
    });
  });

  // Sidebar toggle (expand / collapse)
  const sidebar = document.getElementById('sidebar');
  document.getElementById('sidebarToggle')?.addEventListener('click', () => {
    sidebar?.classList.toggle('collapsed');
  });
}

// ── Plazos — helpers ─────────────────────────────────────────────────────────
function _dateDiffDays(laterIso, earlierIso) {
  if (!laterIso || !earlierIso) return 0;
  return Math.round((new Date(laterIso) - new Date(earlierIso)) / 86400000);
}
function _isoAddDays(iso, days) {
  if (!iso) return null;
  const d = new Date(iso);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().split('T')[0];
}

// ── Plazos — Accordion por área ───────────────────────────────────────────────
function renderPlazos() {
  if (!D) return;
  const cutDate    = D.meta.dataDate;
  const weeks      = Math.max(1, parseInt(document.getElementById('plUpWeeks')?.value || '4'));
  const futureCut  = _isoAddDays(cutDate, weeks * 7);
  const q          = (document.getElementById('plSearch')?.value || '').toLowerCase();
  const statusF    = document.getElementById('plStatusFilter')?.value || 'all';

  const areas = D.areas
    .filter(a => a.nivel === 3 && a.incidencia > 0)
    .filter(a => !q || a.tarea.toLowerCase().includes(q) || a.edt.toLowerCase().includes(q))
    .sort((a, b) => a.edt.localeCompare(b.edt));

  const container = document.getElementById('plazosAccordion');
  if (!container) return;

  const cards = areas.map(area => {
    const prefix = area.edt + '.';
    const leaves = D.allLeaves.filter(r => r.edt.startsWith(prefix) || r.edt === area.edt);

    const notStarted = leaves.filter(r =>
      r.inicio && r.inicio <= cutDate && r.pctCompReal === 0 && r.incidencia > 0
    ).sort((a, b) => a.inicio.localeCompare(b.inicio));

    const startedLate = leaves.filter(r =>
      r.pctCompReal > 0 && r.pctCompReal < 0.995 &&
      r.inicio && r.inicio <= cutDate &&
      r.pctCompPlan > r.pctCompReal + 0.005 &&
      r.incidencia > 0
    ).sort((a, b) => a.inicio.localeCompare(b.inicio));

    const upcoming = leaves.filter(r =>
      r.inicio && r.inicio > cutDate && r.inicio <= futureCut && r.incidencia > 0
    ).sort((a, b) => a.inicio.localeCompare(b.inicio));

    const total = notStarted.length + startedLate.length + upcoming.length;

    // apply status filter
    const show = statusF === 'all'
      || (statusF === 'notStarted' && notStarted.length  > 0)
      || (statusF === 'behind'     && startedLate.length > 0)
      || (statusF === 'upcoming'   && upcoming.length    > 0)
      || (statusF === 'ok'         && total === 0);
    if (!show) return '';

    return `<div class="pl-card">
      <div class="pl-card-hdr" onclick="togglePlCard(this)">
        <i class="bi bi-chevron-right pl-chevron"></i>
        <span class="pl-card-edt">${area.edt}</span>
        <span class="pl-card-name">${area.tarea.trim()}</span>
        <div class="pl-card-badges">
          ${notStarted.length  ? `<span class="pl-badge pl-badge-late">${notStarted.length} No iniciadas</span>`  : ''}
          ${startedLate.length ? `<span class="pl-badge pl-badge-behind">${startedLate.length} c/ atraso</span>`  : ''}
          ${upcoming.length    ? `<span class="pl-badge pl-badge-up">${upcoming.length} Próximas</span>`           : ''}
          ${total === 0        ? `<span class="pl-badge pl-badge-ok">✓ Sin alertas</span>`                         : ''}
        </div>
      </div>
      <div class="pl-card-body" style="display:none">
        ${_buildPlBody(notStarted, startedLate, upcoming, cutDate, statusF)}
      </div>
    </div>`;
  }).join('');

  container.innerHTML = cards || `<p class="plazos-empty">✓ ${t('pl.noAlerts')}</p>`;
}

function _buildPlBody(notStarted, startedLate, upcoming, cutDate, statusF = 'all') {
  const parts = [];
  const show = s => statusF === 'all' || statusF === s;

  if (show('notStarted') && notStarted.length) {
    parts.push(`<div class="pl-sub pl-sub-late">
      <div class="pl-sub-hdr"><i class="bi bi-clock-history"></i> ${t('pl.lateTitle')} (${notStarted.length})</div>
      ${_plTable(notStarted, 'notStarted', cutDate)}
    </div>`);
  }
  if (show('behind') && startedLate.length) {
    parts.push(`<div class="pl-sub pl-sub-behind">
      <div class="pl-sub-hdr"><i class="bi bi-exclamation-triangle"></i> ${t('pl.behindTitle')} (${startedLate.length})</div>
      ${_plTable(startedLate, 'behind', cutDate)}
    </div>`);
  }
  if (show('upcoming') && upcoming.length) {
    parts.push(`<div class="pl-sub pl-sub-up">
      <div class="pl-sub-hdr"><i class="bi bi-calendar-check"></i> ${t('pl.upTitle')} (${upcoming.length})</div>
      ${_plTable(upcoming, 'upcoming', cutDate)}
    </div>`);
  }
  return parts.join('') || `<p class="plazos-empty">✓ ${t('pl.noAlerts')}</p>`;
}

function _plTable(rows, type, cutDate) {
  const extraH = type === 'notStarted' ? `<th>${t('pl.daysLate')}</th>`
               : type === 'upcoming'   ? `<th>${t('pl.daysToStart')}</th>`
               : '';

  const head = `<tr>
    <th>${t('th.num')}</th><th class="left">${t('th.activity')}</th><th>${t('th.edt')}</th>
    <th>${t('th.start')}</th><th>${t('th.end')}</th><th>${t('th.hh')}</th>
    <th>${t('th.pctPlan')}</th><th>${t('th.pctActual')}</th><th>% Desvío</th>
    <th>Incid Total</th><th>Incid Plan</th><th>Incid Real</th><th>Incid Desvío</th>
    ${extraH}
  </tr>`;

  const body = rows.map((r, i) => {
    const pctReal      = type === 'notStarted' ? 0 : r.pctCompReal;
    const desviacion   = pctReal - r.pctCompPlan;
    const incidPlan    = r.incidencia * r.pctCompPlan;
    const incidReal    = r.incidencia * pctReal;
    const incidDesv    = r.incidencia * desviacion;
    const realCls      = devClass(pctReal - r.pctCompPlan);
    const desvCls      = devClass(desviacion);
    const incidRealCls = devClass(incidReal - incidPlan);
    const incidDesvCls = devClass(incidDesv);
    const overdue      = type === 'behind' && r.fin && r.fin <= cutDate;

    const extraTd = type === 'notStarted'
      ? `<td class="plazos-days-late">${_dateDiffDays(cutDate, r.inicio)}d</td>`
      : type === 'upcoming'
      ? `<td class="plazos-days-upcoming">${_dateDiffDays(r.inicio, cutDate)}d</td>`
      : '';

    return `<tr${overdue ? ' class="pl-row-overdue"' : ''}>
      <td>${i+1}</td>
      <td class="left">${r.tarea.trim()}${overdue ? ` <span class="pl-overdue-tag">${t('pl.overdueTag')}</span>` : ''}</td>
      <td>${r.edt}</td>
      <td>${fmtDate(r.inicio)}</td><td>${fmtDate(r.fin)}</td>
      <td>${Math.round(r.hh).toLocaleString()}</td>
      <td>${pct(r.pctCompPlan)}</td>
      <td class="${realCls}">${pct(pctReal)}</td>
      <td class="${desvCls}">${signPct(desviacion)}</td>
      <td>${pct(r.incidencia, 4)}</td>
      <td>${pct(incidPlan, 4)}</td>
      <td class="${incidRealCls}">${pct(incidReal, 4)}</td>
      <td class="${incidDesvCls}">${signPct(incidDesv, 4)}</td>
      ${extraTd}
    </tr>`;
  }).join('');

  return tableWrap(head, body);
}

// ── Plazos — Excel Export ─────────────────────────────────────────────────────
async function exportPlazosXLSX() {
  if (!D) return;
  const btn = document.getElementById('plExportBtn');
  const origHtml = btn ? btn.innerHTML : '';
  if (btn) { btn.disabled = true; btn.textContent = '⏳ Generando...'; }

  try {
    if (!window.ExcelJS) { alert('ExcelJS no disponible. Revise la conexión a internet.'); return; }

    // ── Same filters as renderPlazos ──────────────────────────────────────────
    const cutDate   = D.meta.dataDate;
    const weeks     = Math.max(1, parseInt(document.getElementById('plUpWeeks')?.value || '4'));
    const futureCut = _isoAddDays(cutDate, weeks * 7);
    const q         = (document.getElementById('plSearch')?.value || '').toLowerCase();
    const statusF   = document.getElementById('plStatusFilter')?.value || 'all';

    const areas = D.areas
      .filter(a => a.nivel === 3 && a.incidencia > 0)
      .filter(a => !q || a.tarea.toLowerCase().includes(q) || a.edt.toLowerCase().includes(q))
      .sort((a, b) => a.edt.localeCompare(b.edt));

    // ── Workbook setup ────────────────────────────────────────────────────────
    const wb = new window.ExcelJS.Workbook();
    wb.creator = 'PowerChina — La Pampina';
    wb.created = new Date();
    const ws = wb.addWorksheet('Plazos', { views: [{ state: 'frozen', ySplit: 3 }] });

    const NCOLS = 14;
    ws.columns = [
      { width: 5  }, // #
      { width: 42 }, // Actividad
      { width: 15 }, // EDT
      { width: 12 }, // Inicio
      { width: 12 }, // Fin
      { width: 10 }, // H-H
      { width: 10 }, // %Plan
      { width: 10 }, // %Real
      { width: 11 }, // %Desvío
      { width: 12 }, // Incid Total
      { width: 12 }, // Incid Plan
      { width: 12 }, // Incid Real
      { width: 13 }, // Incid Desvío
      { width: 12 }, // Días
    ];

    // ── Helper: apply fill + font to all cells in a row ───────────────────────
    const fillRow = (row, bg, fg, bold = false, sz = 10) => {
      for (let c = 1; c <= NCOLS; c++) {
        const cell = row.getCell(c);
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: bg } };
        cell.font = { bold, size: sz, color: { argb: fg } };
        cell.alignment = { vertical: 'middle' };
      }
    };

    // ── Helper: merge a full row ──────────────────────────────────────────────
    const mergeRow = (row) => {
      const n = row.number;
      ws.mergeCells(`A${n}:N${n}`);
      row.getCell(1).alignment = { horizontal: 'left', vertical: 'middle' };
    };

    // ── Row 1: Title ──────────────────────────────────────────────────────────
    const titleRow = ws.addRow([`Seguimiento de Plazos — La Pampina`]);
    titleRow.height = 26;
    fillRow(titleRow, 'FF1E40AF', 'FFFFFFFF', true, 13);
    mergeRow(titleRow);

    // ── Row 2: Metadata ───────────────────────────────────────────────────────
    const metaRow = ws.addRow([`Fecha de corte: ${fmtDate(cutDate)}  |  Semana: ${D.meta.dataWeek || '—'}  |  Próximas: ${weeks} semanas`]);
    metaRow.height = 18;
    fillRow(metaRow, 'FFE0E7FF', 'FF3730A3', false, 10);
    mergeRow(metaRow);

    ws.addRow([]); // spacer

    // ── Section configs ───────────────────────────────────────────────────────
    const secCfg = {
      notStarted: { hdrBg: 'FFFEE2E2', hdrFg: 'FFDC2626', rowBg: 'FFFFF5F5', label: 'No iniciadas' },
      behind:     { hdrBg: 'FFFEF3C7', hdrFg: 'FFB45309', rowBg: 'FFFEFCE8', label: 'Con atraso'   },
      upcoming:   { hdrBg: 'FFDBEAFE', hdrFg: 'FF1D4ED8', rowBg: 'FFEFF6FF', label: 'Próximas'     },
    };

    const COL_HDRS = ['#', 'Actividad', 'EDT', 'Inicio', 'Fin', 'H-H',
                      '% Plan', '% Real', '% Desvío',
                      'Incid Total', 'Incid Plan', 'Incid Real', 'Incid Desvío', 'Días'];

    // ── Build rows per area ───────────────────────────────────────────────────
    for (const area of areas) {
      const prefix = area.edt + '.';
      const leaves = D.allLeaves.filter(r => r.edt.startsWith(prefix) || r.edt === area.edt);

      const notStarted = leaves.filter(r =>
        r.inicio && r.inicio <= cutDate && r.pctCompReal === 0 && r.incidencia > 0
      ).sort((a, b) => a.inicio.localeCompare(b.inicio));

      const startedLate = leaves.filter(r =>
        r.pctCompReal > 0 && r.pctCompReal < 0.995 &&
        r.inicio && r.inicio <= cutDate &&
        r.pctCompPlan > r.pctCompReal + 0.005 && r.incidencia > 0
      ).sort((a, b) => a.inicio.localeCompare(b.inicio));

      const upcoming = leaves.filter(r =>
        r.inicio && r.inicio > cutDate && r.inicio <= futureCut && r.incidencia > 0
      ).sort((a, b) => a.inicio.localeCompare(b.inicio));

      const total = notStarted.length + startedLate.length + upcoming.length;
      const show = statusF === 'all'
        || (statusF === 'notStarted' && notStarted.length > 0)
        || (statusF === 'behind'     && startedLate.length > 0)
        || (statusF === 'upcoming'   && upcoming.length > 0)
        || (statusF === 'ok'         && total === 0);
      if (!show) continue;

      // Area header
      const areaRow = ws.addRow([`${area.edt}   ${area.tarea.trim()}`]);
      areaRow.height = 22;
      fillRow(areaRow, 'FF1E293B', 'FFFFFFFF', true, 11);
      mergeRow(areaRow);

      // Sections
      const sections = [
        { key: 'notStarted', rows: notStarted, extra: 'late'     },
        { key: 'behind',     rows: startedLate, extra: ''        },
        { key: 'upcoming',   rows: upcoming,    extra: 'upcoming' },
      ];

      for (const sec of sections) {
        if (!(statusF === 'all' || statusF === sec.key)) continue;
        if (!sec.rows.length) continue;

        const cfg = secCfg[sec.key];

        // Sub-section header
        const subHdr = ws.addRow([`  ${cfg.label}  (${sec.rows.length})`]);
        subHdr.height = 18;
        fillRow(subHdr, cfg.hdrBg, cfg.hdrFg, true, 10);
        mergeRow(subHdr);

        // Column headers
        const colHdrRow = ws.addRow(COL_HDRS);
        colHdrRow.height = 16;
        colHdrRow.eachCell((cell, ci) => {
          cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
          cell.font = { bold: true, size: 9, color: { argb: 'FF475569' } };
          cell.alignment = { horizontal: ci === 2 ? 'left' : 'center', vertical: 'middle' };
          cell.border = { bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } } };
        });

        // Data rows
        sec.rows.forEach((r, i) => {
          const pctR      = sec.key === 'notStarted' ? 0 : r.pctCompReal;
          const desv      = pctR - r.pctCompPlan;
          const incidPlan = r.incidencia * r.pctCompPlan;
          const incidReal = r.incidencia * pctR;
          const incidDesv = r.incidencia * desv;
          const overdue   = sec.key === 'behind' && r.fin && r.fin <= cutDate;

          const extraVal = sec.key === 'notStarted' ? `${_dateDiffDays(cutDate, r.inicio)}d`
                         : sec.key === 'upcoming'   ? `${_dateDiffDays(r.inicio, cutDate)}d`
                         : '';

          const dataRow = ws.addRow([
            i + 1,
            r.tarea.trim() + (overdue ? '  ⚠ VENCIDA' : ''),
            r.edt,
            fmtDate(r.inicio),
            fmtDate(r.fin),
            Math.round(r.hh),
            +((r.pctCompPlan * 100).toFixed(2)),
            +((pctR * 100).toFixed(2)),
            +((desv * 100).toFixed(2)),
            +((r.incidencia * 100).toFixed(4)),
            +((incidPlan * 100).toFixed(4)),
            +((incidReal * 100).toFixed(4)),
            +((incidDesv * 100).toFixed(4)),
            extraVal,
          ]);
          dataRow.height = 15;

          const evenBg  = i % 2 === 0 ? 'FFFFFFFF' : 'FFF8FAFC';
          const rowBgAr = overdue ? 'FFFFF5F5' : evenBg;

          dataRow.eachCell((cell, ci) => {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: rowBgAr } };
            cell.font = { size: 10, color: { argb: 'FF1E293B' } };
            cell.alignment = { vertical: 'middle', horizontal: ci === 2 ? 'left' : 'center' };
            cell.border = { bottom: { style: 'hair', color: { argb: 'FFCBD5E1' } } };
          });

          // Number format for percentage columns
          [7, 8, 9, 10, 11, 12, 13].forEach(ci => {
            dataRow.getCell(ci).numFmt = '0.00"%"';
          });

          // Color deviation cells
          const dCell = dataRow.getCell(9);
          if      (desv < -0.0001) dCell.font = { size: 10, bold: true, color: { argb: 'FFDC2626' } };
          else if (desv >  0.0001) dCell.font = { size: 10, bold: true, color: { argb: 'FF16A34A' } };

          const idCell = dataRow.getCell(13);
          if      (incidDesv < -0.000001) idCell.font = { size: 10, color: { argb: 'FFDC2626' } };
          else if (incidDesv >  0.000001) idCell.font = { size: 10, color: { argb: 'FF16A34A' } };

          // Color days cell
          const daysCell = dataRow.getCell(14);
          if (sec.key === 'notStarted') daysCell.font = { size: 10, bold: true, color: { argb: 'FFDC2626' } };
          else if (sec.key === 'upcoming') daysCell.font = { size: 10, color: { argb: 'FF1D4ED8' } };
        });
      }

      ws.addRow([]); // area separator
    }

    // ── Generate and download ─────────────────────────────────────────────────
    const buffer = await wb.xlsx.writeBuffer();
    const blob   = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url    = URL.createObjectURL(blob);
    const a      = document.createElement('a');
    a.href       = url;
    a.download   = `Plazos_LaPampina_${D.meta.dataWeek || fmtDate(cutDate).replace(/\//g,'-')}.xlsx`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);

  } catch(e) {
    console.error('exportPlazosXLSX error:', e);
    alert('Error al exportar: ' + e.message);
  } finally {
    if (btn) { btn.disabled = false; btn.innerHTML = origHtml; }
  }
}

function togglePlCard(hdr) {
  const body    = hdr.nextElementSibling;
  const chevron = hdr.querySelector('.pl-chevron');
  const isOpen  = body.style.display !== 'none';
  body.style.display = isOpen ? 'none' : 'block';
  chevron.classList.toggle('bi-chevron-right', isOpen);
  chevron.classList.toggle('bi-chevron-down',  !isOpen);
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
function devClass(v) { return v==null?'dev-neutral':v<-0.00001?'dev-neg':v>0.00001?'dev-pos':'dev-neutral'; }
function statusBadge(r) {
  const cls = { completed:'badge-completed', inProgress:'badge-inProgress',
                late:'badge-late', notStarted:'badge-notStarted' };
  const key = { completed:'status.completed', inProgress:'status.inProgress',
                late:'status.late', notStarted:'status.notStarted' };
  return `<span class="badge ${cls[r.status]||'badge-notStarted'}">${t(key[r.status]||'status.notStarted')}</span>`;
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

// ── PDF Export ────────────────────────────────────────────────────────────────
function setupPdfExports() {
  on('scurvePdfBtn',            'click', exportScurvePDF);
  on('cronogramaPdfBtn',        'click', () => _exportCronogramaPDFBase(false));
  on('cronogramaPdfBtnResumido','click', () => _exportCronogramaPDFBase(true));
}

/** Draw a branded header bar; returns the Y coordinate below it (mm). */
function _pdfHeader(doc, title, pageW, margin) {
  doc.setFillColor(0, 57, 115);
  doc.rect(0, 0, pageW, 19, 'F');
  doc.setTextColor(255, 255, 255);
  doc.setFont('helvetica', 'bold');
  doc.setFontSize(12);
  doc.text('PowerChina · La Pampina', margin, 7.5);
  doc.setFont('helvetica', 'normal');
  doc.setFontSize(9.5);
  doc.text(title, margin, 14);
  doc.setTextColor(0, 0, 0);
  return 23;
}

/** Print one line of project metadata (date / % plan / % real / deviation). */
function _pdfMeta(doc, y, margin, pageW) {
  if (typeof D === 'undefined' || !D || !D.meta) return y;
  const m = D.meta;
  const dev = m.pctReal - m.pctPlan;
  const parts = [
    t('res.controlDate') + ': ' + fmtDate(m.dataDate) + ' (' + m.dataWeek + ')',
    t('kpi.planned')   + ': ' + pct(m.pctPlan),
    t('kpi.actual')    + ': ' + pct(m.pctReal),
    t('kpi.deviation') + ': ' + signPct(dev),
  ];
  doc.setFontSize(7.5);
  doc.setTextColor(80, 80, 80);
  doc.text(parts.join('   |   '), margin, y);
  // Divider line
  doc.setDrawColor(200, 200, 200);
  doc.line(margin, y + 2.5, pageW - margin, y + 2.5);
  doc.setTextColor(0, 0, 0);
  return y + 7;
}

/** Add page-number footer to every page. */
function _pdfPageNumbers(doc, pageW, pageH, margin) {
  const n = doc.internal.getNumberOfPages();
  const prefix = t('pdf.page');
  for (let i = 1; i <= n; i++) {
    doc.setPage(i);
    doc.setFontSize(7.5);
    doc.setTextColor(160, 160, 160);
    doc.text(`${prefix} ${i} / ${n}`, pageW - margin, pageH - 4, { align: 'right' });
    doc.text('PowerChina · La Pampina', margin, pageH - 4);
  }
}

/** Export both S-Curves (main + filtered) to a landscape A4 PDF. */
async function exportScurvePDF() {
  if (!window.jspdf) {
    showToast('jsPDF not available — check CDN connection', true);
    return;
  }
  const btn = document.getElementById('scurvePdfBtn');
  if (btn) { btn.disabled = true; }
  showToast(t('pdf.generating'), false, 60000);

  try {
    const { jsPDF } = window.jspdf;
    const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();
    const pageH = doc.internal.pageSize.getHeight();
    const margin = 12;

    // ── Page 1: main S-curve ──────────────────────────────────────────────────
    let y = _pdfHeader(doc, t('pdf.scurve'), pageW, margin);
    y     = _pdfMeta(doc, y, margin, pageW);

    const mainCanvas = document.getElementById('scurveChart');
    if (mainCanvas && mainCanvas.width && mainCanvas.height) {
      const imgW  = pageW - 2 * margin;
      const imgH  = imgW * (mainCanvas.height / mainCanvas.width);
      const maxH  = pageH - y - margin - 8;
      doc.addImage(mainCanvas.toDataURL('image/png'), 'PNG', margin, y, imgW, Math.min(imgH, maxH));
    }

    // ── Page 2: filtered S-curve ──────────────────────────────────────────────
    const filtCanvas = document.getElementById('scurveFilteredChart');
    if (filtCanvas && filtCanvas.width && filtCanvas.height) {
      doc.addPage();
      y = _pdfHeader(doc, t('sc.filteredTitle'), pageW, margin);
      y = _pdfMeta(doc, y, margin, pageW);

      // Filter label
      const labelEl = document.getElementById('scurveFilterLabel');
      if (labelEl && labelEl.textContent.trim()) {
        doc.setFontSize(8.5);
        doc.setTextColor(80, 80, 80);
        doc.text(labelEl.textContent.trim(), margin, y);
        y += 6;
        doc.setTextColor(0, 0, 0);
      }

      const imgW  = pageW - 2 * margin;
      const imgH  = imgW * (filtCanvas.height / filtCanvas.width);
      const maxH  = pageH - y - margin - 8;
      doc.addImage(filtCanvas.toDataURL('image/png'), 'PNG', margin, y, imgW, Math.min(imgH, maxH));
    }

    _pdfPageNumbers(doc, pageW, pageH, margin);
    doc.save('CurvaS_LaPampina.pdf');
    showToast(t('pdf.success'));

  } catch (err) {
    console.error('[exportScurvePDF]', err);
    showToast(t('toast.error') + err.message, true);
  } finally {
    if (btn) { btn.disabled = false; }
  }
}

/**
 * DETAILED — all rows matching the current cascade + search filters,
 * expand/collapse state is ignored (PDF shows every record).
 */
function _getArbolPdfRows() {
  if (!D) return [];
  const q       = (document.getElementById('arbolSearch')?.value || '').toLowerCase().trim();
  const allRecs = (_consolTree || D.allRecords).filter(r => r.edt);

  if (q) {
    return allRecs.filter(r =>
      r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)
    );
  }
  if (_wbsFilterEdt) {
    return allRecs.filter(r =>
      r.edt === _wbsFilterEdt || r.edt.startsWith(_wbsFilterEdt + '.')
    );
  }
  return allRecs;
}

/**
 * SUMMARY — only the rows that are currently VISIBLE in the collapsed tree,
 * i.e. same result as what the user sees on screen right now.
 * Applies cascade filter + search + collapsedNodes (via isArbolHidden).
 */
/**
 * SUMMARY — only parent/summary rows (resumen === true or isVirtual).
 * No leaf activities. Gives a compact project-structure overview.
 * Respects cascade area filter and search text.
 */
function _getArbolResumidasRows() {
  if (!D) return [];
  const q       = (document.getElementById('arbolSearch')?.value || '').toLowerCase().trim();
  const allRecs = (_consolTree || D.allRecords).filter(r => r.edt);

  // Keep only parent/summary nodes
  const parents = allRecs.filter(r => r.resumen || r.isVirtual);

  // Apply cascade filter
  const inScope = _wbsFilterEdt
    ? parents.filter(r => r.edt === _wbsFilterEdt || r.edt.startsWith(_wbsFilterEdt + '.'))
    : parents;

  // Apply search filter (if active)
  if (q) {
    return inScope.filter(r =>
      r.tarea.toLowerCase().includes(q) || r.edt.toLowerCase().includes(q)
    );
  }

  return inScope;
}

/** Export the WBS Cronograma as a data-driven table PDF (no screenshot).
 *  @param {boolean} summarized  true → only visible (collapsed) rows; false → all rows
 */
function _exportCronogramaPDFBase(summarized) {
  if (!window.jspdf) {
    showToast('jsPDF not available — check CDN connection', true);
    return;
  }
  if (!D) {
    showToast(t('toast.error') + 'Sem dados carregados', true);
    return;
  }
  const btnId  = summarized ? 'cronogramaPdfBtnResumido' : 'cronogramaPdfBtn';
  const btn    = document.getElementById(btnId);
  if (btn) btn.disabled = true;
  showToast(t('pdf.generating'), false, 60000);

  try {
    const { jsPDF } = window.jspdf;
    const doc   = new jsPDF({ orientation: 'landscape', unit: 'mm', format: 'a4' });
    const pageW = doc.internal.pageSize.getWidth();   // 297 mm
    const pageH = doc.internal.pageSize.getHeight();  // 210 mm
    const ML    = 10;   // left/right margin

    // ── Título muda conforme o modo ───────────────────────────────────────────
    const pdfTitle = t(summarized ? 'pdf.cronogramaResumido' : 'pdf.cronograma');

    // ── Page-1 header + meta ─────────────────────────────────────────────────
    let startY = _pdfHeader(doc, pdfTitle, pageW, ML);
    startY     = _pdfMeta(doc, startY, ML, pageW);

    // ── Active filter label ───────────────────────────────────────────────────
    const q = (document.getElementById('arbolSearch')?.value || '').trim();
    let filterLabel = '';
    if (q) {
      filterLabel = `🔍 "${q}"`;
    } else if (_wbsFilterEdt) {
      const node = (_consolTree || D.allRecords).find(r => r.edt === _wbsFilterEdt);
      filterLabel = `📂 ${_wbsFilterEdt}${node ? ' — ' + node.tarea.trim() : ''}`;
    }
    if (filterLabel) {
      doc.setFontSize(8);
      doc.setTextColor(40, 80, 160);
      doc.text(filterLabel, ML, startY);
      startY += 6;
      doc.setTextColor(0, 0, 0);
    }

    // ── Collect rows (detalhado = tudo; resumido = só visíveis) ──────────────
    const rows = summarized ? _getArbolResumidasRows() : _getArbolPdfRows();

    if (!rows.length) {
      doc.setFontSize(10);
      doc.setTextColor(120);
      doc.text('Sem registros para exportar.', ML, startY + 6);
      doc.save('Cronograma_LaPampina.pdf');
      showToast(t('pdf.success'));
      return;
    }

    // ── Status label helper ───────────────────────────────────────────────────
    const statusLbl = {
      completed: t('status.completed'), inProgress: t('status.inProgress'),
      late: t('status.late'),           notStarted: t('status.notStarted'),
    };

    // Colour palettes (RGB arrays for autoTable)
    const STATUS_CLR = {
      completed:  [22, 130, 60],
      inProgress: [25, 100, 200],
      late:       [200, 40,  40],
      notStarted: [110, 110, 110],
    };

    // Cores para linhas pai (resumen/isVirtual) — escala de cinza por nível.
    //   nivel 1 → cinza escuro  [80,80,80]    texto branco
    //   nivel 2 → cinza médio   [120,120,120] texto branco
    //   nivel 3 → cinza         [160,160,160] texto branco
    //   nivel 4 → cinza médio   [200,200,200] texto escuro
    //   nivel 5+→ cinza claro   [220,220,220] texto escuro
    //   folhas  → branco
    const PAI_BG  = [
      null,              // nivel 0 (unused)
      [80,  80,  80 ],   // nivel 1 — cinza escuro
      [120, 120, 120],   // nivel 2 — cinza médio-escuro
      [160, 160, 160],   // nivel 3 — cinza
      [200, 200, 200],   // nivel 4 — cinza médio
      [220, 220, 220],   // nivel 5 — cinza claro
      [220, 220, 220],   // nivel 6+
    ];
    const PAI_TXT = [
      null,
      [255, 255, 255],   // nivel 1: branco (navy escuro)
      [40,  40,  40 ],   // nivel 2: escuro (cinza médio-escuro)
      [40,  40,  40 ],   // nivel 3: escuro (cinza médio)
      [40,  40,  40 ],   // nivel 4: escuro
      [40,  40,  40 ],   // nivel 5
      [40,  40,  40 ],   // nivel 6+
    ];

    // ── Build table body ──────────────────────────────────────────────────────
    const head = [[
      t('th.edt'), t('th.activity'),
      t('th.hh'), t('th.start'), t('th.end'),
      'PBs', 'PB Plan.', 'PB Av.', 'PB Dev.',
      t('th.incidence'), 'INCD.PLAN', 'INCD.REAL', 'INCID. DESVÍO',
      t('th.pctPlan'), t('th.pctActual'), '% DESV.',
      t('th.status'),
    ]];

    const body = rows.map(r => {
      const hh        = r.hh > 0        ? Math.round(r.hh).toLocaleString('es-CL')       : '—';
      const incid     = r.incidencia > 0 ? (r.incidencia * 100).toFixed(3) + '%'          : '—';
      const pPlan     = r.pctCompPlan != null ? (r.pctCompPlan * 100).toFixed(2) + '%'    : '—';
      const pReal     = r.pctCompReal != null ? (r.pctCompReal * 100).toFixed(2) + '%'    : '—';
      const incdPlan  = r.incidencia > 0
        ? (r.incidencia * (r.pctCompPlan || 0) * 100).toFixed(3) + '%' : '—';
      const incdReal  = r.incidencia > 0
        ? (r.incidencia * (r.pctCompReal || 0) * 100).toFixed(3) + '%' : '—';
      const incidDesv = r.incidencia > 0.0001
        ? ((r.incidencia * (r.desviacion || 0)) >= 0 ? '+' : '')
          + (r.incidencia * (r.desviacion || 0) * 100).toFixed(3) + '%'
        : '—';
      const pctDesv   = r.incidencia > 0.0001
        ? ((r.desviacion || 0) >= 0 ? '+' : '') + ((r.desviacion || 0) * 100).toFixed(2) + '%'
        : '—';
      const hasPB = r.pbTotal != null;
      // Indent activity name by level to show hierarchy
      const indent  = r.nivel > 1 ? '  '.repeat(r.nivel - 1) : '';
      const actName = indent + (r.tarea || '').trim();

      return [
        r.edt || '',
        actName,
        hh,
        fmtDate(r.inicio),
        fmtDate(r.fin),
        hasPB ? String(r.pbTotal)  : '—',
        hasPB ? String(r.pbPlan)   : '—',
        hasPB ? String(r.pbAv)     : '—',
        hasPB && r.pbDev != null ? String(r.pbDev) : '—',
        incid,
        incdPlan,
        incdReal,
        incidDesv,
        pPlan,
        pReal,
        pctDesv,
        statusLbl[r.status] || '—',
      ];
    });

    // ── Total row (row index 0) using D.meta ──────────────────────────────────
    const m = D.meta;
    const wbsTotalRow = [
      '',
      '',
      Math.round(m.totalHH).toLocaleString('es-CL'),
      '—', '—',
      '—', '—', '—', '—',
      '100.000%',
      (m.pctPlan * 100).toFixed(3) + '%',
      (m.pctReal * 100).toFixed(3) + '%',
      ((m.desvio * 100) >= 0 ? '+' : '') + (m.desvio * 100).toFixed(3) + '%',
      (m.pctPlan * 100).toFixed(2) + '%',
      (m.pctReal * 100).toFixed(2) + '%',
      ((m.desvio * 100) >= 0 ? '+' : '') + (m.desvio * 100).toFixed(2) + '%',
      '—',
    ];
    const bodyWithTotal = [wbsTotalRow, ...body];

    // ── Draw autoTable ────────────────────────────────────────────────────────
    doc.autoTable({
      head,
      body: bodyWithTotal,
      startY,
      margin: { top: 23, left: ML, right: ML, bottom: 14 },

      styles: {
        fontSize: 6.2,
        cellPadding: { top: 1.4, right: 2, bottom: 1.4, left: 2 },
        overflow: 'linebreak',
        valign: 'middle',
        lineColor: [210, 218, 230],
        lineWidth: 0.2,
      },
      headStyles: {
        fillColor: [0, 57, 115],
        textColor: 255,
        fontStyle: 'bold',
        fontSize: 6.8,
        halign: 'center',
      },
      alternateRowStyles: { fillColor: [250, 252, 255] },

      // Column widths — 17 cols, total ≈ 274 mm (landscape A4, 277 mm usable)
      columnStyles: {
        0:  { cellWidth: 22,  fontStyle: 'bold', halign: 'left'  },  // EDT
        1:  { cellWidth: 55,  halign: 'left'                      },  // Activity
        2:  { cellWidth: 11,  halign: 'right'                     },  // H-H
        3:  { cellWidth: 16,  halign: 'center'                    },  // Inicio
        4:  { cellWidth: 16,  halign: 'center'                    },  // Fin
        5:  { cellWidth:  9,  halign: 'center'                    },  // PBs
        6:  { cellWidth: 10,  halign: 'center'                    },  // PB Plan.
        7:  { cellWidth: 10,  halign: 'center'                    },  // PB Av.
        8:  { cellWidth: 10,  halign: 'center'                    },  // PB Dev.
        9:  { cellWidth: 13,  halign: 'right'                     },  // Incid.
        10: { cellWidth: 13,  halign: 'right'                     },  // INCD.PLAN
        11: { cellWidth: 13,  halign: 'right'                     },  // INCD.REAL
        12: { cellWidth: 15,  halign: 'right', fontStyle: 'bold'  },  // INCID. DESVÍO
        13: { cellWidth: 13,  halign: 'right'                     },  // % Plan
        14: { cellWidth: 13,  halign: 'right'                     },  // % Real
        15: { cellWidth: 15,  halign: 'right', fontStyle: 'bold'  },  // % DESV.
        16: { cellWidth: 20,  halign: 'center', fontStyle: 'bold' },  // Status
      },

      // didParseCell: hook correto para modificar estilos de célula no autoTable.
      // (willDrawCell é para desenho nativo jsPDF — não altera o estilo da tabela)
      didParseCell(data) {
        if (data.section !== 'body') return;

        // ── Row 0 = total row (green, bold black) ────────────────────────────
        if (data.row.index === 0) {
          data.cell.styles.fillColor = [214, 240, 224];
          data.cell.styles.fontStyle = 'bold';
          data.cell.styles.textColor = [0, 0, 0];
          data.cell.styles.lineColor = [22, 163, 74];
          data.cell.styles.lineWidth = { top: 0.6, bottom: 0.6, left: 0.1, right: 0.1 };
          return;
        }

        // ── Data rows (offset by 1 because row 0 is total) ────────────────────
        const r = rows[data.row.index - 1];
        if (!r) return;

        // ── Linhas pai (resumen / virtual): fundo cinza por nível ────────────
        if (r.resumen || r.isVirtual) {
          const lvl = Math.min(r.nivel || 1, PAI_BG.length - 1);
          data.cell.styles.fillColor = PAI_BG[lvl];
          data.cell.styles.textColor = PAI_TXT[lvl] || [0, 0, 0];
          data.cell.styles.fontStyle = 'bold';
        }

        // ── INCID. DESVÍO (col 12): vermelho / verde ─────────────────────────
        if (data.column.index === 12) {
          const v = (r.incidencia || 0) * (r.desviacion || 0);
          if      (v < -0.00001) data.cell.styles.textColor = [200, 30, 30];
          else if (v >  0.00001) data.cell.styles.textColor = [22, 130, 60];
        }

        // ── % DESV. (col 15): vermelho / verde ───────────────────────────────
        if (data.column.index === 15) {
          const dev = r.desviacion || 0;
          if      (dev < -0.00001) data.cell.styles.textColor = [200, 30, 30];
          else if (dev >  0.00001) data.cell.styles.textColor = [22, 130, 60];
        }

        // ── Status (col 16): cor por estado ──────────────────────────────────
        if (data.column.index === 16) {
          const c = STATUS_CLR[r.status];
          if (c) data.cell.styles.textColor = c;
        }
      },

      // Header bar repeated on every continuation page
      didDrawPage(data) {
        _pdfHeader(doc, pdfTitle, pageW, ML);
        // Footer (page number without total — updated below)
        const pg = doc.internal.getCurrentPageInfo().pageNumber;
        doc.setFontSize(7);
        doc.setTextColor(160, 160, 160);
        doc.text('PowerChina · La Pampina', ML, pageH - 4);
        doc.text(`${t('pdf.page')} ${pg}`, pageW - ML, pageH - 4, { align: 'right' });
        doc.setTextColor(0, 0, 0);
      },
    });

    // Overwrite page numbers with correct X / N total
    const totalPages = doc.internal.getNumberOfPages();
    for (let i = 1; i <= totalPages; i++) {
      doc.setPage(i);
      doc.setFillColor(255, 255, 255);
      doc.rect(pageW - ML - 32, pageH - 7, 34, 6, 'F');
      doc.setFontSize(7);
      doc.setTextColor(160, 160, 160);
      doc.text(`${t('pdf.page')} ${i} / ${totalPages}`, pageW - ML, pageH - 4, { align: 'right' });
      doc.setTextColor(0, 0, 0);
    }

    const fileName = summarized
      ? 'Cronograma_Resumido_LaPampina.pdf'
      : 'Cronograma_LaPampina.pdf';
    doc.save(fileName);
    showToast(t('pdf.success'));

  } catch (err) {
    console.error('[_exportCronogramaPDFBase]', err);
    showToast(t('toast.error') + err.message, true);
  } finally {
    if (btn) btn.disabled = false;
  }
}
