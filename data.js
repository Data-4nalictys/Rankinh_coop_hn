/**
 * data.js — Carga los 3 CSV y los procesa
 * LOCAL: lee desde data/ | PRODUCCIÓN: lee desde GitHub Raw
 */

const BASE = 'https://raw.githubusercontent.com/Data-4nalictys/Rankinh_coop_hn/main/data/';

// ── Parser CSV robusto (maneja comas dentro de comillas) ──
function parseCSV(text) {
  text = text.replace(/^\uFEFF/, '');
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (!lines.length) return [];

  function parseLine(line) {
    const vals = [];
    let cur = '', inQ = false;
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') { inQ = !inQ; continue; }
      if (c === ',' && !inQ) { vals.push(cur); cur = ''; continue; }
      cur += c;
    }
    vals.push(cur);
    return vals;
  }

  // Headers: NO hacer trim para preservar espacios como en 'Total Afiliados '
  const rawHeaders = parseLine(lines[0]);
  const headers = rawHeaders.map(h => h.replace(/^\uFEFF/, '')); // solo quitar BOM

  return lines.slice(1).map(line => {
    const vals = parseLine(line);
    const row = {};
    headers.forEach((h, i) => { row[h] = (vals[i] || '').trim(); });
    return row;
  }).filter(r => Object.values(r).some(v => v));
}

// ── Caché de sesión ────────────────────────────────────────
async function fetchCSV(filename) {
  if (_cache[filename]) return _cache[filename];

  const url = BASE + filename + '?v=' + Date.now();

  const res = await fetch(url);

  if (!res.ok)
    throw new Error('No se pudo cargar ' + filename + ' (' + res.status + ')');

  const data = parseCSV(await res.text());

  _cache[filename] = data;

  return data;
}

// ── Número entero seguro ───────────────────────────────────
function toNum(v) {
  if (!v && v !== 0) return 0;
  const n = parseInt(String(v).replace(/,/g, '').trim(), 10);
  return isNaN(n) ? 0 : n;
}

// ── Periodo YYYY-MM ────────────────────────────────────────
function mkPeriodo(anio, mes) {
  return String(anio || '').trim().padStart(4, '0') + '-' + String(mes || '').trim().padStart(2, '0');
}

// ── Buscar campo ignorando espacios al inicio/fin ──────────
// Necesario para 'Total Afiliados ' que tiene espacio al final
function getField(row, name) {
  // Primero buscar exacto
  if (row[name] !== undefined) return row[name];
  // Luego buscar ignorando espacios
  const nameTrim = name.trim();
  for (const key of Object.keys(row)) {
    if (key.trim() === nameTrim) return row[key];
  }
  return '';
}

// ════════════════════════════════════════════════════════════
// BALANCE GENERAL
// ════════════════════════════════════════════════════════════
async function loadBalanceGeneral() {
  const rows = await fetchCSV('Balance_General.csv');
  return rows
    .filter(r => r['COOPERATIVA'] && !r['COOPERATIVA'].toUpperCase().includes('FACACH'))
    .map(r => ({
      coop:        r['COOPERATIVA'].trim(),
      periodo:     mkPeriodo(r['AÑO'], r['MES#']),
      activos_net: toNum(r['ACTIVOS NETOS']),
      activos_tot: toNum(r['ACTIVOS TOTALES']),
      prestamos:   toNum(r['PRESTAMOS']),
      pasivos:     toNum(r['PASIVOS TOTALES']),
      obligaciones:toNum(r['OBLIGACIONES BANCARIAS']),
      patrimonio:  toNum(r['PATRIMONIO']),
      provisiones: toNum(r['PROVISIONES DE VALUACION']),
    }));
}

// ════════════════════════════════════════════════════════════
// ESTADO DE RESULTADOS
// ════════════════════════════════════════════════════════════
async function loadEstadoResultados() {
  const rows = await fetchCSV('Estado_de_Resultados.csv');
  return rows
    .filter(r => r['COOPERATIVA'] && !r['COOPERATIVA'].toUpperCase().includes('FACACH'))
    .map(r => ({
      coop:       r['COOPERATIVA'].trim(),
      periodo:    mkPeriodo(r['AÑO'], r['MES#']),
      ingresos:   toNum(r['INGRESOS']),
      egresos:    toNum(r['EGRESOS']),
      excedentes: toNum(r['EXCEDENTES']),
      prod_fin:   toNum(r['Productos Financieros']),
      gast_fin:   toNum(r['Gastos Financieros']),
      gast_adm:   toNum(r['Gastos de Administracion']),
    }));
}

// ════════════════════════════════════════════════════════════
// AFILIADOS
// Nota: 'Total Afiliados ' tiene espacio al final — usar getField()
// Nota: 'Cierre' ya viene como 'YYYY-MM'
// ════════════════════════════════════════════════════════════
async function loadAfiliados() {
  const rows = await fetchCSV('Afiliados.csv');
  return rows
    .filter(r => r['Cooperativa'] && !r['Cooperativa'].toUpperCase().includes('FACACH'))
    .map(r => ({
      coop:      r['Cooperativa'].trim(),
      periodo:   r['Cierre'].trim(),
      total:     toNum(getField(r, 'Total Afiliados')),
      hombres:   toNum(r['Afiliado Hombres']),
      mujeres:   toNum(r['Afiliados Mujeres']),
      menores:   toNum(r['Menores Ahorrantes']),
      juridicas: toNum(r['Personas Jurídicas']),
    }));
}

// ════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════

function agrupar(rows, camposNum) {
  const map = {};
  rows.forEach(r => {
    const key = r.coop + '|' + r.periodo;
    if (!map[key]) {
      map[key] = { coop: r.coop, periodo: r.periodo };
      camposNum.forEach(c => map[key][c] = 0);
    }
    camposNum.forEach(c => map[key][c] += (r[c] || 0));
  });
  return Object.values(map).sort((a, b) => a.periodo.localeCompare(b.periodo));
}

function periodos(rows) {
  return [...new Set(rows.map(r => r.periodo))].sort();
}

function ultimoPeriodo(rows) {
  return periodos(rows).slice(-1)[0] || '';
}

function top20(rows, periodo, campo) {
  return rows
    .filter(r => r.periodo === periodo)
    .sort((a, b) => b[campo] - a[campo])
    .slice(0, 20);
}

function buildIndex(rows) {
  const idx = {};
  rows.forEach(r => {
    if (!idx[r.periodo]) idx[r.periodo] = {};
    idx[r.periodo][r.coop] = r;
  });
  return idx;
}

window.RD = {
  loadBalanceGeneral,
  loadEstadoResultados,
  loadAfiliados,
  agrupar, periodos, ultimoPeriodo, top20, buildIndex,
};
