/**
 * data.js — Carga los CSV y los procesa
 */

const BASE = 'https://raw.githubusercontent.com/Data-4nalictys/Rankinh_coop_hn/main/data/';

const _cache = {};

// ════════════════════════════════════════════════════════════
// PARSER CSV ROBUSTO
// ════════════════════════════════════════════════════════════

function parseCSV(text) {

  text = text.replace(/^\uFEFF/, '');

  const lines = text
    .split(/\r?\n/)
    .filter(l => l.trim());

  if (!lines.length) return [];

  function parseLine(line) {

    const vals = [];

    let cur = '';
    let inQ = false;

    for (let i = 0; i < line.length; i++) {

      const c = line[i];

      if (c === '"') {
        inQ = !inQ;
        continue;
      }

      if (c === ',' && !inQ) {
        vals.push(cur);
        cur = '';
        continue;
      }

      cur += c;
    }

    vals.push(cur);

    return vals;
  }

  const headers = parseLine(lines[0]).map(h =>
    h.replace(/^\uFEFF/, '')
  );

  return lines.slice(1).map(line => {

    const vals = parseLine(line);

    const row = {};

    headers.forEach((h, i) => {
      row[h] = (vals[i] || '').trim();
    });

    return row;

  }).filter(r =>
    Object.values(r).some(v => v)
  );
}

// ════════════════════════════════════════════════════════════
// FETCH CSV
// ════════════════════════════════════════════════════════════

async function fetchCSV(filename) {

  if (_cache[filename])
    return _cache[filename];

  const url =
    BASE + filename + '?v=' + new Date().getTime();

  console.log('Cargando:', url);

  const res = await fetch(url);

  console.log('Status:', res.status);

  if (!res.ok) {
    throw new Error(
      'No se pudo cargar ' +
      filename +
      ' (' + res.status + ')'
    );
  }

  const data =
    parseCSV(await res.text());

  _cache[filename] = data;

  return data;
}

// ════════════════════════════════════════════════════════════
// NUMÉRICO
// ════════════════════════════════════════════════════════════

function toNum(v) {

  if (!v && v !== 0)
    return 0;

  const n = parseFloat(
    String(v)
      .replace(/,/g, '')
      .replace('%', '')
      .trim()
  );

  return isNaN(n) ? 0 : n;
}

// ════════════════════════════════════════════════════════════
// PERIODO
// ════════════════════════════════════════════════════════════

function mkPeriodo(anio, mes) {

  return String(anio || '')
      .trim()
      .padStart(4, '0')

      +

      '-'

      +

      String(mes || '')
      .trim()
      .padStart(2, '0');
}

// ════════════════════════════════════════════════════════════
// GET FIELD FLEXIBLE
// ════════════════════════════════════════════════════════════

function getField(row, name) {

  if (row[name] !== undefined)
    return row[name];

  const target = name.trim();

  for (const k of Object.keys(row)) {

    if (k.trim() === target)
      return row[k];
  }

  return '';
}

// ════════════════════════════════════════════════════════════
// BALANCE GENERAL
// ════════════════════════════════════════════════════════════

async function loadBalanceGeneral() {

  const rows =
    await fetchCSV('Balance_General.csv');

  return rows

    .filter(r =>
      r['COOPERATIVA'] &&
      !r['COOPERATIVA']
        .toUpperCase()
        .includes('FACACH')
    )

    .map(r => ({

      coop:
        r['COOPERATIVA'].trim(),

      periodo:
        mkPeriodo(r['AÑO'], r['MES#']),

      activos_net:
        toNum(r['ACTIVOS NETOS']),

      activos_tot:
        toNum(r['ACTIVOS TOTALES']),

      prestamos:
        toNum(r['PRESTAMOS']),

      pasivos:
        toNum(r['PASIVOS TOTALES']),

      obligaciones:
        toNum(r['OBLIGACIONES BANCARIAS']),

      patrimonio:
        toNum(r['PATRIMONIO']),

      provisiones:
        toNum(r['PROVISIONES DE VALUACION']),
    }));
}

// ════════════════════════════════════════════════════════════
// ESTADO RESULTADOS
// ════════════════════════════════════════════════════════════

async function loadEstadoResultados() {

  const rows =
    await fetchCSV('Estado_de_Resultados.csv');

  return rows

    .filter(r =>
      r['COOPERATIVA'] &&
      !r['COOPERATIVA']
        .toUpperCase()
        .includes('FACACH')
    )

    .map(r => ({

      coop:
        r['COOPERATIVA'].trim(),

      periodo:
        mkPeriodo(r['AÑO'], r['MES#']),

      ingresos:
        toNum(r['INGRESOS']),

      egresos:
        toNum(r['EGRESOS']),

      excedentes:
        toNum(r['EXCEDENTES']),

      prod_fin:
        toNum(r['Productos Financieros']),

      gast_fin:
        toNum(r['Gastos Financieros']),

      gast_adm:
        toNum(r['Gastos de Administracion']),
    }));
}

// ════════════════════════════════════════════════════════════
// AFILIADOS
// ════════════════════════════════════════════════════════════

async function loadAfiliados() {

  const rows =
    await fetchCSV('Afiliados.csv');

  return rows

    .filter(r =>
      r['Cooperativa'] &&
      !r['Cooperativa']
        .toUpperCase()
        .includes('FACACH')
    )

    .map(r => ({

      coop:
        r['Cooperativa'].trim(),

      periodo:
        r['Cierre'].trim(),

      total:
        toNum(getField(r, 'Total Afiliados')),

      hombres:
        toNum(r['Afiliado Hombres']),

      mujeres:
        toNum(r['Afiliados Mujeres']),

      menores:
        toNum(r['Menores Ahorrantes']),

      juridicas:
        toNum(r['Personas Jurídicas']),
    }));
}

// ════════════════════════════════════════════════════════════
// UTILIDADES
// ════════════════════════════════════════════════════════════

function agrupar(rows, camposNum) {

  const map = {};

  rows.forEach(r => {

    const key =
      r.coop + '|' + r.periodo;

    if (!map[key]) {

      map[key] = {
        coop: r.coop,
        periodo: r.periodo
      };

      camposNum.forEach(c =>
        map[key][c] = 0
      );
    }

    camposNum.forEach(c =>
      map[key][c] += (r[c] || 0)
    );
  });

  return Object.values(map)
    .sort((a,b)=>
      a.periodo.localeCompare(b.periodo)
    );
}

function periodos(rows) {

  return [...new Set(
    rows.map(r => r.periodo)
  )].sort();
}

function ultimoPeriodo(rows) {

  return periodos(rows)
    .slice(-1)[0] || '';
}

function top20(rows, periodo, campo) {

  return rows
    .filter(r => r.periodo === periodo)
    .sort((a,b)=>b[campo]-a[campo])
    .slice(0,20);
}

function buildIndex(rows) {

  const idx = {};

  rows.forEach(r => {

    if (!idx[r.periodo])
      idx[r.periodo] = {};

    idx[r.periodo][r.coop] = r;
  });

  return idx;
}

// ════════════════════════════════════════════════════════════
// EXPORT
// ════════════════════════════════════════════════════════════

window.RD = {
  loadBalanceGeneral,
  loadEstadoResultados,
  loadAfiliados,
  agrupar,
  periodos,
  ultimoPeriodo,
  top20,
  buildIndex,
};