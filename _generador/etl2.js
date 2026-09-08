const XLSX = require('xlsx');
const fs = require('fs');
const dir = "C:/Users/Estiven/OneDrive - PYC Consultoria Integral SAS/Estiven/Sotracar/";
const files = {
 2020:"LIBRO_AUXILIAR_SOTRACAR - (año 2020) 2026-09-01T121531.985.xlsx",
 2021:"LIBRO_AUXILIAR_SOTRACAR - (2021) 2026-09-01T121557.849.xlsx",
 2022:"LIBRO_AUXILIAR_SOTRACAR - (2022) 2026-09-01T121627.890.xlsx",
 2023:"LIBRO_AUXILIAR_SOTRACAR - (2023) 2026-09-01T121655.045.xlsx",
 2024:"LIBRO_AUXILIAR_SOTRACAR -(2024)  2026-09-01T121728.762.xlsx",
 2025:"LIBRO_AUXILIAR_SOTRACAR -(año 2025) 2026-08-26T161646.856.xlsx",
 2026:"LIBRO_AUXILIAR_SOTRACAR - (año 2026) 2026-08-26T161714.548.xlsx",
};
const YEARS = Object.keys(files).map(Number).sort();
const codeRe = /^(\d{6,10})\s+(.+)$/;
const R = v => Math.round((+v || 0) * 100) / 100;
const R0 = v => Math.round(+v || 0);
const Z = () => Array(12).fill(0);
const grp = c => c.slice(0, 2);
const acc4 = c => c.slice(0, 4);
const cls = c => c[0];

// ---------- 1. PARSE ----------
const cuentas = {};                       // code -> name
const mov = {};                           // year -> code -> [12][deb,cre]  (real movements, excl saldo inicial)
const monthEndRaw = {};                   // year -> code -> [12] (month-end running saldo; null where unknown before carry)
const openBal = {};                       // year -> code -> opening saldo (from SALDO INICIAL)
const terceros = {};                      // year -> code -> key -> {nombre,nit,deb,cre}
const lastTxnMonth = {};                  // year -> last month (1-12) with a real movement

for (const y of YEARS) {
  const wb = XLSX.readFile(dir + files[y], { cellDates: true });
  const rows = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  mov[y] = {}; monthEndRaw[y] = {}; openBal[y] = {}; terceros[y] = {};
  let cur = null, curEnd = null;
  for (let i = 2; i < rows.length; i++) {
    const r = rows[i];
    const a = r[0];
    if (a == null || a === '') continue;
    const s = String(a).trim();
    if (s.startsWith('Total Movimientos')) {
      if (cur) {
        for (let m = 0; m < 12; m++) if (curEnd[m] == null) curEnd[m] = (m > 0 ? curEnd[m - 1] : openBal[y][cur]);
        monthEndRaw[y][cur] = curEnd.map(v => v == null ? null : R(v));
      }
      cur = null; continue;
    }
    const m = s.match(codeRe);
    if (m && !(r[6] instanceof Date)) {
      cur = m[1];
      cuentas[cur] = cuentas[cur] || m[2].trim();
      mov[y][cur] = Array.from({ length: 12 }, () => [0, 0]);
      terceros[y][cur] = {};
      openBal[y][cur] = 0;
      curEnd = Array(12).fill(null);
      continue;
    }
    if (!cur) continue;
    const fecha = r[6];
    const obs = (r[7] == null ? '' : String(r[7])).trim().toUpperCase();
    const deb = +r[14] || 0, cre = +r[15] || 0;
    const saldo = (r[16] == null || r[16] === '') ? null : +r[16];
    if (obs === 'SALDO INICIAL') { openBal[y][cur] = saldo != null ? saldo : (deb - cre); continue; }
    if (!(fecha instanceof Date)) continue;
    let mo = fecha.getUTCMonth();
    if (fecha.getUTCFullYear() < y) mo = 0;
    else if (fecha.getUTCFullYear() > y) mo = 11;
    mo = Math.max(0, Math.min(11, mo));
    mov[y][cur][mo][0] = R(mov[y][cur][mo][0] + deb);
    mov[y][cur][mo][1] = R(mov[y][cur][mo][1] + cre);
    if (saldo != null) curEnd[mo] = saldo;
    if (deb || cre) lastTxnMonth[y] = Math.max(lastTxnMonth[y] || 0, mo + 1);
    const tn = (r[0] != null ? String(r[0]) : '').trim();
    const nit = r[2] != null ? String(r[2]).trim() : '';
    if (tn) {
      const k = nit || tn;
      const t = terceros[y][cur][k] || (terceros[y][cur][k] = { nombre: tn, nit, deb: 0, cre: 0 });
      t.deb = R(t.deb + deb); t.cre = R(t.cre + cre);
      if (tn.length > t.nombre.length) t.nombre = tn;
    }
  }
  if (cur && curEnd) {
    for (let m = 0; m < 12; m++) if (curEnd[m] == null) curEnd[m] = (m > 0 ? curEnd[m - 1] : openBal[y][cur]);
    monthEndRaw[y][cur] = curEnd.map(v => v == null ? null : R(v));
  }
}

// ---------- 2. CROSS-YEAR BALANCE CHAINING (class 1,2,3) ----------
// monthEnd[y][code] = full 12-month series, carrying forward prior year close when the account is absent.
const monthEnd = {};
// seed: earliest recorded opening balance per balance-sheet account, so pre-2020 static
// assets/liabilities that had no movement in early years are still reflected (approx).
const carry = {};
for (const y of YEARS) for (const c of Object.keys(openBal[y])) {
  if (['1', '2', '3'].includes(cls(c)) && !(c in carry) && Math.abs(openBal[y][c]) > 0) carry[c] = openBal[y][c];
}
const seededCodes = new Set(Object.keys(carry));
for (const y of YEARS) {
  monthEnd[y] = {};
  const present = new Set(Object.keys(monthEndRaw[y]));
  // accounts present this year
  for (const c of present) {
    if (!['1', '2', '3'].includes(cls(c))) continue;
    const ser = monthEndRaw[y][c].slice();
    for (let m = 0; m < 12; m++) if (ser[m] == null) ser[m] = m > 0 ? ser[m - 1] : (carry[c] ?? openBal[y][c] ?? 0);
    monthEnd[y][c] = ser.map(R);
    carry[c] = ser[11];
  }
  // accounts known from prior years but absent now -> flat carry
  for (const c of Object.keys(carry)) {
    if (present.has(c)) continue;
    if (!['1', '2', '3'].includes(cls(c))) continue;
    monthEnd[y][c] = Array(12).fill(R(carry[c]));
  }
}

// ---------- 3. HELPERS ----------
const isDep = c => ['5160', '5165', '6160', '6165', '7160', '7165'].includes(acc4(c));
const sumA = a => R(a.reduce((x, y) => x + y, 0));
function movSeries(y, code, sign) {
  const mm = mov[y][code]; const o = Z();
  if (!mm) return o;
  for (let m = 0; m < 12; m++) o[m] = R(sign === 'DC' ? mm[m][0] - mm[m][1] : mm[m][1] - mm[m][0]);
  return o;
}
const addS = (dst, src) => { for (let m = 0; m < 12; m++) dst[m] = R(dst[m] + src[m]); };
function lastIdx(y) { return (lastTxnMonth[y] || 12) - 1; }

// ---------- 4. P&L PER YEAR (standard methodology) ----------
const pyg = {};
for (const y of YEARS) {
  const codes = Object.keys(mov[y]);
  const L = {
    ingresos_op: Z(), devoluciones: Z(), ingresos_noop: Z(),
    costo_servicio: Z(), gasto_admon: Z(), gasto_ventas: Z(),
    gasto_fin: Z(), otros_egresos: Z(), impuesto: Z(), dep: Z(),
  };
  for (const c of codes) {
    const g = grp(c), a4 = acc4(c), k = cls(c);
    if (k === '4') {
      if (a4 === '4175') addS(L.devoluciones, movSeries(y, c, 'DC'));
      else if (g === '41') addS(L.ingresos_op, movSeries(y, c, 'CD'));
      else addS(L.ingresos_noop, movSeries(y, c, 'CD'));
    } else if (k === '6' || k === '7') {
      addS(L.costo_servicio, movSeries(y, c, 'DC'));
      if (isDep(c)) addS(L.dep, movSeries(y, c, 'DC'));
    } else if (k === '5') {
      if (g === '51') addS(L.gasto_admon, movSeries(y, c, 'DC'));
      else if (g === '52') addS(L.gasto_ventas, movSeries(y, c, 'DC'));
      else if (g === '54') addS(L.impuesto, movSeries(y, c, 'DC'));
      else if (a4 === '5305') addS(L.gasto_fin, movSeries(y, c, 'DC'));
      else addS(L.otros_egresos, movSeries(y, c, 'DC'));
      if (isDep(c)) addS(L.dep, movSeries(y, c, 'DC'));
    }
  }
  const ingresos = L.ingresos_op.map((v, m) => R(v - L.devoluciones[m]));
  const util_bruta = ingresos.map((v, m) => R(v - L.costo_servicio[m]));
  const ebit = util_bruta.map((v, m) => R(v - L.gasto_admon[m] - L.gasto_ventas[m]));
  const ebitda = ebit.map((v, m) => R(v + L.dep[m]));
  const uai = ebit.map((v, m) => R(v + L.ingresos_noop[m] - L.gasto_fin[m] - L.otros_egresos[m]));
  const util_neta = uai.map((v, m) => R(v - L.impuesto[m]));
  pyg[y] = {
    ingresos, ingresos_noop: L.ingresos_noop, costo_servicio: L.costo_servicio,
    gasto_admon: L.gasto_admon, gasto_ventas: L.gasto_ventas, gasto_fin: L.gasto_fin,
    otros_egresos: L.otros_egresos, impuesto: L.impuesto, dep: L.dep,
    util_bruta, ebit, ebitda, uai, util_neta,
  };
}

// ---------- 5. BALANCE SHEET PER YEAR ----------
const GRP_AC = ['11', '12', '13', '14'];
const GRP_PC = ['21', '22', '23', '24', '25', '26', '27', '28'];
const bsheet = {};
for (const y of YEARS) {
  const B = {
    activo_corriente: Z(), activo_no_corriente: Z(), pasivo_corriente: Z(), pasivo_no_corriente: Z(),
    patrimonio_libro: Z(), caja: Z(), cartera: Z(), inventario: Z(), ppe: Z(),
    proveedores: Z(), deuda_fin: Z(), impuestos_xp: Z(), laborales: Z(),
  };
  for (const c of Object.keys(monthEnd[y])) {
    const me = monthEnd[y][c]; const g = grp(c), k = cls(c);
    const into = dst => { for (let m = 0; m < 12; m++) dst[m] = R(dst[m] + (me[m] || 0)); };
    if (k === '1') {
      into(GRP_AC.includes(g) ? B.activo_corriente : B.activo_no_corriente);
      if (g === '11') into(B.caja);
      if (g === '13') into(B.cartera);
      if (g === '14') into(B.inventario);
      if (g === '15') into(B.ppe);
    } else if (k === '2') {
      into(GRP_PC.includes(g) ? B.pasivo_corriente : B.pasivo_no_corriente);
      if (g === '21') into(B.deuda_fin);
      if (g === '22') into(B.proveedores);
      if (g === '24') into(B.impuestos_xp);
      if (g === '25' || g === '26') into(B.laborales);
    } else if (k === '3') into(B.patrimonio_libro);
  }
  const activo_total = B.activo_corriente.map((v, m) => R(v + B.activo_no_corriente[m]));
  const pasivo_total = B.pasivo_corriente.map((v, m) => R(v + B.pasivo_no_corriente[m]));
  const patrimonio = activo_total.map((v, m) => R(v - pasivo_total[m]));         // plug
  const capital_trabajo = B.activo_corriente.map((v, m) => R(v - B.pasivo_corriente[m]));
  const deuda_neta = B.deuda_fin.map((v, m) => R(v - B.caja[m]));
  const li0 = lastIdx(y);
  const patLibro = B.patrimonio_libro[li0];
  const descuadre = activo_total[li0] ? Math.abs(activo_total[li0] - pasivo_total[li0] - patLibro) / Math.abs(activo_total[li0]) : 1;
  const confiable = descuadre < 0.05 && B.activo_no_corriente[li0] > 0;
  bsheet[y] = { ...B, activo_total, pasivo_total, patrimonio, capital_trabajo, deuda_neta, confiable, descuadre: R(descuadre * 1e4) / 1e4 };
}

// ---------- 6. ACCOUNT DETAIL (top-N + "otras") ----------
function topMov(y, filt, sign, n) {
  const list = Object.keys(mov[y]).filter(filt).map(c => {
    const vals = movSeries(y, c, sign);
    return { codigo: c, nombre: cuentas[c], vals, total: sumA(vals) };
  }).filter(x => Math.abs(x.total) > 1).sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
  const top = list.slice(0, n), rest = list.slice(n);
  const otras = Z(); rest.forEach(x => addS(otras, x.vals));
  const total = Z(); list.forEach(x => addS(total, x.vals));
  return { top, otras, otras_count: rest.length, total };
}
function topBal(y, filt, n) {
  const list = Object.keys(monthEnd[y]).filter(filt).map(c => {
    const vals = monthEnd[y][c].map(v => R(v || 0));
    let li = 11; for (let i = 11; i >= 0; i--) if (vals[i]) { li = i; break; }
    return { codigo: c, nombre: cuentas[c], vals, total: vals[Math.min(li, lastIdx(y))] || vals[li] || 0 };
  }).filter(x => x.vals.some(v => Math.abs(v) > 1)).sort((a, b) => Math.abs(b.total) - Math.abs(a.total));
  const top = list.slice(0, n), rest = list.slice(n);
  const otras = Z(); rest.forEach(x => addS(otras, x.vals));
  const total = Z(); list.forEach(x => addS(total, x.vals));
  return { top, otras, otras_count: rest.length, total };
}
const detalle = {};
for (const y of YEARS) {
  detalle[y] = {
    ingresos_op: topMov(y, c => grp(c) === '41', 'CD', 8),
    costo_servicio: topMov(y, c => cls(c) === '6' || cls(c) === '7', 'DC', 10),
    gasto_admon: topMov(y, c => grp(c) === '51', 'DC', 10),
    gasto_fin: topMov(y, c => acc4(c) === '5305', 'DC', 8),
    otros_egresos: topMov(y, c => grp(c) === '53' && acc4(c) !== '5305', 'DC', 6),
    caja: topBal(y, c => grp(c) === '11', 8),
    cartera: topBal(y, c => grp(c) === '13', 8),
    ppe: topBal(y, c => grp(c) === '15' || grp(c) === '16', 8),
    proveedores: topBal(y, c => grp(c) === '22', 6),
    deuda_fin: topBal(y, c => grp(c) === '21', 10),
    laborales: topBal(y, c => grp(c) === '25' || grp(c) === '26', 8),
    impuestos_xp: topBal(y, c => grp(c) === '24', 8),
  };
}

// ---------- 7. TERCEROS ----------
function topTerc(y, filt, sign, n) {
  const agg = {};
  for (const c of Object.keys(terceros[y])) {
    if (!filt(c)) continue;
    for (const [k, t] of Object.entries(terceros[y][c])) {
      const key = t.nit || k;
      const a = agg[key] || (agg[key] = { nombre: t.nombre, nit: t.nit, v: 0 });
      a.v = R(a.v + (sign === 'DC' ? t.deb - t.cre : t.cre - t.deb));
      if (t.nombre.length > a.nombre.length) a.nombre = t.nombre;
    }
  }
  const arr = Object.values(agg).filter(a => Math.abs(a.v) > 1).sort((x, z) => z.v - x.v);
  const top = arr.slice(0, n).map(a => ({ nombre: a.nombre, nit: a.nit, valor: R0(a.v) }));
  const total = R0(arr.reduce((s, a) => s + a.v, 0));
  return { top, total, n_total: arr.length };
}
const tercAgg = {};
for (const y of YEARS) {
  tercAgg[y] = {
    clientes: topTerc(y, c => grp(c) === '41', 'CD', 15),
    proveedores: topTerc(y, c => grp(c) === '22' || cls(c) === '6', 'DC', 15),
    acreedores_fin: topTerc(y, c => grp(c) === '21', 'CD', 12),
  };
}

// ---------- 8. RATIOS ----------
const ratios = {};
for (const y of YEARS) {
  const p = pyg[y], b = bsheet[y];
  const li = lastIdx(y);
  const meses = li + 1;
  const ing = sumA(p.ingresos);
  const ub = sumA(p.util_bruta), eb = sumA(p.ebit), ebd = sumA(p.ebitda), un = sumA(p.util_neta), ua = sumA(p.uai);
  const at = b.activo_total[li], pt = b.pasivo_total[li], pat = b.patrimonio[li];
  const ac = b.activo_corriente[li], pc = b.pasivo_corriente[li];
  const cart = b.cartera[li], prov = b.proveedores[li], df = b.deuda_fin[li], caja = b.caja[li];
  const gf = sumA(p.gasto_fin), costo = sumA(p.costo_servicio);
  const annualize = v => R(v * 12 / meses);
  ratios[y] = {
    meses, ultimo_mes: li + 1, balance_confiable: b.confiable, balance_descuadre: b.descuadre,
    ingresos: ing, ingresos_anualizado: annualize(ing),
    util_bruta: ub, ebit: eb, ebitda: ebd, uai: ua, util_neta: un,
    dep: sumA(p.dep), gasto_admon: sumA(p.gasto_admon), gasto_fin: gf,
    costo_servicio: costo, impuesto: sumA(p.impuesto), ingresos_noop: sumA(p.ingresos_noop), otros_egresos: sumA(p.otros_egresos),
    margen_bruto: ing ? R(ub / ing * 1e4) / 1e4 : 0,
    margen_op: ing ? R(eb / ing * 1e4) / 1e4 : 0,
    margen_ebitda: ing ? R(ebd / ing * 1e4) / 1e4 : 0,
    margen_neto: ing ? R(un / ing * 1e4) / 1e4 : 0,
    activo_total: R0(at), pasivo_total: R0(pt), patrimonio: R0(pat),
    activo_corriente: R0(ac), pasivo_corriente: R0(pc),
    caja_fin: R0(caja), cartera_fin: R0(cart), proveedores_fin: R0(prov), deuda_fin: R0(df),
    deuda_neta: R0(df - caja),
    razon_corriente: pc ? R(ac / pc * 1e4) / 1e4 : 0,
    capital_trabajo: R0(ac - pc),
    endeudamiento: at ? R(pt / at * 1e4) / 1e4 : 0,
    apalancamiento: pat ? R(pt / pat * 1e4) / 1e4 : 0,
    deuda_ebitda: ebd ? R((df - caja) / annualize(ebd) * 100) / 100 : 0,
    cobertura_int: gf ? R(eb / gf * 100) / 100 : 0,
    roe: pat ? R(annualize(un) / pat * 1e4) / 1e4 : 0,
    roa: at ? R(annualize(un) / at * 1e4) / 1e4 : 0,
    dias_cartera: ing ? R0(cart / annualize(ing) * 365) : 0,
    dias_proveedor: costo ? R0(prov / annualize(costo) * 365) : 0,
    carga_financiera: ing ? R(gf / ing * 1e4) / 1e4 : 0,
  };
}

// ---------- 9. OUTPUT ----------
const out = {
  meta: {
    generado: new Date().toISOString().slice(0, 10),
    años: YEARS,
    ultimo_mes: lastTxnMonth,
    n_cuentas: Object.keys(cuentas).length,
    fuente: "Libro Auxiliar Sotracar 2020–2026 · mensualizado desde movimientos débito/crédito",
    metodologia: "Cifras calculadas directamente del libro auxiliar (PUC). Depreciación y amortización (5160/5165) tratadas como gasto operativo; EBITDA = utilidad operacional + D&A. Balance encadenado entre años arrastrando el saldo de cierre de cuentas sin movimiento. Patrimonio = activo − pasivo.",
  },
  cuentas,
  meses: ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"],
  pyg,
  bsheet_serie: (() => {
    const o = {};
    for (const y of YEARS) {
      const b = bsheet[y];
      o[y] = {
        activo_corriente: b.activo_corriente, activo_no_corriente: b.activo_no_corriente, activo_total: b.activo_total,
        pasivo_corriente: b.pasivo_corriente, pasivo_no_corriente: b.pasivo_no_corriente, pasivo_total: b.pasivo_total,
        patrimonio: b.patrimonio, caja: b.caja, cartera: b.cartera, ppe: b.ppe,
        proveedores: b.proveedores, deuda_fin: b.deuda_fin, laborales: b.laborales, impuestos_xp: b.impuestos_xp,
        capital_trabajo: b.capital_trabajo, deuda_neta: b.deuda_neta,
        confiable: b.confiable, descuadre: b.descuadre,
      };
    }
    return o;
  })(),
  detalle, ratios, terceros: tercAgg,
};
fs.writeFileSync("datos.json", JSON.stringify(out));
console.log("datos.json", (fs.statSync("datos.json").size / 1024).toFixed(0) + " KB\n");

const M = v => (v / 1e6).toLocaleString('en', { maximumFractionDigits: 0 }).padStart(9);
console.log("AÑO    INGRESOS   U.BRUTA  %BR     EBIT   EBITDA   D&A    GTO.FIN   U.NETA  %NETA");
for (const y of YEARS) {
  const r = ratios[y];
  console.log(y, M(r.ingresos), M(r.util_bruta), (r.margen_bruto * 100).toFixed(1).padStart(5),
    M(r.ebit), M(r.ebitda), M(r.dep), M(r.gasto_fin), M(r.util_neta), (r.margen_neto * 100).toFixed(1).padStart(5));
}
console.log("\nAÑO    ACTIVO   PASIVO  PATRIM   A.CTE  P.CTE  RAZ.CTE  CAP.TRAB   CAJA  DEUDA.FIN D.NETA/EBITDA END%");
for (const y of YEARS) {
  const r = ratios[y];
  console.log(y, M(r.activo_total), M(r.pasivo_total), M(r.patrimonio), M(r.activo_corriente), M(r.pasivo_corriente),
    r.razon_corriente.toFixed(2).padStart(6), M(r.capital_trabajo), M(r.caja_fin), M(r.deuda_fin),
    r.deuda_ebitda.toFixed(2).padStart(6), (r.endeudamiento * 100).toFixed(0).padStart(5),
    "  desc=" + (r.balance_descuadre * 100).toFixed(1) + "% " + (r.balance_confiable ? "OK" : "parcial"));
}
console.log("\nTerceros 2025 — top 3 clientes:", tercAgg[2025].clientes.top.slice(0, 3).map(t => t.nombre + " " + (t.valor / 1e6).toFixed(0) + "M"));
console.log("Cuadre balance (activo vs pasivo+patrim libro) 2025:",
  (bsheet[2025].activo_total[11] / 1e6).toFixed(0), "vs",
  ((bsheet[2025].pasivo_total[11] + bsheet[2025].patrimonio_libro[11]) / 1e6).toFixed(0));
