/**
 * Tasador de Alquiler de Local Comercial
 * app.js — lógica principal
 *
 * Requiere: precios.json cargado como window.DATA (inyectado por build_html.py)
 * o importado como módulo en un bundler.
 */

"use strict";

// ── Constantes del modelo (Tabla_calculo_renta_locales.xlsx) ──────────────────
const IVA_DIV   = 1.104;  // divisor para pasar CVB → CVN
const DIAS_DOM  = 350;    // días de venta/año con apertura domingos
const DIAS_NODOM= 308;    // días de venta/año sin domingos
const TE_RIESGO = 0.80;   // umbral superior tasa de esfuerzo
const TE_OPTIMO = 0.50;   // umbral inferior tasa de esfuerzo

// ── Calles prime conocidas: nombre, coeficiente, descripción ─────────────────
// Cobertura limitada y manual (sin fuente de precios por calle real).
// El coeficiente es orientativo: sirve para ajustar el precio del distrito,
// no es un dato medido calle a calle.
const STREETS = [
  { name: "Montera",                coef: "2.2", desc: "eje prime de Centro, Madrid" },
  { name: "Preciados",              coef: "2.2", desc: "eje prime de Centro, Madrid" },
  { name: "Gran Vía",               coef: "2.2", desc: "eje prime, Madrid" },
  { name: "Carmen",                 coef: "1.4", desc: "calle comercial de Centro" },
  { name: "Fuencarral",             coef: "1.4", desc: "calle comercial principal, Madrid" },
  { name: "Serrano",                coef: "2.2", desc: "eje prime de Salamanca, Madrid" },
  { name: "José Ortega y Gasset",   coef: "2.2", desc: "eje prime de Salamanca (lujo)" },
  { name: "Goya",                   coef: "1.4", desc: "calle comercial principal, Salamanca" },
  { name: "Príncipe de Vergara",    coef: "1.4", desc: "calle comercial, Salamanca" },
  { name: "Velázquez",              coef: "1.4", desc: "calle comercial, Salamanca" },
  { name: "Bravo Murillo",          coef: "1.0", desc: "calle media, Tetuán" },
  { name: "Alcalá",                 coef: "1.4", desc: "eje principal, Madrid" },
  { name: "Portal de l'Àngel",      coef: "2.2", desc: "eje prime de Ciutat Vella, Barcelona" },
  { name: "Passeig de Gràcia",      coef: "2.2", desc: "eje prime de Eixample, Barcelona (lujo)" },
  { name: "Rambla de Catalunya",    coef: "1.4", desc: "calle comercial principal, Eixample" },
  { name: "Portaferrissa",          coef: "2.2", desc: "eje prime de Ciutat Vella, Barcelona" },
  { name: "Pelai",                  coef: "2.2", desc: "eje prime, Barcelona" },
  { name: "Diagonal",               coef: "1.4", desc: "eje principal, Barcelona" },
  { name: "La Rambla",              coef: "2.2", desc: "eje prime turístico, Barcelona" },
  { name: "Colón",                  coef: "2.2", desc: "eje prime, Valencia" },
  { name: "Sorní",                  coef: "1.4", desc: "calle comercial, Valencia" },
  { name: "Tetuán",                 coef: "2.2", desc: "eje prime, Sevilla" },
  { name: "Sierpes",                coef: "2.2", desc: "eje prime peatonal, Sevilla" },
  { name: "Gran Vía",               coef: "2.2", desc: "eje prime, Bilbao" },
  { name: "Alameda de Colón",       coef: "1.4", desc: "calle comercial, Málaga" },
  { name: "Larios",                 coef: "2.2", desc: "eje prime peatonal, Málaga" },
];

const EJE_LABEL = {
  "2.2": "eje prime",
  "1.4": "calle principal",
  "1.0": "calle media",
  "0.65": "calle secundaria",
};

// ── Rangos de superficie del fichero ─────────────────────────────────────────
const RANGES = [
  { key: "r1", lo: 300,  hi: 600,  label: "300–600 m²" },
  { key: "r2", lo: 1000, hi: 1400, label: "1.000–1.400 m²" },
  { key: "r3", lo: 1400, hi: 1800, label: "1.400–1.800 m²" },
  { key: "r4", lo: 1800, hi: 2200, label: "1.800–2.200 m²" },
];

// ── Utilidades ────────────────────────────────────────────────────────────────

function norm(s) {
  return (s || "").toString().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s*-\s*a\d\s*$/i, "").trim();
}

/** Acepta 3.500 (miles ES), 7,80 (decimal ES) y 3500 (normal) */
function parseNum(s) {
  if (s == null) return NaN;
  s = s.toString().trim().replace(/[€\s]/g, "");
  if (!s) return NaN;
  if (s.includes(","))                          s = s.replace(/\./g, "").replace(",", ".");
  else if ((s.match(/\./g) || []).length > 1)  s = s.replace(/\./g, "");
  else if (/\.\d{3}$/.test(s))                 s = s.replace(/\./g, "");
  return parseFloat(s);
}

/** Formatea con separador de miles "." y decimales ",", forzado a mano:
 *  toLocaleString("es-ES") en algunos motores no agrupa números de 4 cifras (8750 en vez de 8.750). */
function fmtEs(n, decimals) {
  const neg = n < 0;
  const [intPart, decPart] = Math.abs(n).toFixed(decimals).split(".");
  const grouped = intPart.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return (neg ? "-" : "") + grouped + (decPart ? "," + decPart : "");
}
const eur  = n => fmtEs(n, 0);
const eur2 = n => fmtEs(n, 2);
const pct1 = n => (n * 100).toLocaleString("es-ES", { minimumFractionDigits: 1, maximumFractionDigits: 1 }) + "%";
const clean = name => name.replace(/ - A\d$/i, "");

function pickRange(m2) {
  let best = null, bd = Infinity, inside = false;
  for (const r of RANGES) {
    let d, ins = false;
    if (m2 >= r.lo && m2 <= r.hi) { d = 0; ins = true; }
    else d = Math.min(Math.abs(m2 - r.lo), Math.abs(m2 - r.hi));
    if (d < bd) { bd = d; best = r; inside = ins; }
  }
  return { range: best, inside };
}

// ── Autocomplete ──────────────────────────────────────────────────────────────

let selected = null, hi = -1, lastList = [];

function searchLoc(q) {
  q = norm(q.split("·")[0]);
  if (q.length < 2) return [];
  const out = [], seen = new Set();
  for (const d of window.DATA) {
    const n = norm(d.localizacion);
    if (n.includes(q)) {
      const k = n + d.provincia;
      if (!seen.has(k)) {
        seen.add(k);
        out.push({ d, exact: n === q, starts: n.startsWith(q), has: d.hasData ? 1 : 0 });
      }
    }
  }
  out.sort((a, b) => b.exact - a.exact || b.starts - a.starts || b.has - a.has);
  return out.map(o => o.d);
}

function setSelected(d) {
  selected = d;
  document.getElementById("loc").value = clean(d.localizacion) + " · " + d.provincia;
  const pill = document.getElementById("okpill");
  pill.style.color = d.hasData ? "var(--good)" : "var(--accent-ink)";
  pill.textContent = (d.hasData ? "✓ " : "⚠ ")
    + clean(d.localizacion) + " (" + d.provincia + " · " + d.tipo + ")"
    + (d.hasData ? "" : " — sin precio en el fichero, solo Idealista");
  pill.classList.add("show");
  document.getElementById("sugg").classList.remove("open");
}

function renderSugg(list) {
  const sugg = document.getElementById("sugg");
  if (!list.length) { sugg.classList.remove("open"); return; }
  lastList = list;
  sugg.innerHTML = list.map((d, i) =>
    `<div data-i="${i}"><b>${clean(d.localizacion)}</b><span class="prov">${d.provincia}</span>`
    + `<span class="tp ${d.hasData ? "has" : "no"}">${d.hasData ? d.tipo : "Idealista"}</span></div>`
  ).join("");
  sugg.classList.add("open");
  hi = -1;
  [...sugg.children].forEach(el => el.onclick = () => setSelected(lastList[+el.dataset.i]));
}

function resolveLoc() {
  if (selected) return selected;
  const m = searchLoc(document.getElementById("loc").value);
  if (m.length) { setSelected(m[0]); return m[0]; }
  return null;
}

// ── Autocomplete de calle ──────────────────────────────────────────────────────

let selectedCalle = null, hiCalle = -1, lastListCalle = [];

function searchStreet(q) {
  const nq = norm(q);
  if (nq.length < 2) return [];
  return STREETS
    .map(s => ({ s, n: norm(s.name) }))
    .filter(x => x.n.includes(nq))
    .sort((a, b) => (b.n === nq) - (a.n === nq) || (b.n.startsWith(nq)) - (a.n.startsWith(nq)))
    .map(x => x.s);
}

function setSelectedCalle(s) {
  selectedCalle = s;
  document.getElementById("calle").value = s.name;
  document.getElementById("eje").value = s.coef;
  const hint = document.getElementById("hint");
  hint.textContent = "Detectado: " + s.desc + " → " + EJE_LABEL[s.coef]
    + " (×" + s.coef.replace(".", ",") + "). Puedes cambiarlo.";
  hint.classList.add("show");
  document.getElementById("suggCalle").classList.remove("open");
}

function renderSuggCalle(list) {
  const sugg = document.getElementById("suggCalle");
  if (!list.length) { sugg.classList.remove("open"); return; }
  lastListCalle = list;
  sugg.innerHTML = list.map((s, i) =>
    `<div data-i="${i}"><b>${s.name}</b><span class="prov">${s.desc}</span>`
    + `<span class="tp has">×${s.coef.replace(".", ",")}</span></div>`
  ).join("");
  sugg.classList.add("open");
  hiCalle = -1;
  [...sugg.children].forEach(el => el.onclick = () => setSelectedCalle(lastListCalle[+el.dataset.i]));
}

// ── HTML builders ─────────────────────────────────────────────────────────────

function ideaBlock(sel, calleTxt) {
  const gq = encodeURIComponent("site:idealista.com alquiler locales "
    + (calleTxt ? calleTxt + " " : "") + clean(sel.localizacion) + " " + sel.provincia);
  return `<div class="idea">
    <div class="ttl">Oferta real en <b>Idealista</b> — locales en alquiler en ${clean(sel.localizacion)}${calleTxt ? " · " + calleTxt : ""}</div>
    <a href="${sel.idealista}" target="_blank" rel="noopener">Ver locales en el distrito →</a>
    <a class="ghost" href="https://www.google.com/search?q=${gq}" target="_blank" rel="noopener">Buscar por calle</a>
  </div>`;
}

function esfuerzoBlock(rentMonthly) {
  const cvb = parseNum(document.getElementById("ventas").value);
  if (!(cvb > 0)) return "";
  const conIVA  = document.getElementById("ivaSi").classList.contains("on");
  const abreDom = document.getElementById("domSi").classList.contains("on");
  const cvn          = conIVA ? cvb / IVA_DIV : cvb;
  const dias         = abreDom ? DIAS_DOM : DIAS_NODOM;
  const ventaDiaria  = cvn / dias;
  const rentAnnual   = rentMonthly * 12;
  const te           = rentMonthly / ventaDiaria;
  const pctVtaNeta   = rentAnnual / cvn;
  let val, col;
  if      (te > TE_RIESGO) { val = "Riesgo"; col = "var(--bad)"; }
  else if (te < TE_OPTIMO) { val = "Óptimo"; col = "var(--good)"; }
  else                      { val = "Viable"; col = "var(--ok)"; }
  return `
  <div class="esf" style="border-color:${col}">
    <div class="eh">
      <div>
        <div class="et">Tasa de esfuerzo · renta mensual / venta media diaria</div>
        <div class="ebig" style="color:${col}">${pct1(te)}</div>
      </div>
      <div class="eval" style="background:${col};color:var(--text)">${val}</div>
    </div>
    <div class="stats">
      <div class="stat"><div class="k">Renta anual</div><div class="v">${eur(rentAnnual)}<small> €</small></div></div>
      <div class="stat"><div class="k">Renta mensual</div><div class="v">${eur(rentMonthly)}<small> €</small></div></div>
      <div class="stat"><div class="k">Venta media diaria</div><div class="v">${eur(ventaDiaria)}<small> €/día</small></div></div>
      <div class="stat"><div class="k">Renta s/ venta neta</div><div class="v">${pct1(pctVtaNeta)}</div></div>
    </div>
    <div class="stats">
      <div class="stat"><div class="k">Ventas ${conIVA ? "(con IVA)" : "(netas)"}</div><div class="v" style="font-size:16px">${eur(cvb)}<small> €</small></div></div>
      <div class="stat"><div class="k">Ventas netas (÷1,104)</div><div class="v" style="font-size:16px">${eur(cvn)}<small> €</small></div></div>
      <div class="stat"><div class="k">Días de venta/año</div><div class="v">${dias}<small> ${abreDom ? "c/dom" : "s/dom"}</small></div></div>
      <div class="stat"><div class="k">Equivale a</div><div class="v" style="font-size:15px">${te.toLocaleString("es-ES", { maximumFractionDigits: 2 })}<small> días vta/mes</small></div></div>
    </div>
    <div class="legend">
      <span><b style="color:var(--good)">&lt;50%</b> Óptimo</span>
      <span><b style="color:var(--ok)">50–80%</b> Viable</span>
      <span><b style="color:var(--bad)">&gt;80%</b> Riesgo</span>
      <span style="color:var(--muted)">Metodología: Tabla cálculo renta locales</span>
    </div>
  </div>`;
}

// ── Calculadora: precio de venta → renta objetivo ────────────────────────────

function calcYield() {
  const out    = document.getElementById("yieldOut");
  const precio = parseNum(document.getElementById("ventaPrecio").value);
  const pct    = parseNum(document.getElementById("rentabPct").value);
  if (!(precio > 0) || !(pct > 0)) { out.innerHTML = ""; return; }

  const anual   = precio * (pct / 100);
  const mensual = anual / 12;
  const pctTxt  = pct.toLocaleString("es-ES", { maximumFractionDigits: 2 });

  out.innerHTML = `
    <div class="stats" style="margin-top:14px">
      <div class="stat"><div class="k">Alquiler anual</div><div class="v">${eur(anual)}<small> €/año</small></div></div>
      <div class="stat"><div class="k">Alquiler mensual</div><div class="v">${eur(mensual)}<small> €/mes</small></div></div>
    </div>
    <div class="note">
      <b>Cálculo:</b> el ${pctTxt}% de ${eur(precio)} € son <b>${eur(anual)} €</b> al año (antes de impuestos y gastos)
      → alquiler de <b>${eur(mensual)} €/mes</b> (IVA no incluido, si procede).
    </div>
    <button type="button" class="usebtn" id="useYieldRent">Usar esta renta (${eur(mensual)} €/mes) en el comparador ↑</button>`;

  document.getElementById("useYieldRent").onclick = () => {
    document.getElementById("mTotal").click();
    document.getElementById("rent").value = String(Math.round(mensual));
    document.getElementById("rent").focus();
    document.getElementById("rent").scrollIntoView({ behavior: "smooth", block: "center" });
  };
}

// ── Cálculo principal ─────────────────────────────────────────────────────────

function calcular() {
  document.getElementById("err").style.display = "none";
  const sel     = resolveLoc();
  const m2      = parseNum(document.getElementById("m2").value);
  const rentVal = parseNum(document.getElementById("rent").value);

  function showErr(t) { const e = document.getElementById("err"); e.textContent = t; e.style.display = "block"; }
  if (!sel)        { showErr("Elige una ubicación de la lista (municipio o distrito)."); document.getElementById("loc").focus(); return; }
  if (!(m2 > 0))   { showErr("Introduce los metros cuadrados del local (ej. 450).");    document.getElementById("m2").focus();  return; }
  if (!(rentVal > 0)) { showErr("Introduce la renta que te piden (ej. 3.500).");        document.getElementById("rent").focus(); return; }

  const unitMode = document.getElementById("mUnit").classList.contains("on");
  const unit  = unitMode ? rentVal : rentVal / m2;
  const total = unitMode ? rentVal * m2 : rentVal;
  const anual = total * 12;
  const calleTxt  = document.getElementById("calle").value.trim();
  const { range, inside } = pickRange(m2);
  const base  = sel[range.key];
  const esf   = esfuerzoBlock(total);
  const result = document.getElementById("result");

  // Municipio sin datos de precio
  if (!sel.hasData || !base) {
    result.innerHTML = `
      <div class="verdict" style="border-color:var(--line)">
        <span class="tag">${clean(sel.localizacion)}${calleTxt ? " · " + calleTxt : ""} · ${range.label}</span>
        <h2 style="color:var(--soft)">Sin precio de referencia</h2>
        <p>El fichero no incluye precio de local comercial para <b>${clean(sel.localizacion)}</b>. Compara tu renta con la oferta actual en Idealista.</p>
      </div>
      <div class="pricebox">
        <div class="pb">
          <div class="k">Tu renta</div>
          <div class="big">${eur2(unit)}<small> €/m²/mes</small></div>
          <div class="src">${eur(total)} €/mes · <b style="color:var(--text)">${eur(anual)} €/año</b> · ${eur(m2)} m²</div>
        </div>
        <div class="pb live">
          <div>
            <div class="k">Precio medio de la zona</div>
            <div class="big" style="color:var(--idealista)">en vivo</div>
            <div class="src">Fuente fiable y actual: Idealista</div>
          </div>
          <a href="${sel.idealista}" target="_blank" rel="noopener">Ver precio medio en Idealista →</a>
        </div>
      </div>
      ${esf}
      ${ideaBlock(sel, calleTxt)}`;
    result.style.display = "block";
    result.scrollIntoView({ behavior: "smooth", block: "nearest" });
    return;
  }

  // Con datos del fichero
  const coef    = parseFloat(document.getElementById("eje").value);
  const coefLbl = EJE_LABEL[document.getElementById("eje").value];
  const min = base.min * coef, medio = base.medio * coef, max = base.max * coef;

  let v, color, msg;
  if      (unit < min)    { v = "Por debajo de mercado";    color = "var(--good)"; msg = "Buen precio: por debajo de la horquilla para esta calle y superficie."; }
  else if (unit <= medio) { v = "Buen precio";              color = "var(--good)"; msg = "Dentro de mercado y por debajo de la media ajustada a la calle. Encaja bien."; }
  else if (unit <= max)   { v = "En mercado";               color = "var(--ok)";   msg = "En mercado, en la parte alta de la horquilla. Margen para negociar."; }
  else                    { v = "Por encima de mercado";    color = "var(--bad)";  msg = "Caro: supera el máximo estimado para esta calle y superficie. Conviene negociar."; }

  const lo = min * 0.9, hiv = max * 1.1;
  const pos = Math.max(0, Math.min(1, (unit - lo) / (hiv - lo))) * 100;
  const mp  = Math.max(0, Math.min(1, (medio - lo) / (hiv - lo))) * 100;
  const expMed = medio * m2, expMin = min * m2, expMax = max * m2;
  const diff = total - expMed, diffPct = diff / expMed * 100;
  const adjLine = coef === 1
    ? "Referencia del distrito sin ajuste de calle."
    : `Base del distrito <b>${eur2(base.medio)}</b> €/m²/mes × <b>×${document.getElementById("eje").value.replace(".", ",")}</b> (${coefLbl}) = <b>${eur2(medio)}</b> €/m²/mes ajustado a la calle.`;
  const rangeNote = inside ? "" :
    `<div class="note">La superficie de ${eur(m2)} m² queda fuera de los rangos medidos; uso el más cercano (<b>${range.label}</b>). Resultado orientativo.</div>`;

  result.innerHTML = `
    <div class="verdict" style="background:linear-gradient(180deg,rgba(20,30,45,.03),transparent);border-color:${color}">
      <span class="tag"><span class="vchip" style="background:${color}"></span>${clean(sel.localizacion)}${calleTxt ? " · " + calleTxt : ""} · ${coefLbl} · ${range.label}</span>
      <h2 style="color:${color}">${v}</h2><p>${msg}</p>
      <div class="bar">
        <div class="mk" style="left:${mp}%" data-l="media"></div>
        <div class="mk" style="left:${pos}%;background:${color}" data-l="tú · ${eur2(unit)} €/m²"></div>
      </div>
      <div class="scale"><span>${eur2(min)} €/m²</span><span>${eur2(max)} €/m²</span></div>
    </div>
    <div class="pricebox">
      <div class="pb">
        <div class="k">Precio medio de referencia${calleTxt && coef !== 1 ? " (calle)" : ""}</div>
        <div class="big">${eur2(medio)}<small> €/m²/mes</small></div>
        <div class="src">Fuente: fichero de referencia (2009)${coef !== 1 ? " · ajustado ×" + document.getElementById("eje").value.replace(".", ",") : ""} · ${clean(sel.localizacion)}, ${range.label}</div>
      </div>
      <div class="pb live">
        <div>
          <div class="k">Precio medio actual</div>
          <div class="big" style="color:var(--idealista)">en vivo</div>
          <div class="src">Contrasta el dato de 2009 con el mercado de hoy en Idealista</div>
        </div>
        <a href="${sel.idealista}" target="_blank" rel="noopener">Ver precio medio en Idealista →</a>
      </div>
    </div>
    <div class="adj">${adjLine}</div>
    <div class="stats">
      <div class="stat"><div class="k">Tu renta /m²</div><div class="v">${eur2(unit)}<small> €/m²/mes</small></div></div>
      <div class="stat"><div class="k">Media ajustada</div><div class="v">${eur2(medio)}<small> €/m²/mes</small></div></div>
      <div class="stat"><div class="k">Horquilla calle</div><div class="v" style="font-size:15px">${eur2(min)}–${eur2(max)}<small> €/m²</small></div></div>
      <div class="stat"><div class="k">vs. media</div><div class="v" style="color:${diff > 0 ? "var(--bad)" : "var(--good)"}">${diff > 0 ? "+" : ""}${diffPct.toFixed(0)}%</div></div>
    </div>
    <div class="stats">
      <div class="stat"><div class="k">Tu renta mensual</div><div class="v">${eur(total)}<small> €/mes</small></div></div>
      <div class="stat"><div class="k">Tu renta anual</div><div class="v">${eur(anual)}<small> €/año</small></div></div>
      <div class="stat"><div class="k">Esperable (media)</div><div class="v">${eur(expMed)}<small> €/mes</small></div></div>
      <div class="stat"><div class="k">Diferencia</div><div class="v" style="color:${diff > 0 ? "var(--bad)" : "var(--good)"}">${diff > 0 ? "+" : ""}${eur(diff)}<small> €/mes</small></div></div>
    </div>
    ${esf}
    ${ideaBlock(sel, calleTxt)}
    ${rangeNote}`;

  result.style.display = "block";
  result.scrollIntoView({ behavior: "smooth", block: "nearest" });
}

// ── Init ──────────────────────────────────────────────────────────────────────

function limpiar() {
  ["loc", "m2", "calle", "rent", "ventas", "ventaPrecio"].forEach(id => document.getElementById(id).value = "");
  document.getElementById("eje").value = "1.0";
  document.getElementById("rentabPct").value = "5";
  selected = null;
  selectedCalle = null;
  document.getElementById("okpill").classList.remove("show");
  document.getElementById("hint").classList.remove("show");
  document.getElementById("err").style.display = "none";
  document.getElementById("result").style.display = "none";
  document.getElementById("result").innerHTML = "";
  document.getElementById("yieldOut").innerHTML = "";
  document.getElementById("sugg").classList.remove("open");
  document.getElementById("suggCalle").classList.remove("open");
  document.getElementById("mTotal").click();
  document.getElementById("ivaSi").click();
  document.getElementById("domNo").click();
  document.getElementById("loc").focus();
}

document.addEventListener("DOMContentLoaded", () => {
  const locEl = document.getElementById("loc");
  const sugg  = document.getElementById("sugg");

  locEl.addEventListener("input", () => {
    selected = null;
    document.getElementById("okpill").classList.remove("show");
    renderSugg(searchLoc(locEl.value).slice(0, 50));
  });
  locEl.addEventListener("keydown", e => {
    if (!sugg.classList.contains("open")) return;
    const items = [...sugg.children];
    if      (e.key === "ArrowDown")  { hi = Math.min(hi + 1, items.length - 1); e.preventDefault(); }
    else if (e.key === "ArrowUp")    { hi = Math.max(hi - 1, 0); e.preventDefault(); }
    else if (e.key === "Enter")      { if (hi >= 0) { setSelected(lastList[hi]); e.preventDefault(); } return; }
    else return;
    items.forEach((el, i) => el.classList.toggle("hi", i === hi));
  });
  locEl.addEventListener("blur", () => { setTimeout(() => { if (!selected) resolveLoc(); }, 150); });

  // calles conocidas (autocompletado real, igual que la ubicación)
  const calleEl    = document.getElementById("calle");
  const suggCalle  = document.getElementById("suggCalle");

  calleEl.addEventListener("input", () => {
    selectedCalle = null;
    document.getElementById("hint").classList.remove("show");
    renderSuggCalle(searchStreet(calleEl.value));
  });
  calleEl.addEventListener("keydown", e => {
    if (!suggCalle.classList.contains("open")) return;
    const items = [...suggCalle.children];
    if      (e.key === "ArrowDown") { hiCalle = Math.min(hiCalle + 1, items.length - 1); e.preventDefault(); }
    else if (e.key === "ArrowUp")   { hiCalle = Math.max(hiCalle - 1, 0); e.preventDefault(); }
    else if (e.key === "Enter")     { if (hiCalle >= 0) { setSelectedCalle(lastListCalle[hiCalle]); e.preventDefault(); } return; }
    else return;
    items.forEach((el, i) => el.classList.toggle("hi", i === hiCalle));
  });

  document.addEventListener("click", e => {
    if (!e.target.closest("#loc")   && !e.target.closest("#sugg"))       sugg.classList.remove("open");
    if (!e.target.closest("#calle") && !e.target.closest("#suggCalle")) suggCalle.classList.remove("open");
  });

  // calculadora precio de venta → renta
  document.getElementById("ventaPrecio").addEventListener("input", calcYield);
  document.getElementById("rentabPct").addEventListener("input", calcYield);

  // segmented
  const seg = (a, b, onA, onB) => {
    document.getElementById(a).onclick = () => { onA(); document.getElementById(a).classList.add("on"); document.getElementById(b).classList.remove("on"); };
    document.getElementById(b).onclick = () => { onB(); document.getElementById(b).classList.add("on"); document.getElementById(a).classList.remove("on"); };
  };
  seg("mTotal", "mUnit",
    () => document.getElementById("rent").placeholder = "Ej. 3.500",
    () => document.getElementById("rent").placeholder = "Ej. 7,80");
  seg("ivaSi", "ivaNo", () => {}, () => {});
  seg("domNo", "domSi", () => {}, () => {});

  document.getElementById("go").onclick    = calcular;
  document.getElementById("clear").onclick = limpiar;
});
