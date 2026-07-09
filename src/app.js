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

// ── Calles prime conocidas → [coeficiente, descripción] ──────────────────────
const STREETS = {
  "montera":              ["2.2", "eje prime de Centro, Madrid"],
  "preciados":            ["2.2", "eje prime de Centro, Madrid"],
  "gran via":             ["2.2", "eje prime, Madrid"],
  "carmen":               ["1.4", "calle comercial de Centro"],
  "fuencarral":           ["1.4", "calle comercial principal, Madrid"],
  "serrano":              ["2.2", "eje prime de Salamanca, Madrid"],
  "jose ortega y gasset": ["2.2", "eje prime de Salamanca (lujo)"],
  "ortega y gasset":      ["2.2", "eje prime de Salamanca (lujo)"],
  "goya":                 ["1.4", "calle comercial principal, Salamanca"],
  "principe de vergara":  ["1.4", "calle comercial, Salamanca"],
  "velazquez":            ["1.4", "calle comercial, Salamanca"],
  "bravo murillo":        ["1.0", "calle media, Tetuán"],
  "alcala":               ["1.4", "eje principal, Madrid"],
  "portal de l angel":    ["2.2", "eje prime de Ciutat Vella, Barcelona"],
  "portal del angel":     ["2.2", "eje prime de Ciutat Vella, Barcelona"],
  "passeig de gracia":    ["2.2", "eje prime de Eixample, Barcelona (lujo)"],
  "paseo de gracia":      ["2.2", "eje prime de Eixample, Barcelona (lujo)"],
  "rambla catalunya":     ["1.4", "calle comercial principal, Eixample"],
  "portaferrissa":        ["2.2", "eje prime de Ciutat Vella, Barcelona"],
  "pelai":                ["2.2", "eje prime, Barcelona"],
  "pelayo":               ["2.2", "eje prime, Barcelona"],
  "diagonal":             ["1.4", "eje principal, Barcelona"],
  "la rambla":            ["2.2", "eje prime turístico, Barcelona"],
  "ramblas":              ["2.2", "eje prime turístico, Barcelona"],
};

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

function streetKey(s) {
  return norm(s)
    .replace(/[''`]/g, " ")
    .replace(/\b(calle|c|carrer|avenida|avda|av|passeig|paseo|plaza|pza|de|del|la|las|los|el|i)\b/g, " ")
    .replace(/\s+/g, " ").trim();
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

const eur  = n => n.toLocaleString("es-ES", { maximumFractionDigits: 0 });
const eur2 = n => n.toLocaleString("es-ES", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
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
  pill.style.color = d.hasData ? "var(--good)" : "#c9a23a";
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
      <div class="eval" style="background:${col};color:#10161c">${val}</div>
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
      <span style="color:#5d7384">Metodología: Tabla cálculo renta locales</span>
    </div>
  </div>`;
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
    <div class="verdict" style="background:linear-gradient(180deg,rgba(255,255,255,.02),transparent);border-color:${color}">
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
  ["loc", "m2", "calle", "rent", "ventas"].forEach(id => document.getElementById(id).value = "");
  document.getElementById("eje").value = "1.0";
  selected = null;
  document.getElementById("okpill").classList.remove("show");
  document.getElementById("hint").classList.remove("show");
  document.getElementById("err").style.display = "none";
  document.getElementById("result").style.display = "none";
  document.getElementById("result").innerHTML = "";
  document.getElementById("sugg").classList.remove("open");
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
  document.addEventListener("click", e => { if (!e.target.closest(".autobox")) sugg.classList.remove("open"); });

  // calles conocidas
  document.getElementById("calle").addEventListener("input", () => {
    const k = streetKey(document.getElementById("calle").value);
    const hint = document.getElementById("hint");
    hint.classList.remove("show");
    if (k.length < 3) return;
    for (const key in STREETS) {
      if (k.includes(key)) {
        const [coef, desc] = STREETS[key];
        document.getElementById("eje").value = coef;
        hint.textContent = "Detectado: " + desc + " → " + EJE_LABEL[coef]
          + " (×" + coef.replace(".", ",") + "). Puedes cambiarlo.";
        hint.classList.add("show");
        return;
      }
    }
  });

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
