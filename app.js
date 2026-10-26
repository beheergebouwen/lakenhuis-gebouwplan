
(() => {
let ORGS = [];   // from SharePoint list LK_Organisaties (name + colour)
const USES = ['Verhuurbaar', 'Gemeenschappelijk', 'Technisch / berging', 'Buitenruimte'];
const USEV = { 'Verhuurbaar': '--v', 'Gemeenschappelijk': '--g', 'Technisch / berging': '--t', 'Buitenruimte': '--x' };
const USEC = { 'Verhuurbaar': 'u-v', 'Gemeenschappelijk': 'u-g', 'Technisch / berging': 'u-t', 'Buitenruimte': 'u-x' };
const SRCS = ['Op plan', 'Berekend uit maten', 'Te meten', 'Gemeten'];
const LOCKS = ['', 'Cilinder', 'Badgelezer', 'Cilinder + badgelezer', 'Codeklavier', 'Geen slot', 'Onbekend'];
const $ = (s, el = document) => el.querySelector(s);
const css = v => getComputedStyle(document.documentElement).getPropertyValue(v).trim();
const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const eur = n => (n == null || isNaN(n)) ? '–' : n.toLocaleString('nl-BE', { style: 'currency', currency: 'EUR' });
const m2f = n => (n == null || n === '' || isNaN(n)) ? '–' : Number(n).toLocaleString('nl-BE', { maximumFractionDigits: 1 }) + ' m²';
const short = c => String(c).replace(/^LK-/, '');
const orgColor = o => { const x = ORGS.find(g => g.name === o); return x ? x.color : null; };

let GEO, canWrite = true; const IMG = {};
// ---------- bezetting (wie zat waar, van wanneer tot wanneer)
let OCC = [], COSTS = [], VOS = [], IMPS = [];                                    // {_id, room, org, from, to, note}  (datums YYYY-MM-DD, leeg = onbekend / lopend)
const ymd = d => new Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Brussels' }).format(d);
const dnum = s => Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) / 864e5;
const addDays = (s, n) => new Date((dnum(s) + n) * 864e5).toISOString().slice(0, 10);
const dfmt = s => s ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : '';
let peil = ymd(new Date());
const occOf = code => OCC.filter(o => o.room === code).sort((a, b) => (b.from || '').localeCompare(a.from || ''));
const isOpen = (o, d) => (!o.from || o.from <= d) && (!o.to || o.to >= d);
function orgAt(code, d) { const o = occOf(code).find(o => isOpen(o, d)); return o ? o.org : ''; }
function orgOf(r) { if (!r) return ''; return OCC.some(o => o.room === r.code) ? orgAt(r.code, peil) : (r.org || ''); }
function orgDays(r, a, b) {                      // dagen per organisatie binnen [a, b]
  const out = {}; const rows = occOf(r.code);
  if (!rows.length) { if (r.org) out[r.org] = dnum(b) - dnum(a) + 1; return out; }
  for (const o of rows) { const s = Math.max(dnum(a), o.from ? dnum(o.from) : -1e9), e = Math.min(dnum(b), o.to ? dnum(o.to) : 1e9); if (e >= s) out[o.org] = (out[o.org] || 0) + e - s + 1; }
  return out;
}
function istaPeriod() { const m = String(S.settings.period || '').match(/(\d\d)\/(\d\d)\/(\d{4}).*?(\d\d)\/(\d\d)\/(\d{4})/); return m ? { from: `${m[3]}-${m[2]}-${m[1]}`, to: `${m[6]}-${m[5]}-${m[4]}` } : { from: peil.slice(0, 4) + '-01-01', to: peil.slice(0, 4) + '-12-31' }; }
let per = null;
const S = { rooms: {}, doors: {}, meters: {}, settings: {} };      // seed baseline overlaid with live docs
let floor = 'f2', sel = null, zoom = 1;
try { const t = localStorage.getItem('lk.floor'); if (t) floor = t; } catch (e) {}

// ---------- data
function setStatus(kind, text) { $('#dot').className = 'dot ' + kind; $('#statusText').textContent = text; }
async function boot() {
  GEO = await fetch('geo.json?t=' + Date.now(), { cache: 'no-store' }).then(r => r.json());
  if (!GEO.floors.some(f => f.id === floor)) floor = 'f2';
  setStatus('', 'Aanmelden bij Microsoft 365…');
  const me = await LK.login(); if (!me) return;            // redirecting to login
  $('#who').textContent = me.name; $('#logout').hidden = false; $('#logout').onclick = () => LK.logout();
  setStatus('', 'Gegevens laden uit SharePoint…');
  await reload(true);
  buildStatic(); renderAll();
  await Promise.all(GEO.floors.map(async f => { try { IMG[f.id] = await LK.image(f.img); } catch (e) { IMG[f.id] = ''; } }));
  renderPlan();
  setStatus('live', 'Verbonden met SharePoint · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' }));
  setInterval(() => { if (!document.hidden && !document.activeElement?.closest('form,select,input,textarea')) reload().then(() => renderAll(true)); }, 60000);
}
async function reload(first) {
  const [rooms, doors, meters, orgs, settings, occ, costs, vo, imp] = await Promise.all(['rooms', 'doors', 'meters', 'orgs', 'settings', 'occ', 'costs', 'vo', 'imp'].map(k => LK.load(k).catch(() => null)));
  COSTS = costs || []; VOS = (vo || []).filter(v => v.org); IMPS = imp || [];
  OCC = (occ || []).filter(o => o.room && o.org);
  if (!rooms) { setStatus('ro', 'Lijsten niet gevonden'); const n = $('#notice'); n.hidden = false; n.textContent = 'De SharePoint-lijsten bestaan nog niet. Laat IT eenmalig setup/ uitvoeren (zie handleiding).'; throw new Error('no lists'); }
  S.rooms = Object.fromEntries(rooms.map(r => [r.code, r])); S.doors = Object.fromEntries(doors.map(d => [d.code, d]));
  S.meters = Object.fromEntries(meters.map(m => [m.nr, m]));
  ORGS = (orgs || []).filter(o => o.name).map(o => ({ _id: o._id, name: o.name, color: o.color || '#e6e2d6', type: o.type || 'Huurder' })).sort((a, b) => a.name.localeCompare(b.name));
  S.settings = (settings || []).find(x => x.key === 'ista') || S.settings;
  for (const k of ['heatCost', 'totalUnits', 'rentTotal', 'meterCount']) S.settings[k] = Number(S.settings[k]) || 0;
  if (!per) { const y = +peil.slice(0, 4) - 1; per = { from: y + '-01-01', to: y + '-12-31' }; }   // standaard: vorig kalenderjaar
  if (!first) setStatus('live', 'Verbonden met SharePoint · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' }));
}
async function save(col, id, body) {
  const kind = col; const cur = col === 'occ' ? body : (S[col][id] || {});
  const obj = { ...body, _id: cur._id };
  try { await LK.save(kind, obj); if (col === 'occ') body._id = obj._id; }
  catch (e) { if (e.status === 403 || e.status === 401) { const er = new Error('geen bewerkrechten op de SharePoint-lijst'); er.code = 'permission_denied'; throw er; } throw e; }
  setStatus('live', 'Opgeslagen in SharePoint · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' }));
}

// ---------- derived
const m2 = r => (r.m2meas != null && r.m2meas !== '') ? Number(r.m2meas) : (r.m2plan != null ? Number(r.m2plan) : null);
const price = () => S.settings.totalUnits ? S.settings.heatCost / S.settings.totalUnits : 0;
const rent = () => S.settings.meterCount ? S.settings.rentTotal / S.settings.meterCount : 0;
const meterCost = m => (Number(m.units) || 0) * price() + rent();
function metersOf(code) { return Object.values(S.meters).filter(m => m.room === code); }
function roomOf(code) { return S.rooms[code]; }
function floorLabel(id) { return GEO.floors.find(f => f.id === id)?.label || id; }

// ---------- static UI
function buildStatic() {
  const fr = $('#floors');
  fr.innerHTML = GEO.floors.map(f => `<button data-f="${f.id}" aria-pressed="${f.id === floor}"><b>${esc(f.label)}</b><span>${esc(f.sub)}</span></button>`).join('');
  fr.onclick = e => { const b = e.target.closest('button'); if (!b) return; floor = b.dataset.f; sel = null; try { localStorage.setItem('lk.floor', floor); } catch (er) {} renderPlan(); renderPanel(); };
  const fopts = '<option value="">Alle verdiepingen</option>' + GEO.floors.map(f => `<option value="${f.id}">${esc(f.label)}</option>`).join('');
  $('#rFloor').innerHTML = fopts; $('#dFloor').innerHTML = fopts;
  $('#rOrg').innerHTML = '<option value="">Alle organisaties</option><option value="-">Zonder organisatie</option>' + ORGS.map(o => `<option>${esc(o.name)}</option>`).join('');
  ['#rFloor', '#rOrg', '#rQ'].forEach(s => $(s).oninput = renderRooms);
  ['#dFloor', '#dQ'].forEach(s => $(s).oninput = renderDoors);
  $('#q').oninput = renderPlan; $('#showDoors').onchange = renderPlan;
  $('#zIn').onclick = () => { zoom = Math.min(3, zoom * 1.25); applyZoom(); };
  $('#zOut').onclick = () => { zoom = Math.max(0.5, zoom / 1.25); applyZoom(); };
  $('#zFit').onclick = () => { zoom = 1; applyZoom(); };
  $('#peil').value = peil; $('#peil').onchange = e => { if (e.target.value) { peil = e.target.value; renderAll(); } };
  $('#pFrom').value = per.from; $('#pTo').value = per.to;
  const setPer = () => setPeriod($('#pFrom').value, $('#pTo').value);
  $('#pFrom').onchange = setPer; $('#pTo').onchange = setPer;
  yearSel('#pYear');
  $('#expOcc').onclick = () => exportCsv('Lakenhuis_bezetting.csv', ['ruimtecode', 'ruimte', 'organisatie', 'van', 'tot', 'opmerking'],
    [...OCC].sort((a, b) => a.room.localeCompare(b.room) || (a.from || '').localeCompare(b.from || '')).map(o => [o.room, roomOf(o.room)?.name || '', o.org, dfmt(o.from) || 'onbekend', dfmt(o.to) || 'lopend', o.note]));
  $('#tabs').onclick = e => { const b = e.target.closest('button'); if (!b) return; showView(b.dataset.v); };
  $('#legend').innerHTML = USES.map(u => `<span><i style="border-color:var(${USEV[u]})"></i>${u}</span>`).join('') +
    ORGS.map(o => `<span><i style="border-color:transparent;background:${o.color}"></i>${esc(o.name)}</span>`).join('');
  $('#expRooms').onclick = () => exportCsv('Lakenhuis_ruimten.csv', ['code', 'verdieping', 'ruimte', 'gebruik', 'organisatie', 'm2_plan', 'bron', 'm2_gemeten', 'ok', 'ista_meters', 'verwarming_eur', 'opmerking'],
    Object.values(S.rooms).sort((a, b) => a.code.localeCompare(b.code)).map(r => { const ms = metersOf(r.code); return [r.code, floorLabel(r.floor), r.name, r.use, orgOf(r), r.m2plan ?? '', r.src, r.m2meas ?? '', r.ok ? 'ja' : 'nee', ms.map(m => m.nr).join(' '), ms.reduce((s, m) => s + meterCost(m), 0).toFixed(2), r.note]; }));
  $('#expDoors').onclick = () => exportCsv('Lakenhuis_deuren.csv', ['code', 'verdieping', 'van', 'naar', 'type', 'maat', 'slot_badge', 'sleutel_cilinder', 'gecontroleerd', 'opmerking'],
    Object.values(S.doors).sort((a, b) => a.code.localeCompare(b.code)).map(d => [d.code, floorLabel(d.floor), d.fr, d.to, d.type, d.size, d.lock, d.key, d.checked ? 'ja' : 'nee', d.note]));
  $('#expIsta').onclick = () => exportCsv('Lakenhuis_ista.csv', ['meternummer', 'ista_unit', 'ista_ruimte', 'verbruik_eenheden', 'ruimtecode', 'ruimte', 'organisatie', 'verwarming_eur', 'huur_meter_eur', 'totaal_eur'],
    Object.values(S.meters).map(m => { const r = roomOf(m.room) || {}; return [m.nr, m.unit, m.istaRoom, m.units, m.room, r.name || '', orgOf(r) || '', ((Number(m.units) || 0) * price()).toFixed(2), rent().toFixed(2), meterCost(m).toFixed(2)]; }));
}
function refreshOrgUI() {
  const cur = $('#rOrg').value;
  $('#rOrg').innerHTML = '<option value="">Alle organisaties</option><option value="-">Zonder organisatie</option>' + ORGS.map(o => `<option>${esc(o.name)}</option>`).join('');
  $('#rOrg').value = cur;
  $('#legend').innerHTML = USES.map(u => `<span><i style="border-color:var(${USEV[u]})"></i>${u}</span>`).join('') +
    ORGS.map(o => `<span><i style="border-color:transparent;background:${o.color}"></i>${esc(o.name)}</span>`).join('');
}
async function saveOrg(o, msgEl) {
  try { await LK.save('orgs', o); if (msgEl) { msgEl.className = 'saved'; msgEl.textContent = 'Opgeslagen'; }
    const orgs = await LK.load('orgs'); ORGS = orgs.filter(x => x.name).map(x => ({ _id: x._id, name: x.name, color: x.color || '#e6e2d6', type: x.type || 'Huurder' })).sort((a, b) => a.name.localeCompare(b.name));
    refreshOrgUI(); renderAll(true); }
  catch (e) { if (msgEl) { msgEl.className = 'err'; msgEl.textContent = 'Niet opgeslagen: ' + (e.status === 403 ? 'geen bewerkrechten.' : e.message); } }
}
function showView(v) {
  document.querySelectorAll('#tabs button').forEach(b => b.setAttribute('aria-selected', b.dataset.v === v));
  document.querySelectorAll('section.view').forEach(s => s.hidden = s.id !== 'v-' + v);
}
async function exportCsv(name, head, rows) {
  const q = v => { const s = String(v ?? ''); return /[;"\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; };
  const csv = '﻿' + [head, ...rows].map(r => r.map(q).join(';')).join('\r\n');
  const a = document.createElement('a'); a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv;charset=utf-8' })); a.download = name; document.body.appendChild(a); a.click(); setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
}

// ---------- plan
function applyZoom() { const svg = $('#sheet svg'), sh = $('#sheet'); if (!svg) return; const F = GEO.floors.find(f => f.id === floor); const fitH = Math.max(340, sh.clientHeight - 2), fitW = sh.clientWidth - 2; const base = Math.min(fitW, fitH * F.w / F.h); svg.style.width = (base * zoom) + 'px'; svg.style.margin = zoom <= 1 ? '0 auto' : '0'; }
addEventListener('resize', () => applyZoom());
function renderPlan() {
  document.querySelectorAll('#floors button').forEach(b => b.setAttribute('aria-pressed', b.dataset.f === floor));
  const F = GEO.floors.find(f => f.id === floor);
  const q = $('#q').value.trim().toLowerCase();
  let g = '';
  for (const [id, geo] of Object.entries(GEO.rooms)) {
    if (geo.floor !== floor) continue;
    if (geo.out) { g += `<path class="n out" d="${geo.poly.map(p => 'M' + p.replace(/ /g, 'L') + 'Z').join('')}" fill="${css('--x')}" stroke="${css('--x')}" style="pointer-events:none"/>`; continue; }
    const isRef = !!geo.ref; const code = isRef ? geo.ref : id;
    const r = roomOf(code) || {};
    const stroke = css(USEV[r.use] || '--t');
    const fill = isRef ? css('--ovm') : (orgColor(orgOf(r)) || stroke);
    const hit = !q || [code, r.name, orgOf(r)].join(' ').toLowerCase().includes(q);
    const d = geo.poly.map(p => 'M' + p.replace(/ /g, 'L') + 'Z').join('');
    const lbl = short(code) + (isRef ? '*' : '');
    const tw = lbl.length * 11.6 + 14;
    g += `<g class="room${sel === code ? ' sel' : ''}${hit ? '' : ' dim'}" data-code="${esc(code)}" tabindex="0" role="button" aria-label="${esc(code + ' ' + (r.name || ''))}">
      <path class="${(orgColor(orgOf(r)) || isRef) ? 'o' : 'n'}" d="${d}" fill="${fill}" stroke="${stroke}" ${isRef ? 'stroke-dasharray="10 7"' : ''}/>
      <g class="lbl"><rect x="${geo.lbl[0] - tw / 2}" y="${geo.lbl[1] - 15}" width="${tw}" height="30" rx="6" stroke="${stroke}" stroke-width="2.5" fill="${orgColor(orgOf(r)) || '#fff'}"/>
      <text x="${geo.lbl[0]}" y="${geo.lbl[1] + 7}" text-anchor="middle" fill="${stroke}" style="stroke:none">${esc(lbl)}</text></g></g>`;
  }
  if (F.lift) g += `<g class="lbl"><rect x="${F.lift[0] - 30}" y="${F.lift[1] - 15}" width="60" height="30" rx="6" stroke="${css('--g')}" stroke-width="2.5"/><text x="${F.lift[0]}" y="${F.lift[1] + 7}" text-anchor="middle" fill="${css('--g')}" style="stroke:none">LIFT</text></g>`;
  if ($('#showDoors').checked) for (const [id, dg] of Object.entries(GEO.doors)) {
    if (dg.floor !== floor) continue; const t = 'D-' + short(id.replace(/^D-/, '')).replace(/^/, ''); const lab = id; const w = lab.length * 8.6 + 10;
    g += `<g class="door"><rect x="${dg.pt[0] - w / 2}" y="${dg.pt[1] - 11}" width="${w}" height="22" rx="3"/><text x="${dg.pt[0]}" y="${dg.pt[1] + 5}" text-anchor="middle">${esc(lab)}</text></g>`;
  }
  $('#sheet').innerHTML = `<svg viewBox="0 0 ${F.w} ${F.h}" role="img" aria-label="Plan ${esc(F.label)}"><image href="${IMG[F.id] || ''}" width="${F.w}" height="${F.h}"/>${g}</svg>`;
  applyZoom();
  $('#sheet svg').onclick = e => { const el = e.target.closest('.room'); if (el) select(el.dataset.code); };
  $('#sheet svg').onkeydown = e => { if ((e.key === 'Enter' || e.key === ' ') && e.target.closest('.room')) { e.preventDefault(); select(e.target.closest('.room').dataset.code); } };
}
function select(code) {
  const r = roomOf(code); if (!r) return;
  if (r.floor !== floor) { floor = r.floor; }
  sel = code; showView('plan'); renderPlan(); renderPanel();
  if (window.innerWidth < 860) $('#panel').scrollIntoView({ behavior: 'smooth', block: 'start' });
}

// ---------- panel
function renderPanel() {
  const P = $('#panel');
  if (!sel || !roomOf(sel)) { P.innerHTML = `<p class="empty">Klik op een ruimte in het plan om de gegevens te zien en aan te passen.</p>${floorSummary()}`; return; }
  const r = roomOf(sel), ms = metersOf(sel), doors = Object.values(S.doors).filter(d => (d.fr + ' / ' + d.to).split(' / ').includes(sel));
  const heat = ms.reduce((s, m) => s + meterCost(m), 0);
  const dis = canWrite ? '' : 'disabled';
  P.innerHTML = `
    <div><h2><span class="mono">${esc(r.code)}</span>${esc(r.name)}</h2>
      <div class="chips" style="margin-top:6px"><span class="chip ${USEC[r.use] || ''}">${esc(r.use)}</span><span class="chip">${esc(floorLabel(r.floor))}</span>
      ${orgOf(r) ? `<span class="chip" style="background:${orgColor(orgOf(r)) || 'var(--paper)'};color:#1d2328">${esc(orgOf(r))}</span>` : ''}
      ${r.ok ? '<span class="chip okc">Gecontroleerd</span>' : (r.src === 'Te meten' && r.m2meas == null ? '<span class="chip warn">Te meten</span>' : '')}</div></div>
    <div class="occbox"><div class="sub">Bezetting (historiek)</div>
      ${occOf(r.code).length ? `<div class="occlist">${occOf(r.code).map(o => `<div class="occitem${isOpen(o, peil) ? ' now' : ''}" data-id="${esc(o._id)}">
        <div class="occorg"><span class="swatch" style="background:${orgColor(o.org)}"></span>${esc(o.org)}${isOpen(o, peil) ? ' <span class="chip okc" style="font-size:11px">nu</span>' : ''}${canWrite ? '<button type="button" class="occdel" title="Deze bezetting verwijderen" aria-label="Verwijderen">✕</button><span class="occconf" hidden>Verwijderen? <button type="button" class="btn occyes">Ja</button> <button type="button" class="btn occno">Nee</button></span>' : ''}</div>
        <div class="two"><label>Van<input type="date" data-k="from" value="${esc(o.from)}" ${dis}></label><label>Tot<input type="date" data-k="to" value="${esc(o.to)}" ${dis}></label></div></div>`).join('')}</div>
        <p class="empty" style="margin:4px 0 0;font-size:12px">Leeg 'Van' = datum onbekend · leeg 'Tot' = loopt nog.</p>` : '<p class="empty" style="margin:0;font-size:13px">Nog geen bezetting geregistreerd.</p>'}
      ${canWrite ? `<form class="edit" id="fOcc" style="margin-top:8px"><div class="two"><div class="row"><label for="oOrg">Nieuwe bezetting / vertrek</label><select id="oOrg">${[...ORGS].sort((a, b) => (b.name === 'De Overmolen') - (a.name === 'De Overmolen')).map(o => `<option ${o.name === 'De Overmolen' ? 'value="De Overmolen"' : ''}>${o.name === 'De Overmolen' ? 'Vertrek → terug naar De Overmolen' : esc(o.name)}</option>`).join('')}</select></div>
        <div class="row"><label for="oDate">Vanaf datum</label><input type="date" id="oDate" value="${peil}" required></div></div>
        <div class="btns"><button class="btn primary" type="submit">Registreren</button><span id="oMsg"></span></div></form>` : ''}</div>
    <form class="edit" id="fRoom">
      <div class="row"><label for="fName">Ruimte</label><input id="fName" value="${esc(r.name)}" ${dis}></div>
      <div class="two"><div class="row"><label for="fUse">Gebruik</label><select id="fUse" ${dis}>${USES.map(u => `<option ${u === r.use ? 'selected' : ''}>${u}</option>`).join('')}</select></div>
      <div class="row"><label>Organisatie op ${dfmt(peil)}</label><div style="padding:7px 0">${orgOf(r) ? `<span class="swatch" style="background:${orgColor(orgOf(r))}"></span>${esc(orgOf(r))}` : '<span class="empty">– leeg –</span>'}</div></div></div>
      <div class="two"><div class="row"><label for="fM2p">m² plan</label><input id="fM2p" class="mono" inputmode="decimal" value="${r.m2plan ?? ''}" ${dis}></div>
      <div class="row"><label for="fM2m">m² gemeten</label><input id="fM2m" class="mono" inputmode="decimal" value="${r.m2meas ?? ''}" ${dis}></div></div>
      <div class="two"><div class="row"><label for="fSrc">Bron</label><select id="fSrc" ${dis}>${SRCS.map(s => `<option ${s === r.src ? 'selected' : ''}>${s}</option>`).join('')}</select></div>
      <div class="row"><label>&nbsp;</label><label class="check" for="fOk"><input type="checkbox" id="fOk" ${r.ok ? 'checked' : ''} ${dis}> Gecontroleerd (OK)</label></div></div>
      <div class="row"><label for="fNote">Opmerking</label><textarea id="fNote" ${dis}>${esc(r.note)}</textarea></div>
      <div class="btns">${canWrite ? '<button class="btn primary" type="submit">Opslaan</button>' : '<span class="empty" style="font-size:12px">Alleen lezen</span>'}<span id="fMsg"></span></div>
    </form>
    <div><div class="sub">Ista-meters (${ms.length})</div>
      ${ms.length ? `<ul class="mini">${ms.map(m => `<li><span class="mono">${esc(m.nr)}</span><span>${Number(m.units).toLocaleString('nl-BE')} eenh.</span><span class="num">${eur(meterCost(m))}</span></li>`).join('')}
        <li><b>Verwarming ${esc(S.settings.period)}</b><b class="num">${eur(heat)}</b></li></ul>` : '<p class="empty" style="margin:0;font-size:13px">Nog geen meter gekoppeld. Koppel meters in de tab Ista-meters.</p>'}</div>
    <div><div class="sub">Deuren (${doors.length})</div>${doors.length ? `<ul class="mini">${doors.map(d => `<li><span class="mono">${esc(d.code)}</span><span>${esc(d.fr === sel ? '→ ' + d.to : '← ' + d.fr)}</span><span>${d.checked ? '✓' : ''}</span></li>`).join('')}</ul>` : '<p class="empty" style="margin:0;font-size:13px">Geen deuren gekoppeld.</p>'}</div>
    ${r.oldCode ? `<p class="empty" style="margin:0;font-size:12px">Oude code: ${esc(r.oldCode)}${r.updatedAt ? ' · gewijzigd ' + new Date(r.updatedAt).toLocaleString('nl-BE') : ''}</p>` : ''}`;
  const P2 = $('#panel');
  const occTable = P2.querySelector('.occlist');
  if (occTable) occTable.onclick = async e => {
    const item = e.target.closest('.occitem'); if (!item) return;
    const conf = item.querySelector('.occconf'), del = item.querySelector('.occdel');
    if (e.target.closest('.occdel')) { conf.hidden = false; del.hidden = true; return; }
    if (e.target.closest('.occno')) { conf.hidden = true; del.hidden = false; return; }
    if (e.target.closest('.occyes')) {
      const o = OCC.find(x => String(x._id) === item.dataset.id); if (!o) return;
      try { await LK.remove('occ', o._id); OCC = OCC.filter(x => x !== o); await syncRoomOrg(r); setStatus('live', 'Bezetting verwijderd · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' })); renderAll(true); renderPanel(); }
      catch (er) { setStatus('ro', 'Niet verwijderd: ' + (er.status === 403 ? 'geen bewerkrechten' : er.message)); }
    }
  };
  if (occTable) occTable.onchange = async e => {
    const inp = e.target.closest('input[data-k]'); if (!inp) return; const o = OCC.find(x => String(x._id) === inp.closest('.occitem').dataset.id);
    const next = { ...o, [inp.dataset.k]: inp.value };
    if (next.from && next.to && next.to < next.from) { inp.style.outline = '2px solid var(--bad)'; setStatus('ro', "'Tot' ligt vóór 'Van'"); return; }
    try { await save('occ', o._id, next); Object.assign(o, next); syncRoomOrg(r); renderAll(true); } catch (er) { inp.style.outline = '2px solid var(--bad)'; setStatus('ro', 'Niet opgeslagen: ' + (er.code || er.message)); }
  };
  if ($('#fOcc')) $('#fOcc').onsubmit = async e => {
    e.preventDefault(); const msg = $('#oMsg'); const D = $('#oDate').value, org = $('#oOrg').value || 'De Overmolen'; if (!D) return;
    const open = occOf(r.code).filter(o => isOpen(o, D));
    if (!org && !open.length) { msg.className = 'err'; msg.textContent = 'Op die datum zit er niemand.'; return; }
    if (org && open.some(o => o.org === org)) { msg.className = 'err'; msg.textContent = org + ' zit hier al op die datum.'; return; }
    if (open.some(o => o.from && o.from > addDays(D, -1))) { msg.className = 'err'; msg.textContent = 'Een bezetting start op die dag zelf. Pas de datums in de tabel aan.'; return; }
    if (OCC.some(o => o.room === r.code && o.from && o.from > D) && org) { msg.className = 'err'; msg.textContent = 'Er is al een latere bezetting. Pas de datums in de tabel aan.'; return; }
    try {
      msg.className = ''; msg.textContent = 'Opslaan…';
      for (const o of open) { const next = { ...o, to: addDays(D, -1) }; await save('occ', o._id, next); Object.assign(o, next); }
      if (org) { const row = { room: r.code, org, from: D, to: '', note: '' }; await save('occ', null, row); OCC.push(row); }
      await syncRoomOrg(r); msg.className = 'saved'; msg.textContent = 'Geregistreerd'; renderAll(true); renderPanel();
    } catch (er) { msg.className = 'err'; msg.textContent = 'Niet opgeslagen: ' + (er.code === 'permission_denied' ? 'geen bewerkrechten.' : (er.message || er.code)); }
  };
  $('#fRoom').onsubmit = async e => {
    e.preventDefault(); const msg = $('#fMsg');
    const num = v => { const t = String(v).trim().replace(',', '.'); return t === '' ? null : (isNaN(+t) ? NaN : +t); };
    const p = num($('#fM2p').value), mm = num($('#fM2m').value);
    if (Number.isNaN(p) || Number.isNaN(mm)) { msg.className = 'err'; msg.textContent = 'm² moet een getal zijn (bv. 18,5).'; return; }
    const body = { ...r, name: $('#fName').value.trim(), use: $('#fUse').value, m2plan: p, m2meas: mm, src: mm != null ? 'Gemeten' : $('#fSrc').value, ok: $('#fOk').checked, note: $('#fNote').value.trim() };
    try { msg.className = ''; msg.textContent = 'Opslaan…'; await save('rooms', r.code, body); S.rooms[r.code] = body; msg.className = 'saved'; msg.textContent = 'Opgeslagen'; renderAll(true); }
    catch (er) { msg.className = 'err'; msg.textContent = 'Niet opgeslagen: ' + (er.code === 'permission_denied' ? 'geen bewerkrechten.' : (er.message || er.code || 'fout')); }
  };
}
async function syncRoomOrg(r) {
  const now = orgAt(r.code, ymd(new Date()));
  if ((r.org || '') !== now) { const body = { ...r, org: now }; try { await save('rooms', r.code, body); S.rooms[r.code] = body; } catch (e) {} }
}
function floorSummary() {
  const rs = Object.values(S.rooms).filter(r => r.floor === floor);
  const tot = rs.reduce((s, r) => s + (m2(r) || 0), 0), todo = rs.filter(r => m2(r) == null).length;
  const byOrg = {}; rs.forEach(r => { if (orgOf(r)) byOrg[orgOf(r)] = (byOrg[orgOf(r)] || 0) + (m2(r) || 0); });
  return `<div><div class="sub">${esc(floorLabel(floor))}</div><dl class="kv"><dt>Ruimten</dt><dd>${rs.length}</dd><dt>Gekende m²</dt><dd class="mono">${m2f(tot)}</dd><dt>Nog te meten</dt><dd>${todo}</dd></dl></div>
    ${Object.keys(byOrg).length ? `<div><div class="sub">Per organisatie</div><ul class="mini">${Object.entries(byOrg).map(([o, v]) => `<li><span><span class="swatch" style="background:${orgColor(o)}"></span>${esc(o)}</span><span class="mono">${m2f(v)}</span></li>`).join('')}</ul></div>` : ''}`;
}

// ---------- tables
function renderRooms() {
  const f = $('#rFloor').value, o = $('#rOrg').value, q = $('#rQ').value.trim().toLowerCase();
  const rs = Object.values(S.rooms).filter(r => (!f || r.floor === f) && (!o || (o === '-' ? !orgOf(r) : orgOf(r) === o)) && (!q || [r.code, r.name, orgOf(r), r.note].join(' ').toLowerCase().includes(q)))
    .sort((a, b) => GEO.floors.findIndex(x => x.id === a.floor) - GEO.floors.findIndex(x => x.id === b.floor) || a.code.localeCompare(b.code));
  $('#tRooms').innerHTML = `<thead><tr><th>Code</th><th>Ruimte</th><th>Gebruik</th><th>Organisatie</th><th class="num">m² plan</th><th>Bron</th><th class="num">m² gemeten</th><th>OK</th><th class="num">Ista</th><th class="num">Verwarming</th></tr></thead><tbody>` +
    rs.map(r => { const ms = metersOf(r.code); return `<tr class="click" data-code="${esc(r.code)}"><td class="mono">${esc(r.code)}</td><td class="wrap">${esc(r.name)}</td><td><span class="chip ${USEC[r.use] || ''}">${esc(r.use)}</span></td>
      <td>${orgOf(r) ? `<span class="swatch" style="background:${orgColor(orgOf(r))}"></span>${esc(orgOf(r))}` : '<span class="empty">–</span>'}</td><td class="num mono">${r.m2plan ?? ''}</td>
      <td style="color:${r.src === 'Te meten' && r.m2meas == null ? 'var(--bad)' : 'inherit'}">${esc(r.src)}</td><td class="num mono">${r.m2meas ?? ''}</td><td>${r.ok ? '✓' : ''}</td><td class="num">${ms.length || ''}</td><td class="num mono">${ms.length ? eur(ms.reduce((s, m) => s + meterCost(m), 0)) : ''}</td></tr>`; }).join('') + '</tbody>';
  $('#tRooms').onclick = e => { const tr = e.target.closest('tr.click'); if (tr) select(tr.dataset.code); };
}
function renderDoors() {
  const f = $('#dFloor').value, q = $('#dQ').value.trim().toLowerCase(); const dis = canWrite ? '' : 'disabled';
  const ds = Object.values(S.doors).filter(d => (!f || d.floor === f) && (!q || [d.code, d.fr, d.to, d.type, d.note, d.key].join(' ').toLowerCase().includes(q)))
    .sort((a, b) => GEO.floors.findIndex(x => x.id === a.floor) - GEO.floors.findIndex(x => x.id === b.floor) || a.code.localeCompare(b.code));
  $('#tDoors').innerHTML = `<thead><tr><th>Code</th><th>Van</th><th>Naar</th><th>Type</th><th>Maat</th><th>Slot / badge</th><th>Sleutel- / cilindernr</th><th>OK</th><th>Opmerking</th></tr></thead><tbody>` +
    ds.map(d => `<tr data-code="${esc(d.code)}"><td class="mono">${esc(d.code)}</td><td class="mono">${esc(d.fr)}</td><td class="mono">${esc(d.to)}</td><td>${esc(d.type)}</td><td class="mono">${esc(d.size)}</td>
      <td><select data-f="lock" ${dis} aria-label="Slot ${esc(d.code)}">${LOCKS.map(l => `<option ${l === (d.lock || '') ? 'selected' : ''} value="${esc(l)}">${l || '–'}</option>`).join('')}</select></td>
      <td><input data-f="key" ${dis} value="${esc(d.key)}" class="mono" style="width:120px;padding:4px 6px;border:1px solid var(--line);border-radius:4px;background:var(--paper)" aria-label="Sleutel ${esc(d.code)}"></td>
      <td><input type="checkbox" data-f="checked" ${d.checked ? 'checked' : ''} ${dis} aria-label="Gecontroleerd ${esc(d.code)}"></td><td class="wrap">${esc(d.note)}</td></tr>`).join('') + '</tbody>';
  $('#tDoors').onchange = async e => {
    const el = e.target.closest('[data-f]'); if (!el) return; const code = el.closest('tr').dataset.code; const d = S.doors[code];
    const body = { ...d, [el.dataset.f]: el.type === 'checkbox' ? el.checked : el.value.trim() };
    try { await save('doors', code, body); S.doors[code] = body; } catch (er) { el.style.outline = '2px solid var(--bad)'; setStatus('ro', 'Niet opgeslagen: ' + (er.code || er.message)); }
  };
}
function renderIsta() {
  const st = S.settings; $('#istaPeriod').textContent = 'Afrekening ' + st.period;
  const ms = Object.values(S.meters); const linked = ms.filter(m => m.room).length;
  const totUnits = ms.reduce((s, m) => s + (Number(m.units) || 0), 0);
  $('#istaCards').innerHTML = `
    <div class="card"><span class="k">Verwarmingskost (Ista, incl. btw)</span><span class="v">${eur(st.heatCost)}</span><span class="k">${Number(st.totalUnits).toLocaleString('nl-BE')} eenheden · ${price().toLocaleString('nl-BE', { maximumFractionDigits: 4 })} €/eenh.</span></div>
    <div class="card"><span class="k">Huur kostenverdelers</span><span class="v">${eur(st.rentTotal)}</span><span class="k">${st.meterCount} meters · ${eur(rent())} per meter</span></div>
    <div class="card"><span class="k">Meters gekoppeld aan een ruimte</span><span class="v">${linked} / ${ms.length}</span><div class="bar"><span style="width:${ms.length ? linked / ms.length * 100 : 0}%;background:var(--accent)"></span></div></div>`;
  const opts = r => '<option value="">– kies ruimte –</option>' + GEO.floors.map(f => `<optgroup label="${esc(f.label)}">` + Object.values(S.rooms).filter(x => x.floor === f.id).sort((a, b) => a.code.localeCompare(b.code))
    .map(x => `<option value="${esc(x.code)}" ${x.code === r ? 'selected' : ''}>${esc(x.code)} · ${esc(x.name)}</option>`).join('') + '</optgroup>').join('');
  const dis = canWrite ? '' : 'disabled';
  $('#tIsta').innerHTML = `<thead><tr><th>Meternr</th><th>Unit</th><th>Ruimte (Ista)</th><th class="num">Verbruik</th><th>Ruimte in gebouw</th><th>Organisatie</th><th class="num">Verwarming</th><th class="num">Totaal</th><th>Opmerking</th></tr></thead><tbody>` +
    ms.sort((a, b) => a.unit.localeCompare(b.unit) || a.nr.localeCompare(b.nr)).map(m => { const r = roomOf(m.room); return `<tr data-nr="${esc(m.nr)}"><td class="mono">${esc(m.nr)}</td><td class="mono">${esc(m.unit)}</td><td>${esc(m.istaRoom)}</td>
      <td class="num mono">${Number(m.units).toLocaleString('nl-BE')}</td><td><select ${dis} aria-label="Ruimte voor meter ${esc(m.nr)}" style="${m.room ? '' : 'border-color:var(--bad)'}">${opts(m.room)}</select>${r ? ` <button class="btn" data-go="${esc(r.code)}" style="padding:2px 8px;font-size:12px">plan</button>` : ''}</td>
      <td>${orgOf(r) ? `<span class="swatch" style="background:${orgColor(orgOf(r))}"></span>${esc(orgOf(r))}` : (r ? '<span class="empty">–</span>' : '')}</td>
      <td class="num mono">${eur((Number(m.units) || 0) * price())}</td><td class="num mono">${eur(meterCost(m))}</td><td class="wrap">${esc(m.note)}</td></tr>`; }).join('') +
    `</tbody><tfoot><tr><th colspan="3">Totaal</th><th class="num mono">${totUnits.toLocaleString('nl-BE')}</th><th colspan="2"></th><th class="num mono">${eur(totUnits * price())}</th><th class="num mono">${eur(ms.reduce((s, m) => s + meterCost(m), 0))}</th><th></th></tr></tfoot>`;
  $('#tIsta').onchange = async e => { const s = e.target.closest('select'); if (!s) return; const nr = s.closest('tr').dataset.nr; const body = { ...S.meters[nr], room: s.value };
    try { await save('meters', nr, body); S.meters[nr] = body; renderAll(true); } catch (er) { s.style.outline = '2px solid var(--bad)'; setStatus('ro', 'Niet opgeslagen: ' + (er.code || er.message)); } };
  $('#tIsta').onclick = e => { const b = e.target.closest('[data-go]'); if (b) select(b.dataset.go); };
}
function bindOrgColors() {
  $('#tOrgs').onchange = e => { const inp = e.target.closest('input.ocol'); if (!inp) return; const o = ORGS.find(x => x.name === inp.closest('tr').dataset.org); if (o) saveOrg({ ...o, color: inp.value }, null); };
}
function renderOrgs() {
  bindOrgColors();
  const A = per.from, B = per.to, P = dnum(B) - dnum(A) + 1;
  const I = istaPeriod(), IP = dnum(I.to) - dnum(I.from) + 1;
  const rs = Object.values(S.rooms);
  const priv = rs.filter(r => r.use === 'Verhuurbaar'), common = rs.filter(r => r.use === 'Gemeenschappelijk');
  const privTot = priv.reduce((s, r) => s + (m2(r) || 0), 0), comTot = common.reduce((s, r) => s + (m2(r) || 0), 0);
  const unmeasured = rs.filter(r => (r.use === 'Verhuurbaar' || r.use === 'Gemeenschappelijk') && m2(r) == null).length;
  const acc = {}; const A0 = n => acc[n] || (acc[n] = { m2d: 0, heat: 0, rooms: new Set() });
  for (const r of priv) for (const [o, d] of Object.entries(orgDays(r, A, B))) { A0(o).m2d += (m2(r) || 0) * d; A0(o).rooms.add(r.code); }
  for (const m of Object.values(S.meters)) { const r = roomOf(m.room); if (!r) continue; for (const [o, d] of Object.entries(orgDays(r, I.from, I.to))) A0(o).heat += meterCost(m) * d / IP; }
  const rows = ORGS.map(o => { const a = acc[o.name] || { m2d: 0, heat: 0, rooms: new Set() }; const avg = a.m2d / P, share = privTot ? avg / privTot : 0;
    return { o, n: a.rooms.size, now: rs.filter(r => orgOf(r) === o.name).length, avg, share, com: share * comTot, heat: a.heat }; });
  const occAvg = rows.reduce((s, x) => s + x.avg, 0), vacant = Math.max(0, privTot - occAvg);
  $('#orgCards').innerHTML = `<div class="card"><span class="k">Periode</span><span class="v" style="font-size:20px">${dfmt(A)} – ${dfmt(B)}</span><span class="k">${P} dagen · m² gewogen per dag bezetting</span></div>
    <div class="card"><span class="k">Verhuurbare m² (gekend)</span><span class="v">${m2f(privTot)}</span><span class="k">Gemeenschappelijk: ${m2f(comTot)}</span></div>
    <div class="card"><span class="k">Gemiddelde leegstand in periode</span><span class="v">${m2f(vacant)}</span><span class="k">${privTot ? (vacant / privTot * 100).toFixed(1) : 0}% van verhuurbaar</span></div>
    <div class="card"><span class="k">Ruimten zonder m²</span><span class="v">${unmeasured}</span><span class="k">Totalen zijn pas volledig na opmeting.</span></div>`;
  const dis = canWrite ? '' : 'disabled';
  $('#orgNew').innerHTML = canWrite ? `<b>Nieuwe organisatie</b>
    <input id="noName" placeholder="Naam (bv. Brussel Helpt vzw)" aria-label="Naam" style="flex:1 1 220px">
    <label>Kleur <input type="color" id="noColor" value="#c9e4f5"></label>
    <select id="noType" aria-label="Soort"><option>Huurder</option><option>Eigenaar</option><option>Partner</option></select>
    <button class="btn primary" id="noAdd">Toevoegen</button><span id="noMsg"></span>` : '';
  if (canWrite) $('#noAdd').onclick = async () => {
    const name = $('#noName').value.trim(), msg = $('#noMsg');
    if (!name) { msg.className = 'err'; msg.textContent = 'Geef een naam in.'; return; }
    if (ORGS.some(o => o.name.toLowerCase() === name.toLowerCase())) { msg.className = 'err'; msg.textContent = 'Deze organisatie bestaat al.'; return; }
    msg.className = ''; msg.textContent = 'Opslaan…';
    await saveOrg({ name, color: $('#noColor').value, type: $('#noType').value }, msg);
  };
  $('#tOrgs').innerHTML = `<thead><tr><th>Organisatie</th><th class="num">Ruimten in periode</th><th class="num">Nu (${dfmt(peil)})</th><th class="num">Gem. privatief m²</th><th>Aandeel</th><th class="num">Aandeel gemeensch. m²</th><th class="num">Totaal m²</th><th class="num">Verwarming Ista ${esc(S.settings.period || '')}</th></tr></thead><tbody>` +
    rows.map(x => `<tr data-org="${esc(x.o.name)}"><td><input type="color" class="ocol" value="${esc(x.o.color)}" ${dis} title="Kleur wijzigen" aria-label="Kleur ${esc(x.o.name)}"> ${esc(x.o.name)} <span class="empty" style="font-size:11px">${esc(x.o.type || '')}</span></td><td class="num">${x.n}</td><td class="num">${x.now}</td><td class="num mono">${m2f(x.avg)}</td>
      <td style="min-width:140px"><div class="bar"><span style="width:${(x.share * 100).toFixed(1)}%;background:${x.o.color}"></span></div><span class="mono" style="font-size:12px">${(x.share * 100).toFixed(2)}%</span></td>
      <td class="num mono">${m2f(x.com)}</td><td class="num mono">${m2f(x.avg + x.com)}</td><td class="num mono">${x.heat ? eur(x.heat) : '–'}</td></tr>`).join('') +
    `</tbody><tfoot><tr><th>Leegstand</th><th></th><th></th><th class="num mono">${m2f(vacant)}</th><th class="mono" style="font-size:12px">${privTot ? (vacant / privTot * 100).toFixed(2) : 0}%</th><th></th><th></th><th></th></tr></tfoot>`;
}
// ---------- bezetting per periode (tijdlijn + tabel)
// jaarkeuze (kalenderjaar) in Organisaties, Bezetting en Afrekening
function yearSel(id) { const el = $(id); if (!el) return; const y0 = +peil.slice(0, 4); el.innerHTML = '<option value="">anders</option>' + [y0 + 1, y0, y0 - 1, y0 - 2, y0 - 3].map(y => `<option>${y}</option>`).join(''); el.onchange = () => { if (el.value) setPeriod(el.value + '-01-01', el.value + '-12-31'); }; syncYears(); }
function syncYears() { const y = per && per.from.slice(5) === '01-01' && per.to.slice(5) === '12-31' && per.from.slice(0, 4) === per.to.slice(0, 4) ? per.from.slice(0, 4) : ''; ['#pYear', '#bYear', '#aYear'].forEach(i => { if ($(i)) $(i).value = y; }); }
function setPeriod(a, b) {
  if (!a || !b || a > b) return; per = { from: a, to: b };
  for (const [f, t] of [['#pFrom', '#pTo'], ['#bFrom', '#bTo'], ['#aFrom', '#aTo']]) { if ($(f)) { $(f).value = a; $(t).value = b; } }
  syncYears();
  renderOrgs(); renderOcc(); renderAfr();
}
function bindOcc() {
  if (bindOcc.done) return; bindOcc.done = true;
  $('#bFrom').onchange = $('#bTo').onchange = () => setPeriod($('#bFrom').value, $('#bTo').value);
  yearSel('#bYear');
  ['#bOrg', '#bFloor'].forEach(x => $(x).oninput = renderOcc);
  $('#bExp').onclick = () => { const rows = occRows(); exportCsv('Lakenhuis_bezetting_' + per.from + '_' + per.to + '.csv', ['ruimtecode', 'ruimte', 'verdieping', 'organisatie', 'van', 'tot', 'dagen_in_periode', 'm2', 'm2_dagen', 'gemiddelde_m2_periode'],
    rows.map(x => [x.r.code, x.r.name, floorLabel(x.r.floor), x.o.org, dfmt(x.o.from) || 'onbekend', dfmt(x.o.to) || 'lopend', x.days, x.m2 ?? '', (x.m2 * x.days || 0).toFixed(1), x.m2 != null ? (x.m2 * x.days / x.P).toFixed(2) : ''])); };
}
function occRows() {
  const A = per.from, B = per.to, P = dnum(B) - dnum(A) + 1, fo = $('#bOrg').value, ff = $('#bFloor').value, out = [];
  for (const o of OCC) {
    const r = roomOf(o.room); if (!r) continue;
    if (fo && o.org !== fo) continue; if (ff && r.floor !== ff) continue;
    const s0 = Math.max(dnum(A), o.from ? dnum(o.from) : -1e9), e0 = Math.min(dnum(B), o.to ? dnum(o.to) : 1e9);
    if (e0 < s0) continue;
    out.push({ o, r, s0, e0, days: e0 - s0 + 1, P, m2: m2(r) });
  }
  return out.sort((a, b) => GEO.floors.findIndex(f => f.id === a.r.floor) - GEO.floors.findIndex(f => f.id === b.r.floor) || a.r.code.localeCompare(b.r.code) || a.s0 - b.s0);
}
function renderOcc() {
  if (!per) return; bindOcc();
  const cur = [$('#bOrg').value, $('#bFloor').value];
  $('#bOrg').innerHTML = '<option value="">Alle organisaties</option>' + ORGS.map(o => `<option>${esc(o.name)}</option>`).join(''); $('#bOrg').value = cur[0];
  $('#bFloor').innerHTML = '<option value="">Alle verdiepingen</option>' + GEO.floors.map(f => `<option value="${f.id}">${esc(f.label)}</option>`).join(''); $('#bFloor').value = cur[1];
  $('#bFrom').value = per.from; $('#bTo').value = per.to;
  const rows = occRows(), A = dnum(per.from), P = dnum(per.to) - A + 1;
  // summary per organisation
  const sum = {}; rows.forEach(x => { const k = x.o.org; sum[k] = sum[k] || { days: 0, m2d: 0, rooms: new Set() }; sum[k].m2d += (x.m2 || 0) * x.days; sum[k].rooms.add(x.r.code); });
  $('#occCards').innerHTML = `<div class="card"><span class="k">Periode</span><span class="v" style="font-size:20px">${dfmt(per.from)} – ${dfmt(per.to)}</span><span class="k">${P} dagen · ${rows.length} bezettingen</span></div>` +
    Object.entries(sum).sort((a, b) => b[1].m2d - a[1].m2d).map(([k, v]) => `<div class="card"><span class="k"><span class="swatch" style="background:${orgColor(k) || '#ddd'}"></span>${esc(k)}</span><span class="v">${m2f(v.m2d / P)}</span><span class="k">gem. m² in periode · ${v.rooms.size} ruimte(n)</span></div>`).join('');
  // timeline
  const months = []; { let d = new Date(per.from + 'T00:00:00Z'); d.setUTCDate(1); while (d.toISOString().slice(0, 10) <= per.to) { const iso = d.toISOString().slice(0, 10); const x = Math.max(0, (dnum(iso) - A) / P * 100); months.push(`<span style="left:${x}%">${['jan','feb','mrt','apr','mei','jun','jul','aug','sep','okt','nov','dec'][d.getUTCMonth()]}${d.getUTCMonth() === 0 ? ' ' + d.getUTCFullYear() : ''}</span>`); d.setUTCMonth(d.getUTCMonth() + 1); } }
  const byRoom = {}; rows.forEach(x => (byRoom[x.r.code] = byRoom[x.r.code] || []).push(x));
  const pp = dnum(peil) >= A && dnum(peil) <= A + P - 1 ? `<i class="peilline" style="left:${(dnum(peil) - A + 0.5) / P * 100}%" title="Peildatum ${dfmt(peil)}"></i>` : '';
  $('#occTimeline').innerHTML = rows.length ? `<div class="tl-head"><div class="tl-l">Ruimte</div><div class="tl-r">${months.join('')}${pp}</div></div>` + Object.entries(byRoom).map(([code, xs]) => {
    const r = roomOf(code);
    return `<div class="tl-row" data-code="${esc(code)}"><div class="tl-l"><span class="mono">${esc(short(code))}</span> <span class="empty">${esc(r.name)}</span></div><div class="tl-r">${pp}` + xs.map(x => {
      const l = (x.s0 - A) / P * 100, w = x.days / P * 100;
      return `<b class="bar" style="left:${l}%;width:${w}%;background:${orgColor(x.o.org) || '#ddd'}" title="${esc(x.o.org)}: ${dfmt(x.o.from) || 'onbekend'} – ${dfmt(x.o.to) || 'lopend'} (${x.days} d in periode)">${w > 12 ? esc(x.o.org) : ''}</b>`;
    }).join('') + '</div></div>';
  }).join('') : '<p class="empty">Geen bezetting in deze periode.</p>';
  $('#occTimeline').onclick = e => { const row = e.target.closest('.tl-row'); if (row) select(row.dataset.code); };
  // table
  $('#tOcc').innerHTML = `<thead><tr><th>Code</th><th>Ruimte</th><th>Verdieping</th><th>Organisatie</th><th>Van</th><th>Tot</th><th class="num">Dagen in periode</th><th class="num">m²</th><th class="num">Gem. m² periode</th></tr></thead><tbody>` +
    rows.map(x => `<tr class="click" data-code="${esc(x.r.code)}"><td class="mono">${esc(x.r.code)}</td><td class="wrap">${esc(x.r.name)}</td><td>${esc(floorLabel(x.r.floor))}</td>
      <td><span class="swatch" style="background:${orgColor(x.o.org) || '#ddd'}"></span>${esc(x.o.org)}</td><td class="mono">${dfmt(x.o.from) || '<span class="empty">onbekend</span>'}</td><td class="mono">${dfmt(x.o.to) || '<span class="empty">lopend</span>'}</td>
      <td class="num mono">${x.days}</td><td class="num mono">${x.m2 != null ? m2f(x.m2) : '–'}</td><td class="num mono">${x.m2 != null ? m2f(x.m2 * x.days / P) : '–'}</td></tr>`).join('') + '</tbody>';
  $('#tOcc').onclick = e => { const tr = e.target.closest('tr.click'); if (tr) select(tr.dataset.code); };
}
// ---------- afrekening: kostenposten + verdeelsleutels
const CATS = ['Water (algemene meter)', 'Elektriciteit algemeen', 'Elektriciteit eigen meter', 'Gas / verwarming', 'Ista – opname/beheer', 'Schoonmaak', 'Herstelling / onderhoud', 'Lift', 'Keuringen / brandveiligheid', 'Verzekering', 'Afval', 'Telefoon / internet', 'Huur / huurvermindering', 'Andere'];
const KEYS = ['m² × dagen bezetting', 'Ista-verbruik', 'Gelijke delen (huurders)', 'Rechtstreeks aan één organisatie', 'Eigen m² + gemeensch. naar aanwezigen (m²)', 'Eigen m² + gemeensch. naar aanwezigen (gelijk)', 'Bezette m² (leegstand niet aangerekend)'];
const KDEF = KEYS[6];   // standaardsleutel voor huurlasten
// gemeenschappelijk deel: per dag verdeeld over wie die dag een verhuurbare ruimte gebruikt (naar m² of gelijk); niemand aanwezig → De Overmolen
let PSC = {}, SIMKEY = '';
const keyOf = c => SIMKEY && c.key !== KEYS[3] ? SIMKEY : (c.key || KDEF);
function orgOnDay(r, day) { const rows = occOf(r.code); if (!rows.length) return r.org || ''; const o = rows.find(o => isOpen(o, day)); return o ? o.org : ''; }
function presentShares(A, B, equal, noOvm) {
  const ck = A + B + equal + !!noOvm; if (PSC[ck]) return PSC[ck];
  const out = {}, priv = Object.values(S.rooms).filter(r => r.use === 'Verhuurbaar');
  for (let d = dnum(A); d <= dnum(B); d++) {
    const day = new Date(d * 864e5).toISOString().slice(0, 10), w = {};
    for (const r of priv) { const o = orgOnDay(r, day); if (o && !(noOvm && o === 'De Overmolen')) w[o] = equal ? 1 : (w[o] || 0) + (m2(r) || 0); }
    const T = Object.values(w).reduce((a, b) => a + b, 0);
    if (T > 0) for (const [o, v] of Object.entries(w)) out[o] = (out[o] || 0) + v / T; else out['De Overmolen'] = (out['De Overmolen'] || 0) + 1;
  }
  return (PSC[ck] = out);   // som = aantal dagen
}
const eur2 = n => (Math.round(n * 100) / 100).toLocaleString('nl-BE', { style: 'currency', currency: 'EUR' });
function weights(key, A, B, org) {
  const P = dnum(B) - dnum(A) + 1, w = {};
  if (key === KEYS[3]) { if (org) w[org] = 1; return w; }
  if (key === KEYS[1]) {   // meter in ruimte van een organisatie → die organisatie; leeg → De Overmolen; gemeenschappelijke ruimte → aanwezigen
    for (const m of Object.values(S.meters)) {
      const r = roomOf(m.room), u = Number(m.units) || 0; if (!r || !u) continue;
      if (r.use === 'Verhuurbaar') { let used = 0; for (const [o, d] of Object.entries(orgDays(r, A, B))) { w[o] = (w[o] || 0) + u * d / P; used += d; } if (used < P) w['De Overmolen'] = (w['De Overmolen'] || 0) + u * (P - used) / P; }
      else for (const [o, s] of Object.entries(presentShares(A, B, false))) w[o] = (w[o] || 0) + u * s / P;
    }
    if (Object.values(w).reduce((a, b) => a + b, 0) > 0) return w;
  }
  if (key === KEYS[6] || key === KEYS[1]) {   // bezette m²: alles over wie er die dag is (naar m²); ruimten op naam van De Overmolen / leegstand betalen niets. Ook terugval voor Ista zonder gekoppelde meters
    for (const [o, s] of Object.entries(presentShares(A, B, false, true))) w[o] = (w[o] || 0) + s;
    return w;
  }
  if (key === KEYS[4] || key === KEYS[5]) {   // eigen m² (leegstand → De Overmolen) + gemeenschappelijke m² naar aanwezigen
    const priv = Object.values(S.rooms).filter(r => r.use === 'Verhuurbaar'), comm = Object.values(S.rooms).filter(r => r.use === 'Gemeenschappelijk');
    const Vm = priv.reduce((a, r) => a + (m2(r) || 0), 0), Gm = comm.reduce((a, r) => a + (m2(r) || 0), 0), pf = Vm + Gm ? Vm / (Vm + Gm) : 1;
    let used = 0; for (const r of priv) for (const [o, d] of Object.entries(orgDays(r, A, B))) { const v = (m2(r) || 0) * d; w[o] = (w[o] || 0) + pf * v / (Vm * P || 1); used += v; }
    if (Vm * P > used) w['De Overmolen'] = (w['De Overmolen'] || 0) + pf * (Vm * P - used) / (Vm * P);
    for (const [o, s] of Object.entries(presentShares(A, B, key === KEYS[5]))) w[o] = (w[o] || 0) + (1 - pf) * s / P;
    return w;
  }
  if (key === KEYS[2]) {
    const huur = new Set(ORGS.filter(o => o.type !== 'Eigenaar').map(o => o.name));
    for (const r of Object.values(S.rooms)) for (const [o, d] of Object.entries(orgDays(r, A, B))) if (huur.has(o)) w[o] = Math.max(w[o] || 0, d / P);
    return w;
  }
  // m² × dagen (verhuurbare ruimten); niet-bezette m²-dagen gaan naar De Overmolen
  let tot = 0;
  for (const r of Object.values(S.rooms)) { if (r.use !== 'Verhuurbaar') continue; const a = m2(r) || 0; tot += a * P; for (const [o, d] of Object.entries(orgDays(r, A, B))) w[o] = (w[o] || 0) + a * d; }
  const used = Object.values(w).reduce((a, b) => a + b, 0); if (tot > used) w['De Overmolen'] = (w['De Overmolen'] || 0) + (tot - used);
  return w;
}
// deel van een kost dat in de gekozen periode valt (pro rata dagen)
function ovl(c) {
  const A = c.from && c.from > per.from ? c.from : per.from, B = c.to && c.to < per.to ? c.to : per.to;
  const d = Math.max(0, dnum(B) - dnum(A) + 1), tot = c.from && c.to ? dnum(c.to) - dnum(c.from) + 1 : d;
  return { A, B, f: tot > 0 ? Math.min(1, d / tot) : 1 };
}
const amtIn = c => (Number(c.amount) || 0) * ovl(c).f;
function split(c) {
  const { A, B, f } = ovl(c), w = weights(keyOf(c), A, B, c.org), T = Object.values(w).reduce((a, b) => a + b, 0), amt = (Number(c.amount) || 0) * f, out = {};
  if (T > 0) for (const [o, v] of Object.entries(w)) out[o] = amt * v / T;
  return out;
}
function costsInPeriod() { return COSTS.filter(c => (!c.to || c.to >= per.from) && (!c.from || c.from <= per.to)).sort((a, b) => (a.cat || '').localeCompare(b.cat || '') || (a.from || '').localeCompare(b.from || '')); }
function bindAfr() {
  if (bindAfr.done) return; bindAfr.done = true;
  $('#aFrom').onchange = $('#aTo').onchange = () => setPeriod($('#aFrom').value, $('#aTo').value);
  $('#aSim').innerHTML = '<option value="">Sleutel per kostenlijn</option>' + KEYS.filter(k => k !== KEYS[3]).map(k => `<option>${esc(k)}</option>`).join('');
  $('#aSim').onchange = () => { SIMKEY = $('#aSim').value; PSC = {}; renderAfr(); };
  yearSel('#aYear');
  $('#aPrint').onclick = () => { document.body.classList.add('print-afr'); window.print(); setTimeout(() => document.body.classList.remove('print-afr'), 500); };
  $('#aPdf').onclick = () => orgPdf($('#aPdfOrg').value);
  $('#aExp').onclick = () => { PSC = {}; const cs = costsInPeriod(), orgs = afrOrgs(cs);
    exportCsv('Lakenhuis_afrekening_' + per.from + '_' + per.to + '.csv', ['categorie', 'omschrijving', 'leverancier', 'factuur', 'meter', 'van', 'tot', 'bedrag', 'sleutel', ...orgs],
      cs.map(c => { const sp = split(c); return [c.cat, c.desc, c.supplier, c.invoice || c.entry, c.meter, dfmt(c.from), dfmt(c.to), amtIn(c).toFixed(2), c.key + (c.org ? ' (' + c.org + ')' : ''), ...orgs.map(o => (sp[o] || 0).toFixed(2))]; })); };
}
// afrekening per organisatie: één pagina per organisatie, via afdrukken → "Opslaan als PDF"
function orgPdf(sel) {
  PSC = {}; const cs = costsInPeriod(), vo = {};
  VOS.filter(v => v.date && v.date >= per.from && v.date <= per.to).forEach(v => { vo[v.org] = (vo[v.org] || 0) + (Number(v.amount) || 0); });
  const all = afrOrgs(cs), list = sel ? [sel] : [...all, ...Object.keys(vo).filter(o => !all.includes(o))];
  if (!list.length) { setStatus('ro', 'Geen organisaties met kosten in deze periode'); return; }
  const today = dfmt(new Date().toISOString().slice(0, 10));
  $('#afrPdf').innerHTML = list.map(o => {
    const rows = cs.map(c => ({ c, sh: split(c)[o] || 0 })).filter(x => Math.abs(x.sh) >= 0.005), sh = rows.reduce((a, x) => a + x.sh, 0), v = vo[o] || 0, sal = sh - v, cat = {};
    rows.forEach(x => { const k = x.c.cat || 'Overige'; cat[k] = (cat[k] || 0) + x.sh; });
    return `<section class="orgpdf"><div class="ph"><div>Lakenhuis · VZW De Overmolen<h1>Afrekening gebouwkosten</h1>Periode ${dfmt(per.from)} – ${dfmt(per.to)}</div><div style="text-align:right">Organisatie<h1>${esc(o)}</h1>Opgemaakt op ${today}</div></div>
      ${SIMKEY ? `<p><b>Simulatie</b> – alle kostenlijnen (behalve 'Rechtstreeks') berekend met "${esc(SIMKEY)}". Geen definitieve afrekening.</p>` : ''}
      <h2>Overzicht</h2><table class="sum"><tbody><tr><td>Aandeel in de kosten</td><td class="num">${eur2(sh)}</td></tr><tr><td>Betaalde voorschotten</td><td class="num">${v ? eur2(v) : '–'}</td></tr></tbody>
      <tfoot><tr><th>${sal > 0.005 ? 'Te betalen' : sal < -0.005 ? 'Terug te krijgen' : 'Saldo'}</th><th class="num">${eur2(Math.abs(sal))}</th></tr></tfoot></table>
      ${rows.length ? `<h2>Per categorie</h2><table class="sum"><tbody>${Object.entries(cat).map(([k, n]) => `<tr><td>${esc(k)}</td><td class="num">${eur2(n)}</td></tr>`).join('')}</tbody></table>
      <h2>Detail kostenposten</h2><table><thead><tr><th>Categorie</th><th>Omschrijving</th><th>Periode</th><th class="num">Bedrag</th><th>Verdeelsleutel</th><th class="num">Uw aandeel</th></tr></thead><tbody>` +
        rows.map(({ c, sh }) => `<tr><td>${esc(c.cat)}</td><td>${esc(c.desc)}${c.supplier || c.invoice ? `<br><span class="sml">${esc([c.supplier, c.invoice && 'factuur ' + c.invoice].filter(Boolean).join(' · '))}</span>` : ''}</td><td style="white-space:nowrap">${dfmt(ovl(c).A)} – ${dfmt(ovl(c).B)}</td><td class="num">${eur2(amtIn(c))}</td><td>${esc(SIMKEY && c.key !== KEYS[3] ? SIMKEY : c.key || KDEF)}</td><td class="num">${eur2(sh)}</td></tr>`).join('') +
        `</tbody><tfoot><tr><th colspan="5">Totaal aandeel</th><th class="num">${eur2(sh)}</th></tr></tfoot></table>` : '<p>Geen kosten aangerekend in deze periode.</p>'}
      <p class="pf">Bedragen = deel van elke kostenpost dat in de periode valt, verdeeld volgens de vermelde sleutel. Voorschotten = betalingen met boekdatum in de periode.</p></section>`;
  }).join('');
  const t = document.title; document.title = 'Lakenhuis_afrekening_' + (sel || 'alle_organisaties').replace(/[^\w-]+/g, '_') + '_' + per.from + '_' + per.to;
  document.body.classList.add('print-org'); window.print();
  setTimeout(() => { document.body.classList.remove('print-org'); document.title = t; }, 500);
}
function afrOrgs(cs) { const set = new Set(); cs.forEach(c => Object.keys(split(c)).forEach(o => set.add(o))); return [...ORGS.map(o => o.name).filter(n => set.has(n)), ...[...set].filter(n => !ORGS.some(o => o.name === n))]; }
function renderAfr() {
  if (!per || !$('#v-afr')) return; bindAfr(); bindImport(); PSC = {};
  const li = IMPS.slice().sort((a, b) => (b._id | 0) - (a._id | 0))[0];
  $('#afrLog').innerHTML = li ? `Laatste import: <b>${esc(li.file)}</b> op ${esc(li.at)} (boekdatum ${dfmt(li.from)} – ${dfmt(li.to)}, ${li.added || 0} nieuw, ${li.changed || 0} bijgewerkt, ${li.removed || 0} verwijderd) · ${IMPS.length} import(s) in totaal` : 'Nog geen import uit Exact.';
  $('#aFrom').value = per.from; $('#aTo').value = per.to;
  const dis = canWrite;
  $('#afrNew').innerHTML = dis ? `<b>Nieuwe kostenpost</b>
    <select id="cCat" aria-label="Categorie">${CATS.map(c => `<option>${c}</option>`).join('')}</select>
    <input id="cDesc" placeholder="Omschrijving (bv. Sibelga factuur Q1)" style="flex:2 1 200px" aria-label="Omschrijving">
    <input id="cAmt" placeholder="Bedrag €" inputmode="decimal" style="width:110px" aria-label="Bedrag">
    <label>Van <input type="date" id="cFrom" value="${per.from}"></label><label>Tot <input type="date" id="cTo" value="${per.to}"></label>
    <select id="cKey" aria-label="Verdeelsleutel">${KEYS.map(k => `<option${k === KDEF ? ' selected' : ''}>${k}</option>`).join('')}</select>
    <select id="cOrg" aria-label="Organisatie" hidden>${ORGS.map(o => `<option>${esc(o.name)}</option>`).join('')}</select>
    <input id="cSup" placeholder="Leverancier" style="width:140px" aria-label="Leverancier"><input id="cInv" placeholder="Factuurnr" style="width:110px" aria-label="Factuurnummer"><input id="cMeter" placeholder="Meter / EAN" style="width:120px" aria-label="Meter">
    <button class="btn primary" id="cAdd">Toevoegen</button><span id="cMsg"></span>` : '';
  if (dis) {
    const sync = () => { $('#cOrg').hidden = $('#cKey').value !== KEYS[3]; };
    $('#cCat').onchange = () => { const v = $('#cCat').value; $('#cKey').value = v === 'Gas / verwarming' ? KEYS[1] : v === 'Elektriciteit eigen meter' ? KEYS[3] : KDEF; sync(); };
    $('#cKey').onchange = sync;
    $('#cAdd').onclick = async () => {
      const msg = $('#cMsg'), amt = Number(String($('#cAmt').value).replace(/\s/g, '').replace(',', '.'));
      if (!$('#cDesc').value.trim() || !isFinite(amt) || !$('#cAmt').value) { msg.className = 'err'; msg.textContent = 'Geef omschrijving en bedrag in.'; return; }
      if ($('#cFrom').value > $('#cTo').value) { msg.className = 'err'; msg.textContent = "'Tot' ligt vóór 'Van'."; return; }
      const c = { desc: $('#cDesc').value.trim(), cat: $('#cCat').value, amount: amt, from: $('#cFrom').value, to: $('#cTo').value, key: $('#cKey').value, org: $('#cKey').value === KEYS[3] ? $('#cOrg').value : '', supplier: $('#cSup').value.trim(), invoice: $('#cInv').value.trim(), meter: $('#cMeter').value.trim(), note: '' };
      try { msg.className = ''; msg.textContent = 'Opslaan…'; await LK.save('costs', c); COSTS.push(c); renderAfr(); setStatus('live', 'Kostenpost opgeslagen · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' })); }
      catch (e) { msg.className = 'err'; msg.textContent = 'Niet opgeslagen: ' + (e.status === 403 ? 'geen bewerkrechten.' : e.message); }
    };
  }
  const cs = costsInPeriod(), orgs = afrOrgs(cs), tot = {}; let total = 0;
  cs.forEach(c => { const sp = split(c); total += amtIn(c); for (const [o, v] of Object.entries(sp)) tot[o] = (tot[o] || 0) + v; });
  const vo = {}; VOS.filter(v => v.date && v.date >= per.from && v.date <= per.to).forEach(v => { vo[v.org] = (vo[v.org] || 0) + (Number(v.amount) || 0); });
  const sorgs = [...orgs, ...Object.keys(vo).filter(o => !orgs.includes(o))];
  const saldo = o => (tot[o] || 0) - (vo[o] || 0);
  const po = $('#aPdfOrg').value; $('#aPdfOrg').innerHTML = '<option value="">Alle organisaties</option>' + sorgs.map(o => `<option>${esc(o)}</option>`).join(''); $('#aPdfOrg').value = sorgs.includes(po) ? po : '';
  $('#afrSimNote').hidden = !SIMKEY; $('#afrSimNote').textContent = SIMKEY ? `Simulatie: alle kostenlijnen (behalve 'Rechtstreeks') berekend met "${SIMKEY}". Niets wordt opgeslagen – zet terug op 'Sleutel per kostenlijn' voor de echte afrekening.` : '';
  $('#afrCards').innerHTML = `<div class="card"><span class="k">Periode</span><span class="v" style="font-size:20px">${dfmt(per.from)} – ${dfmt(per.to)}</span><span class="k">${cs.length} kostenposten</span></div>
    <div class="card"><span class="k">Totale kosten</span><span class="v">${eur2(total)}</span></div>` +
    sorgs.map(o => `<div class="card"><span class="k"><span class="swatch" style="background:${orgColor(o) || '#ddd'}"></span>${esc(o)}</span><span class="v">${eur2(tot[o] || 0)}</span><span class="k">${total ? ((tot[o] || 0) / total * 100).toFixed(1) : 0}% van totaal${vo[o] ? ` · voorschot ${eur2(vo[o])}<br><b class="${saldo(o) > 0 ? 'err' : 'saved'}">${saldo(o) > 0 ? 'te betalen ' + eur2(saldo(o)) : 'terug ' + eur2(-saldo(o))}</b>` : ''}</span></div>`).join('');
  $('#tSaldo').innerHTML = sorgs.length ? `<thead><tr><th>Organisatie</th><th class="num">Aandeel kosten</th><th class="num">Betaalde voorschotten</th><th class="num">Saldo</th><th></th></tr></thead><tbody>` +
    sorgs.map(o => `<tr><td><span class="swatch" style="background:${orgColor(o) || '#ddd'}"></span>${esc(o)}</td><td class="num mono">${eur2(tot[o] || 0)}</td><td class="num mono">${vo[o] ? eur2(vo[o]) : '–'}</td><td class="num mono"><b>${eur2(saldo(o))}</b></td><td>${o === 'De Overmolen' ? '<span class="empty">eigen deel / leegstand</span>' : saldo(o) > 0.005 ? '<span class="err">te betalen</span>' : saldo(o) < -0.005 ? '<span class="saved">terug te krijgen</span>' : ''}</td></tr>`).join('') +
    `</tbody><tfoot><tr><th>Totaal</th><th class="num mono">${eur2(Object.values(tot).reduce((a, b) => a + b, 0))}</th><th class="num mono">${eur2(Object.values(vo).reduce((a, b) => a + b, 0))}</th><th class="num mono">${eur2(sorgs.reduce((a, o) => a + saldo(o), 0))}</th><th></th></tr></tfoot>` : '';
  $('#tAfr').innerHTML = cs.length ? `<thead><tr><th>Categorie</th><th>Omschrijving</th><th>Periode</th><th class="num">Bedrag</th><th>Sleutel</th>${orgs.map(o => `<th class="num"><span class="swatch" style="background:${orgColor(o) || '#ddd'}"></span>${esc(o)}</th>`).join('')}<th></th></tr></thead><tbody>` +
    cs.map(c => { const sp = split(c); return `<tr data-id="${esc(c._id)}"><td>${canWrite ? `<select class="ce" data-f="cat">${[...new Set([...CATS, c.cat || ''])].filter(Boolean).map(x => `<option${x === c.cat ? ' selected' : ''}>${esc(x)}</option>`).join('')}</select>` : esc(c.cat)}</td><td class="wrap">${esc(c.desc)}${c.supplier || c.invoice || c.meter ? `<br><span class="empty" style="font-size:11px">${esc([c.supplier, c.invoice && 'factuur ' + c.invoice, c.entry && 'boekstuk ' + c.entry, c.gl, c.meter && 'meter ' + c.meter].filter(Boolean).join(' · '))}</span>` : ''}</td>
      <td class="mono" style="font-size:12px;white-space:nowrap">${canWrite ? `<input type="date" class="ce" data-f="from" value="${c.from || ''}"><br><input type="date" class="ce" data-f="to" value="${c.to || ''}">` : dfmt(c.from) + '<br>' + dfmt(c.to)}</td><td class="num mono">${eur2(amtIn(c))}${ovl(c).f < 0.9999 ? `<br><span class="empty" style="font-size:10px">${Math.round(ovl(c).f * 100)}% van ${eur2(Number(c.amount) || 0)}</span>` : ''}</td><td style="font-size:12px">${canWrite ? `<select class="ce" data-f="key">${KEYS.map(k => `<option${k === (c.key || KDEF) ? ' selected' : ''}>${esc(k)}</option>`).join('')}</select>${c.key === KEYS[3] ? `<br><select class="ce" data-f="org">${ORGS.map(o => `<option${o.name === c.org ? ' selected' : ''}>${esc(o.name)}</option>`).join('')}</select>` : ''}` : esc(c.key) + (c.org ? '<br><b>' + esc(c.org) + '</b>' : '')}${SIMKEY && c.key !== KEYS[3] ? '<br><span class="imp-warn">simulatie: ' + esc(SIMKEY) + '</span>' : ''}</td>
      ${orgs.map(o => `<td class="num mono">${sp[o] ? eur2(sp[o]) : '–'}</td>`).join('')}<td>${canWrite ? '<button class="occdel cdel" title="Verwijderen">✕</button>' : ''}</td></tr>`; }).join('') +
    `</tbody><tfoot><tr><th colspan="3">Totaal</th><th class="num mono">${eur2(total)}</th><th></th>${orgs.map(o => `<th class="num mono">${eur2(tot[o] || 0)}</th>`).join('')}<th></th></tr></tfoot>` : '<tbody><tr><td class="empty">Nog geen kostenposten in deze periode. Voeg ze hierboven toe.</td></tr></tbody>';
  $('#tAfr').onchange = async e => {
    const el = e.target.closest('.ce'); if (!el) return; const tr = el.closest('tr'), c = COSTS.find(x => String(x._id) === tr.dataset.id); if (!c) return;
    const f = el.dataset.f, patch = { _id: c._id, [f]: el.value };
    if (f === 'key') { patch.org = el.value === KEYS[3] ? (c.org || 'De Overmolen') : ''; }
    if (f === 'cat') { /* enkel categorie */ }
    if ((f === 'from' && el.value > (c.to || '9')) || (f === 'to' && el.value < (c.from || '0'))) { setStatus('ro', "'Tot' ligt vóór 'Van' – niet opgeslagen"); renderAfr(); return; }
    try { await LK.save('costs', patch); Object.assign(c, patch); PSC = {}; renderAfr(); setStatus('live', 'Kostenpost aangepast · herberekend · ' + new Date().toLocaleTimeString('nl-BE', { hour: '2-digit', minute: '2-digit' })); }
    catch (er) { setStatus('ro', 'Niet opgeslagen: ' + (er.status === 403 ? 'geen bewerkrechten' : er.message)); renderAfr(); }
  };
  $('#tAfr').onclick = async e => {
    const b = e.target.closest('.cdel'); if (!b) return; const tr = b.closest('tr');
    if (!tr.dataset.confirm) { tr.dataset.confirm = 1; b.textContent = 'Zeker?'; b.style.width = 'auto'; return; }
    const c = COSTS.find(x => String(x._id) === tr.dataset.id); if (!c) return;
    try { await LK.remove('costs', c._id); COSTS = COSTS.filter(x => x !== c); renderAfr(); } catch (er) { setStatus('ro', 'Niet verwijderd: ' + er.message); }
  };
}
// ---------- import Exact Online (export "FinTransactionSearch": CSV/TXT, tab of ; gescheiden, UTF-16 of UTF-8)
const ORG_ALIAS = [[/accolage/i, 'Accolage'], [/huisartsenkring|\bbhak\b/i, 'BHAK'], [/ondersteuningsplan|^d\.?o\.?p\.?$/i, 'DOP'], [/huis voor gezondheid/i, 'Huis voor Gezondheid'], [/kenniscentrum|\bwwz\b/i, 'Kenniscentrum WWZ'], [/overmolen/i, 'De Overmolen']];
let ORGMAP = {}; try { ORGMAP = JSON.parse(localStorage.getItem('lk.orgmap') || '{}'); } catch (e) {}
function mapOrg(n) {
  if (!n) return ''; if (ORGMAP[n]) return ORGMAP[n];
  const exact = ORGS.find(x => x.name.toLowerCase() === n.toLowerCase()); if (exact) return exact.name;
  for (const [re, v] of ORG_ALIAS) if (re.test(n)) { const o = ORGS.find(x => x.name.toLowerCase() === v.toLowerCase()); return o ? o.name : v; }
  return '';
}
const HDR = { date: ['EntryDate', 'Boekdatum', 'Datum'], entry: ['EntryNumber', 'Boekstuknummer', 'Boekstuk'], journal: ['JournalCodeDescriptionCode', 'Dagboek', 'Dagboekcode'],
  journalName: ['JournalCodeDescriptionDescription', 'Dagboek omschrijving'], gl: ['GLAccountCodeDescriptionCode', 'Grootboekrekening', 'Grootboek'],
  glName: ['GLAccountCodeDescriptionDescription', 'Grootboekrekening omschrijving'], desc: ['Description', 'Omschrijving'], rel: ['AccountCodeNameDescription', 'Relatie', 'Naam'],
  amount: ['AmountDC', 'Bedrag', 'Bedrag DC'], cc: ['CostCenterCodeDescriptionDescription', 'Kostenplaats'] };
function decodeFile(buf) {
  const b = new Uint8Array(buf);
  if (b[0] === 0xFF && b[1] === 0xFE) return new TextDecoder('utf-16le').decode(buf);
  if (b[0] === 0xFE && b[1] === 0xFF) return new TextDecoder('utf-16be').decode(buf);
  return new TextDecoder('utf-8').decode(buf).replace(/^\uFEFF/, '');
}
function parseDelim(t) {
  const first = t.split(/\r?\n/, 1)[0], sep = ['\t', ';', ','].map(s => [s, first.split(s).length]).sort((a, b) => b[1] - a[1])[0][0];
  const rows = []; let row = [], f = '', q = false;
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { f += '"'; i++; } else q = false; } else f += c; }
    else if (c === '"') q = true; else if (c === sep) { row.push(f); f = ''; }
    else if (c === '\n' || c === '\r') { if (c === '\r' && t[i + 1] === '\n') i++; row.push(f); rows.push(row); row = []; f = ''; }
    else f += c;
  }
  if (f || row.length) { row.push(f); rows.push(row); }
  return rows.filter(r => r.some(x => x.trim() !== '')).map(r => r.map(x => x.trim()));
}
const numNL = s => { s = String(s).replace(/[\s\u00a0€]/g, ''); if (!s) return 0; if (s.includes(',')) s = s.replace(/\./g, '').replace(',', '.'); return Number(s) || 0; };
const isoD = (y, m, d) => `${y}-${String(m).padStart(2, '0')}-${String(d).padStart(2, '0')}`;
const lastDay = (y, m) => new Date(Date.UTC(+y, +m, 0)).getUTCDate();
const toIso = s => { const m = String(s).match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})/); if (m) return isoD(m[3], m[2], m[1]); const k = String(s).match(/^(\d{4})-(\d\d)-(\d\d)/); return k ? k[0] : ''; };
const okP = (a, b) => a && b && !isNaN(dnum(a)) && !isNaN(dnum(b)) && a <= b && dnum(b) - dnum(a) <= 400;
function periodOf(desc, date) {
  let m = desc.match(/\((\d\d)-(\d\d)-(\d{4})\s*-\s*(\d\d)-(\d\d)-(\d{4})\)/);
  if (m) { const a = isoD(m[3], m[2], m[1]), b = isoD(m[6], m[5], m[4]); if (okP(a, b)) return [a, b, 'uit omschrijving']; }
  m = desc.match(/\b(20\d\d)(0[1-9]|1[0-2])([0-3]\d)\s*-\s*(20\d\d)(0[1-9]|1[0-2])([0-3]\d)\b/);
  if (m) { const a = isoD(m[1], m[2], m[3]), b = isoD(m[4], m[5], m[6]); if (okP(a, b)) return [a, b, 'uit omschrijving']; }
  m = desc.match(/\b(20\d\d)(0[1-9]|1[0-2])\s*-\s*(20\d\d)(0[1-9]|1[0-2])\b/);
  if (m) { const a = isoD(m[1], m[2], 1), b = isoD(m[3], m[4], lastDay(m[3], m[4])); if (okP(a, b)) return [a, b, 'uit omschrijving']; }
  m = desc.match(/\b(\d\d)(0[1-9]|1[0-2])([0-3]\d)\s*-\s*(\d\d)(0[1-9]|1[0-2])([0-3]\d)\b/);
  if (m) { const a = isoD(2000 + +m[1], m[2], m[3]), b = isoD(2000 + +m[4], m[5], m[6]); if (okP(a, b)) return [a, b, 'uit omschrijving']; }
  const y = date.slice(0, 4), mo = date.slice(5, 7);
  return [isoD(y, mo, 1), isoD(y, mo, lastDay(y, mo)), 'maand van boekdatum'];
}
function catKey(gl, desc, rel) {
  const d = (desc + ' ' + rel).toLowerCase(); let cat = 'Andere', key = KDEF, org = '', warn = '';
  if (/hovenier|passerstr|sint-joris|cellebroers/.test(desc.toLowerCase())) return { cat: 'Andere', key: KEYS[3], org: 'De Overmolen', warn: 'Andere locatie? → De Overmolen' };
  switch (gl.slice(0, 4)) {
    case '6101': cat = 'Schoonmaak'; break;
    case '6104': cat = /poets|ramen|\biss\b|schoonma|opkuis/.test(d) ? 'Schoonmaak' : /afval|container|propret/.test(d) ? 'Afval' : 'Herstelling / onderhoud'; break;
    case '6105': cat = /lift|pareyn/.test(d) ? 'Lift' : 'Herstelling / onderhoud'; break;
    case '6107': cat = 'Keuringen / brandveiligheid'; break;
    case '6108': cat = 'Herstelling / onderhoud'; break;
    case '6110': cat = 'Huur / huurvermindering'; key = KEYS[3]; org = 'De Overmolen'; warn = 'Huur – niet verdelen'; break;
    case '6115': cat = 'Afval'; break;
    case '6121': cat = 'Water (algemene meter)'; break;
    case '6122': if (/\bista\b/.test(d)) { cat = 'Ista – opname/beheer'; } else { cat = 'Gas / verwarming'; key = KEYS[1]; } break;   // Ista-facturen = gewone huurlasten
    case '6123': cat = 'Elektriciteit algemeen'; break;
    case '6125': cat = 'Telefoon / internet'; break;
    case '6141': cat = 'Verzekering'; break;
  }
  return { cat, key, org, warn };
}
function parseExact(text) {
  const rows = parseDelim(text); if (rows.length < 2) throw new Error('Leeg bestand.');
  const h = rows[0].map(x => x.toLowerCase()), ix = {};
  for (const [k, names] of Object.entries(HDR)) { const i = names.map(n => h.indexOf(n.toLowerCase())).find(i => i >= 0); ix[k] = i == null ? -1 : i; }
  const need = ['date', 'entry', 'gl', 'desc', 'amount'].filter(k => ix[k] < 0);
  if (need.length) throw new Error('Kolommen niet gevonden (' + need.join(', ') + '). Gevonden kolommen: ' + rows[0].join(' | '));
  const get = (r, k) => ix[k] >= 0 ? (r[ix[k]] || '') : '';
  const lines = rows.slice(1).map(r => ({ date: toIso(get(r, 'date')), entry: get(r, 'entry'), journal: get(r, 'journal'), journalName: get(r, 'journalName'), gl: get(r, 'gl'), glName: get(r, 'glName'),
    desc: get(r, 'desc'), rel: get(r, 'rel'), amount: numNL(get(r, 'amount')), cc: get(r, 'cc') })).filter(l => l.date && l.entry);
  const groups = {}; for (const l of lines) (groups[l.entry + '|' + l.gl] = groups[l.entry + '|' + l.gl] || []).push(l);
  for (const [k, ls] of Object.entries(groups)) { ls.sort((a, b) => a.desc.localeCompare(b.desc) || a.amount - b.amount); ls.forEach((l, i) => { l.uid = 'EX:' + k + '|' + (i + 1); }); }
  const costs = [], vos = [], info = {};
  for (const l of lines) {
    const amt = Math.round(l.amount * 100) / 100;
    if (/^6/.test(l.gl)) {
      const [from, to, psrc] = periodOf(l.desc, l.date), ck = catKey(l.gl, l.desc, l.rel);
      costs.push({ uid: l.uid, entry: l.entry, bdate: l.date, gl: (l.gl + ' ' + l.glName).trim(), desc: l.desc || l.glName, supplier: l.rel, amount: amt, from, to, psrc, cat: ck.cat, key: ck.key, org: ck.org, warn: ck.warn, src: 'Exact', invoice: '', meter: '', note: '' });
    } else if (l.gl === '702000' && /^huur$/i.test(l.journal)) {
      vos.push({ uid: l.uid, entry: l.entry, date: l.date, rel: l.rel, org: mapOrg(l.rel), amount: -amt, desc: l.desc, src: 'Exact' });
    } else { const k = (l.journalName || l.journal) + ' · ' + l.gl + ' ' + l.glName; info[k] = info[k] || { n: 0, sum: 0 }; info[k].n++; info[k].sum += amt; }
  }
  // lijn zonder periode in de omschrijving: neem de periode van een andere lijn van hetzelfde boekstuk
  for (const c of costs) if (c.psrc !== 'uit omschrijving') { const p = costs.find(x => x.entry === c.entry && x.psrc === 'uit omschrijving'); if (p) { c.from = p.from; c.to = p.to; c.psrc = 'zoals boekstuk ' + c.entry; } }
  const dates = lines.map(l => l.date).sort();
  return { lines, costs, vos, info, from: dates[0], to: dates[dates.length - 1], cc: [...new Set(lines.map(l => l.cc).filter(Boolean))] };
}
function diffImport(P) {
  const R = { cAdd: [], cSame: [], cChg: [], cDel: [], vAdd: [], vSame: [], vChg: [], vDel: [] };
  const inR = d => d && d >= P.from && d <= P.to;
  const ec = new Map(COSTS.filter(c => c.uid).map(c => [c.uid, c])), seenC = new Set();
  for (const c of P.costs) { seenC.add(c.uid); const o = ec.get(c.uid);
    if (!o) R.cAdd.push(c); else if (Math.abs((Number(o.amount) || 0) - c.amount) > 0.004 || (o.desc || '') !== c.desc || (o.supplier || '') !== c.supplier) R.cChg.push({ n: c, o, on: true }); else R.cSame.push(c); }
  for (const o of COSTS) if (o.src === 'Exact' && o.uid && !seenC.has(o.uid) && inR(o.bdate)) R.cDel.push({ o, on: true });
  const ev = new Map(VOS.filter(v => v.uid).map(v => [v.uid, v])), seenV = new Set();
  for (const v of P.vos) { seenV.add(v.uid); const o = ev.get(v.uid);
    if (!o) R.vAdd.push(v); else if (Math.abs((Number(o.amount) || 0) - v.amount) > 0.004 || (o.desc || '') !== v.desc) R.vChg.push({ n: v, o, on: true }); else R.vSame.push(v); }
  for (const o of VOS) if (o.src === 'Exact' && o.uid && !seenV.has(o.uid) && inR(o.date)) R.vDel.push({ o, on: true });
  R.cAdd.forEach(c => { c.on = true; });
  return R;
}
let IMP = null;   // {file, P, R}
function bindImport() {
  if (bindImport.done) return; bindImport.done = true;
  $('#aImp').onclick = () => $('#aImpFile').click();
  $('#aImpFile').onchange = async e => {
    const f = e.target.files[0]; e.target.value = ''; if (!f) return;
    const box = $('#afrImp'); box.hidden = false; box.innerHTML = '<p>Bestand lezen…</p>';
    try { const P = parseExact(decodeFile(await f.arrayBuffer())); IMP = { file: f, P, R: diffImport(P) }; renderImport(); }
    catch (er) { box.innerHTML = `<div class="occbox"><b class="err">Import mislukt:</b> ${esc(er.message)} <button class="btn" id="impX">Sluiten</button></div>`; $('#impX').onclick = () => { box.hidden = true; }; }
  };
}
function renderImport() {
  const box = $('#afrImp'); if (!IMP) { box.hidden = true; box.innerHTML = ''; return; }
  const { file, P, R } = IMP, sum = a => a.reduce((s, x) => s + (Number(x.amount) || 0), 0);
  const unmapped = [...new Set(P.vos.filter(v => !v.org).map(v => v.rel))];
  const opt = (arr, v) => arr.map(x => `<option${x === v ? ' selected' : ''}>${esc(x)}</option>`).join('');
  const orgNames = ORGS.map(o => o.name);
  const cRow = (c, i, kind) => `<tr class="${kind}" data-i="${i}" data-k="${kind}">
    <td><input type="checkbox" class="ion" ${c.on !== false ? 'checked' : ''}></td><td class="mono" style="font-size:12px">${dfmt(c.bdate)}<br>${esc(c.entry)}</td>
    <td class="wrap">${esc(c.desc)}<br><span class="empty" style="font-size:11px">${esc(c.supplier)} · ${esc(c.gl)}</span>${c.warn ? `<br><span class="imp-warn">⚠ ${esc(c.warn)}</span>` : ''}</td>
    <td class="num mono">${eur2(c.amount)}</td>
    <td><select class="icat">${opt(CATS, c.cat)}</select></td>
    <td><select class="ikey">${opt(KEYS, c.key)}</select><select class="iorg" ${c.key === KEYS[3] ? '' : 'hidden'}>${opt(orgNames, c.org)}</select></td>
    <td style="white-space:nowrap"><input type="date" class="ifrom" value="${c.from}"><br><input type="date" class="ito" value="${c.to}"><br><span class="empty" style="font-size:10px">${esc(c.psrc || '')}</span></td></tr>`;
  const infoRows = Object.entries(P.info).sort((a, b) => b[1].n - a[1].n).map(([k, v]) => `<li><span>${esc(k)}</span><span class="mono">${v.n} lijnen · ${eur2(v.sum)}</span></li>`).join('');
  const vByOrg = {}; [...R.vAdd, ...R.vSame, ...R.vChg.map(x => x.n)].forEach(v => { const k = v.org || '(' + v.rel + ')'; vByOrg[k] = (vByOrg[k] || 0) + v.amount; });
  box.innerHTML = `<div class="occbox imp">
    <div class="imp-head"><b>Import Exact: ${esc(file.name)}</b> <span class="chip">${P.lines.length} lijnen</span> <span class="chip">boekdatum ${dfmt(P.from)} – ${dfmt(P.to)}</span>${P.cc.length ? `<span class="chip">kostenplaats: ${esc(P.cc.join(', '))}</span>` : ''}</div>
    <div class="cards imp-cards">
      <div class="card"><span class="k">Kosten (6xxxxx)</span><span class="v">${eur2(sum(P.costs))}</span><span class="k">${P.costs.length} lijnen</span></div>
      <div class="card"><span class="k">Kosten nieuw</span><span class="v">${R.cAdd.length}</span><span class="k">${eur2(sum(R.cAdd))}</span></div>
      <div class="card"><span class="k">Al aanwezig</span><span class="v">${R.cSame.length + R.vSame.length}</span><span class="k">ongewijzigd, overgeslagen</span></div>
      <div class="card ${R.cChg.length + R.vChg.length ? 'warnc' : ''}"><span class="k">Gewijzigd in Exact</span><span class="v">${R.cChg.length + R.vChg.length}</span></div>
      <div class="card ${R.cDel.length + R.vDel.length ? 'badc' : ''}"><span class="k">Niet meer in Exact</span><span class="v">${R.cDel.length + R.vDel.length}</span></div>
      <div class="card"><span class="k">Voorschotten (702000 · HUUR)</span><span class="v">${eur2(sum(P.vos))}</span><span class="k">${P.vos.length} lijnen · ${R.vAdd.length} nieuw</span></div>
    </div>
    ${unmapped.length ? `<div class="imp-map"><b class="imp-warn">Koppel deze huurders aan een organisatie:</b>${unmapped.map(n => `<label>${esc(n)} → <select class="imap" data-n="${esc(n)}"><option value="">– kies –</option>${opt(orgNames, '')}</select></label>`).join('')}</div>` : ''}
    ${R.cChg.length ? `<h3>Gewijzigd in Exact (${R.cChg.length})</h3><div class="tscroll"><table class="imp-t"><thead><tr><th>Bijwerken</th><th>Boekstuk</th><th>Omschrijving</th><th class="num">Oud</th><th class="num">Nieuw</th></tr></thead><tbody>${R.cChg.map((x, i) => `<tr class="chg" data-i="${i}" data-k="cChg"><td><input type="checkbox" class="ion" ${x.on ? 'checked' : ''}></td><td class="mono">${esc(x.n.entry)}</td><td class="wrap">${esc(x.o.desc)}${x.o.desc !== x.n.desc ? '<br>→ ' + esc(x.n.desc) : ''}</td><td class="num mono">${eur2(Number(x.o.amount) || 0)}</td><td class="num mono">${eur2(x.n.amount)}</td></tr>`).join('')}</tbody></table></div><p class="empty" style="font-size:12px">Bij bijwerken veranderen enkel bedrag, omschrijving, leverancier en grootboek. Jouw categorie, sleutel en periode blijven.</p>` : ''}
    ${R.cDel.length ? `<h3>Niet meer in Exact (${R.cDel.length})</h3><div class="tscroll"><table class="imp-t"><thead><tr><th>Verwijderen</th><th>Boekstuk</th><th>Omschrijving</th><th class="num">Bedrag</th></tr></thead><tbody>${R.cDel.map((x, i) => `<tr class="del" data-i="${i}" data-k="cDel"><td><input type="checkbox" class="ion" ${x.on ? 'checked' : ''}></td><td class="mono">${esc(x.o.entry)}</td><td class="wrap">${esc(x.o.desc)}</td><td class="num mono">${eur2(Number(x.o.amount) || 0)}</td></tr>`).join('')}</tbody></table></div>` : ''}
    ${R.cAdd.length ? `<h3>Nieuwe kosten (${R.cAdd.length}) <span class="empty" style="font-size:12px;font-weight:400">– controleer categorie, sleutel en periode; vinkje uit = niet importeren</span></h3>
      <div class="tscroll imp-scroll"><table class="imp-t"><thead><tr><th></th><th>Boekdatum / stuk</th><th>Omschrijving</th><th class="num">Bedrag</th><th>Categorie</th><th>Sleutel</th><th>Periode</th></tr></thead><tbody>${R.cAdd.map((c, i) => cRow(c, i, 'cAdd')).join('')}</tbody></table></div>` : ''}
    ${P.vos.length ? `<h3>Voorschotten per organisatie (in dit bestand)</h3><ul class="mini imp-vo">${Object.entries(vByOrg).sort((a, b) => b[1] - a[1]).map(([k, v]) => `<li><span>${esc(k)}</span><span class="mono">${eur2(v)}</span></li>`).join('')}</ul>` : ''}
    ${infoRows ? `<details style="margin-top:10px"><summary>Niet geïmporteerd (ter info): huur, vorige afrekening, balansrekeningen</summary><ul class="mini" style="margin-top:6px">${infoRows}</ul></details>` : ''}
    <div class="btns" style="margin-top:12px"><button class="btn primary" id="impGo">Opslaan in SharePoint</button><button class="btn" id="impX">Annuleren</button><span id="impMsg"></span></div>
  </div>`;
  const arr = k => k === 'cAdd' ? R.cAdd : k === 'cChg' ? R.cChg : R.cDel;
  box.onchange = e => {
    const t = e.target;
    if (t.classList.contains('imap')) { if (t.value) { ORGMAP[t.dataset.n] = t.value; try { localStorage.setItem('lk.orgmap', JSON.stringify(ORGMAP)); } catch (er) {} [...P.vos].forEach(v => { if (v.rel === t.dataset.n) v.org = t.value; }); renderImport(); } return; }
    const tr = t.closest('tr[data-k]'); if (!tr) return; const it = arr(tr.dataset.k)[+tr.dataset.i];
    if (t.classList.contains('ion')) it.on = t.checked;
    const c = it; if (tr.dataset.k !== 'cAdd') return;
    if (t.classList.contains('icat')) { c.cat = t.value; c.key = t.value === 'Gas / verwarming' ? KEYS[1] : t.value === 'Elektriciteit eigen meter' || t.value === 'Huur / huurvermindering' ? KEYS[3] : KDEF; renderImport(); }
    if (t.classList.contains('ikey')) { c.key = t.value; if (c.key === KEYS[3] && !c.org) c.org = 'De Overmolen'; tr.querySelector('.iorg').hidden = c.key !== KEYS[3]; }
    if (t.classList.contains('iorg')) c.org = t.value;
    if (t.classList.contains('ifrom')) c.from = t.value;
    if (t.classList.contains('ito')) c.to = t.value;
  };
  $('#impX').onclick = () => { IMP = null; renderImport(); };
  $('#impGo').onclick = () => saveImport();
}
async function saveImport() {
  const { file, P, R } = IMP, msg = $('#impMsg'), btn = $('#impGo');
  if (P.vos.some(v => !v.org)) { msg.className = 'err'; msg.textContent = 'Koppel eerst alle huurders aan een organisatie.'; return; }
  if (R.cAdd.some(c => c.on && c.from > c.to)) { msg.className = 'err'; msg.textContent = "Een periode heeft 'Tot' vóór 'Van'."; return; }
  const ops = [];
  for (const c of R.cAdd) if (c.on) { const o = { ...c }; delete o.on; delete o.psrc; delete o.warn; if (o.key !== KEYS[3]) o.org = ''; ops.push({ kind: 'costs', op: 'add', obj: o }); }
  for (const x of R.cChg) if (x.on) ops.push({ kind: 'costs', op: 'upd', obj: { _id: x.o._id, amount: x.n.amount, desc: x.n.desc, supplier: x.n.supplier, gl: x.n.gl, entry: x.n.entry, bdate: x.n.bdate } });
  for (const x of R.cDel) if (x.on) ops.push({ kind: 'costs', op: 'del', obj: { _id: x.o._id } });
  for (const v of R.vAdd) ops.push({ kind: 'vo', op: 'add', obj: v });
  for (const x of R.vChg) if (x.on) ops.push({ kind: 'vo', op: 'upd', obj: { _id: x.o._id, amount: x.n.amount, desc: x.n.desc, org: x.n.org, date: x.n.date } });
  for (const x of R.vDel) if (x.on) ops.push({ kind: 'vo', op: 'del', obj: { _id: x.o._id } });
  btn.disabled = true; msg.className = '';
  try {
    if (ops.length) await LK.batch(ops, (a, b) => { msg.textContent = `Opslaan… ${a}/${b}`; });
    msg.textContent = 'Origineel bestand bewaren…';
    const stamp = new Date().toISOString().slice(0, 19).replace(/[:T]/g, '-');
    let note = '';
    try { await LK.uploadFile(`${window.LK_CONFIG.importFolder}/${stamp}_${file.name}`, file); note = 'Bestand: ' + window.LK_CONFIG.importFolder + '/' + stamp + '_' + file.name; } catch (er) { note = 'Bestand niet bewaard: ' + er.message; }
    const n = k => ops.filter(o => o.op === k).length;
    try { await LK.save('imp', { file: file.name, at: new Date().toLocaleString('nl-BE'), from: P.from, to: P.to, lines: P.lines.length, added: n('add'), changed: n('upd'), removed: n('del'), who: $('#who').textContent, note }); } catch (er) { note += ' · log niet bewaard: ' + er.message; }
    await reload(); IMP = null; renderAll(true);
    setStatus('live', `Import klaar: ${n('add')} nieuw · ${n('upd')} bijgewerkt · ${n('del')} verwijderd`);
    const box = $('#afrImp'); box.hidden = false; box.innerHTML = `<div class="occbox"><b class="saved">✓ Import opgeslagen</b> – ${n('add')} nieuw · ${n('upd')} bijgewerkt · ${n('del')} verwijderd. <span class="empty">${esc(note)}</span> <button class="btn" id="impX">Sluiten</button></div>`;
    $('#impX').onclick = () => { box.hidden = true; };
  } catch (er) { btn.disabled = false; msg.className = 'err'; msg.textContent = 'Niet (volledig) opgeslagen: ' + (er.status === 403 ? 'geen bewerkrechten.' : er.message) + ' Herlaad en importeer opnieuw: wat al bewaard is wordt herkend.'; }
}
let raf = 0;
function renderAll(keepPanel) {
  cancelAnimationFrame(raf);
  raf = requestAnimationFrame(() => { renderPlan(); if (!keepPanel || !document.activeElement?.closest('#panel')) renderPanel(); renderRooms(); renderDoors(); renderIsta(); renderOrgs(); renderOcc(); renderAfr(); });
}
boot().catch(e => { if (e.message === 'no lists') return; console.error(e); setStatus('ro', 'Laden mislukt'); $('#notice').hidden = false; $('#notice').textContent = 'Het plan kon niet geladen worden: ' + (e.message || e) + '. Herlaad de pagina of meld je opnieuw aan.'; });
})();
