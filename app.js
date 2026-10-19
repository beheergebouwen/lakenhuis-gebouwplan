
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
let OCC = [];                                    // {_id, room, org, from, to, note}  (datums YYYY-MM-DD, leeg = onbekend / lopend)
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
  const [rooms, doors, meters, orgs, settings, occ] = await Promise.all(['rooms', 'doors', 'meters', 'orgs', 'settings', 'occ'].map(k => LK.load(k).catch(() => null)));
  OCC = (occ || []).filter(o => o.room && o.org);
  if (!rooms) { setStatus('ro', 'Lijsten niet gevonden'); const n = $('#notice'); n.hidden = false; n.textContent = 'De SharePoint-lijsten bestaan nog niet. Laat IT eenmalig setup/ uitvoeren (zie handleiding).'; throw new Error('no lists'); }
  S.rooms = Object.fromEntries(rooms.map(r => [r.code, r])); S.doors = Object.fromEntries(doors.map(d => [d.code, d]));
  S.meters = Object.fromEntries(meters.map(m => [m.nr, m]));
  ORGS = (orgs || []).filter(o => o.name).map(o => ({ name: o.name, color: o.color || '#e6e2d6', type: o.type }));
  S.settings = (settings || []).find(x => x.key === 'ista') || S.settings;
  for (const k of ['heatCost', 'totalUnits', 'rentTotal', 'meterCount']) S.settings[k] = Number(S.settings[k]) || 0;
  if (!per) per = istaPeriod();
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
const price = () => S.settings.heatCost / S.settings.totalUnits;
const rent = () => S.settings.rentTotal / S.settings.meterCount;
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
  const setPer = () => { if ($('#pFrom').value && $('#pTo').value && $('#pFrom').value <= $('#pTo').value) { per = { from: $('#pFrom').value, to: $('#pTo').value }; renderOrgs(); } };
  $('#pFrom').onchange = setPer; $('#pTo').onchange = setPer;
  $('#pIsta').onclick = () => { per = istaPeriod(); $('#pFrom').value = per.from; $('#pTo').value = per.to; renderOrgs(); };
  $('#pYear').onclick = () => { const y = peil.slice(0, 4); per = { from: y + '-01-01', to: y + '-12-31' }; $('#pFrom').value = per.from; $('#pTo').value = per.to; renderOrgs(); };
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
    <div><div class="sub">Bezetting (historiek)</div>
      ${occOf(r.code).length ? `<table class="occ"><thead><tr><th>Organisatie</th><th>Van</th><th>Tot</th></tr></thead><tbody>${occOf(r.code).map(o => `<tr data-id="${esc(o._id)}" class="${isOpen(o, peil) ? 'now' : ''}"><td><span class="swatch" style="background:${orgColor(o.org)}"></span>${esc(o.org)}</td>
        <td><input type="date" data-k="from" value="${esc(o.from)}" ${dis} aria-label="Van"></td><td><input type="date" data-k="to" value="${esc(o.to)}" ${dis} aria-label="Tot"></td></tr>`).join('')}</tbody></table>
        <p class="empty" style="margin:4px 0 0;font-size:12px">Leeg 'Van' = datum onbekend · leeg 'Tot' = loopt nog.</p>` : '<p class="empty" style="margin:0;font-size:13px">Nog geen bezetting geregistreerd.</p>'}
      ${canWrite ? `<form class="edit" id="fOcc" style="margin-top:8px"><div class="two"><div class="row"><label for="oOrg">Wijziging</label><select id="oOrg"><option value="">Vertrek (ruimte wordt leeg)</option>${ORGS.map(o => `<option>${esc(o.name)}</option>`).join('')}</select></div>
        <div class="row"><label for="oDate">Datum</label><input type="date" id="oDate" value="${peil}" required></div></div>
        <div class="btns"><button class="btn" type="submit">Registreren</button><span id="oMsg"></span></div></form>` : ''}</div>
    <div><div class="sub">Ista-meters (${ms.length})</div>
      ${ms.length ? `<ul class="mini">${ms.map(m => `<li><span class="mono">${esc(m.nr)}</span><span>${Number(m.units).toLocaleString('nl-BE')} eenh.</span><span class="num">${eur(meterCost(m))}</span></li>`).join('')}
        <li><b>Verwarming ${esc(S.settings.period)}</b><b class="num">${eur(heat)}</b></li></ul>` : '<p class="empty" style="margin:0;font-size:13px">Nog geen meter gekoppeld. Koppel meters in de tab Ista-meters.</p>'}</div>
    <div><div class="sub">Deuren (${doors.length})</div>${doors.length ? `<ul class="mini">${doors.map(d => `<li><span class="mono">${esc(d.code)}</span><span>${esc(d.fr === sel ? '→ ' + d.to : '← ' + d.fr)}</span><span>${d.checked ? '✓' : ''}</span></li>`).join('')}</ul>` : '<p class="empty" style="margin:0;font-size:13px">Geen deuren gekoppeld.</p>'}</div>
    ${r.oldCode ? `<p class="empty" style="margin:0;font-size:12px">Oude code: ${esc(r.oldCode)}${r.updatedAt ? ' · gewijzigd ' + new Date(r.updatedAt).toLocaleString('nl-BE') : ''}</p>` : ''}`;
  const P2 = $('#panel');
  const occTable = P2.querySelector('table.occ');
  if (occTable) occTable.onchange = async e => {
    const inp = e.target.closest('input[data-k]'); if (!inp) return; const o = OCC.find(x => String(x._id) === inp.closest('tr').dataset.id);
    const next = { ...o, [inp.dataset.k]: inp.value };
    if (next.from && next.to && next.to < next.from) { inp.style.outline = '2px solid var(--bad)'; setStatus('ro', "'Tot' ligt vóór 'Van'"); return; }
    try { await save('occ', o._id, next); Object.assign(o, next); syncRoomOrg(r); renderAll(true); } catch (er) { inp.style.outline = '2px solid var(--bad)'; setStatus('ro', 'Niet opgeslagen: ' + (er.code || er.message)); }
  };
  if ($('#fOcc')) $('#fOcc').onsubmit = async e => {
    e.preventDefault(); const msg = $('#oMsg'); const D = $('#oDate').value, org = $('#oOrg').value; if (!D) return;
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
function renderOrgs() {
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
  $('#tOrgs').innerHTML = `<thead><tr><th>Organisatie</th><th class="num">Ruimten in periode</th><th class="num">Nu (${dfmt(peil)})</th><th class="num">Gem. privatief m²</th><th>Aandeel</th><th class="num">Aandeel gemeensch. m²</th><th class="num">Totaal m²</th><th class="num">Verwarming Ista ${esc(S.settings.period || '')}</th></tr></thead><tbody>` +
    rows.map(x => `<tr><td><span class="swatch" style="background:${x.o.color}"></span>${esc(x.o.name)}</td><td class="num">${x.n}</td><td class="num">${x.now}</td><td class="num mono">${m2f(x.avg)}</td>
      <td style="min-width:140px"><div class="bar"><span style="width:${(x.share * 100).toFixed(1)}%;background:${x.o.color}"></span></div><span class="mono" style="font-size:12px">${(x.share * 100).toFixed(2)}%</span></td>
      <td class="num mono">${m2f(x.com)}</td><td class="num mono">${m2f(x.avg + x.com)}</td><td class="num mono">${x.heat ? eur(x.heat) : '–'}</td></tr>`).join('') +
    `</tbody><tfoot><tr><th>Leegstand</th><th></th><th></th><th class="num mono">${m2f(vacant)}</th><th class="mono" style="font-size:12px">${privTot ? (vacant / privTot * 100).toFixed(2) : 0}%</th><th></th><th></th><th></th></tr></tfoot>`;
}
let raf = 0;
function renderAll(keepPanel) {
  cancelAnimationFrame(raf);
  raf = requestAnimationFrame(() => { renderPlan(); if (!keepPanel || !document.activeElement?.closest('#panel')) renderPanel(); renderRooms(); renderDoors(); renderIsta(); renderOrgs(); });
}
boot().catch(e => { if (e.message === 'no lists') return; console.error(e); setStatus('ro', 'Laden mislukt'); $('#notice').hidden = false; $('#notice').textContent = 'Het plan kon niet geladen worden: ' + (e.message || e) + '. Herlaad de pagina of meld je opnieuw aan.'; });
})();
