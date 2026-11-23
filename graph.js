// Data layer: Microsoft 365 login (MSAL) + SharePoint lists via Microsoft Graph.
(function () {
  const C = window.LK_CONFIG;
  const SCOPES = ['Sites.ReadWrite.All', 'User.Read'];
  const msal = new window.msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: location.origin + location.pathname.replace(/index\.html$/, '') },
    cache: { cacheLocation: 'localStorage' }
  });
  const ready = msal.initialize();
  // In Teams (of een andere iframe) werken redirects niet: daar aanmelden via een pop-upvenster
  const IFR = (() => { try { return window.self !== window.top; } catch (e) { return true; } })();
  const IS_POPUP = !!window.opener && window.opener !== window && /^msal\./.test(window.name || '');
  function signInButton() {
    return new Promise((ok, no) => {
      const d = document.createElement('div');
      d.style.cssText = 'position:fixed;inset:0;z-index:99999;display:flex;align-items:center;justify-content:center;background:rgba(244,243,239,.96);font-family:inherit';
      d.innerHTML = '<div style="text-align:center;max-width:420px;padding:24px"><h2 style="margin:0 0 8px">Lakenhuis Gebouwplan</h2><p style="margin:0 0 16px">Meld je aan met je De Overmolen-account.</p><button id="lkSignIn" style="font-size:16px;padding:10px 18px;border-radius:8px;border:0;background:#2f6f9f;color:#fff;cursor:pointer">Aanmelden met Microsoft</button><p id="lkSignErr" style="color:#b00;font-size:13px;margin-top:12px"></p></div>';
      document.body.appendChild(d);
      d.querySelector('#lkSignIn').onclick = async () => {
        try { const r = await msal.loginPopup({ scopes: SCOPES }); d.remove(); ok(r.account); }
        catch (e) { d.querySelector('#lkSignErr').textContent = 'Aanmelden niet gelukt (' + (e.errorCode || e.message) + '). Sta pop-ups toe of open de app in de browser.'; }
      };
    });
  }
  let account = null, siteId = null; const listIds = {};

  // Field mapping: app field -> SharePoint internal column name
  const MAP = {
    rooms:   { code: 'Title', floor: 'Verdieping', name: 'Ruimte', use: 'Gebruik', zone: 'Zone', org: 'Organisatie', m2plan: 'M2Plan', src: 'Bron', m2meas: 'M2Gemeten', ok: 'Gecontroleerd', note: 'Opmerking', oldCode: 'OudeCode' },
    doors:   { code: 'Title', floor: 'Verdieping', fr: 'Van', to: 'Naar', arch: 'CodeArchitect', type: 'DeurType', size: 'Maat', lock: 'SlotBadge', key: 'SleutelNr', checked: 'Gecontroleerd', note: 'Opmerking', oldCode: 'OudeCode' },
    meters:  { nr: 'Title', unit: 'IstaUnit', istaCode: 'IstaCode', istaRoom: 'IstaRuimte', old: 'OudeStand', new: 'NieuweStand', units: 'Verbruik', room: 'RuimteCode', note: 'Opmerking' },
    orgs:    { name: 'Title', color: 'Kleur', type: 'Soort' },
    costs:   { desc: 'Title', cat: 'Categorie', from: 'Van', to: 'Tot', amount: 'Bedrag', key: 'Sleutel', org: 'Organisatie', supplier: 'Leverancier', invoice: 'Factuurnr', meter: 'Meter', note: 'Opmerking', uid: 'Uid', entry: 'Boekstuk', bdate: 'Boekdatum', gl: 'Grootboek', src: 'Bron', cu: 'Kostendrager', vat: 'Btw' },
    vo:      { org: 'Title', date: 'Datum', amount: 'Bedrag', desc: 'Omschrijving', rel: 'Relatie', entry: 'Boekstuk', uid: 'Uid', src: 'Bron' },
    imp:     { file: 'Title', at: 'Tijdstip', from: 'Van', to: 'Tot', lines: 'Lijnen', added: 'Nieuw', changed: 'Gewijzigd', removed: 'Verwijderd', who: 'Door', note: 'Opmerking' },
    occ:     { room: 'Title', org: 'Organisatie', from: 'Van', to: 'Tot', note: 'Opmerking' },
    rent:    { org: 'Title', from: 'Van', to: 'Tot', base: 'Beginhuur', baseIdx: 'Beginindex', idxMonth: 'Indexmaand', m2: 'M2Contract', note: 'Opmerking', email: 'Email', parent: 'HoofdContract' },
    idx:     { month: 'Title', value: 'Index2013', src: 'Bron' },
    settings:{ key: 'Title', period: 'Periode', heatCost: 'Verwarmingskost', totalUnits: 'EenhedenTotaal', rentTotal: 'HuurVerdelers', meterCount: 'AantalMeters', rules: 'Regels' }
  };
  const t = (n, type) => ({ name: n, [type || 'text']: type === 'number' ? { decimalPlaces: 'automatic' } : type === 'boolean' ? {} : type === 'choice' ? undefined : {} });
  const COLS = {
    rooms: [t('Verdieping'), t('Ruimte'), { name: 'Gebruik', choice: { choices: ['Verhuurbaar', 'Gemeenschappelijk', 'Technisch / berging', 'Buitenruimte'] } }, t('Zone'), t('Organisatie'),
      t('M2Plan', 'number'), { name: 'Bron', choice: { choices: ['Op plan', 'Berekend uit maten', 'Te meten', 'Gemeten'] } }, t('M2Gemeten', 'number'), t('Gecontroleerd', 'boolean'),
      { name: 'Opmerking', text: { allowMultipleLines: true } }, t('OudeCode')],
    doors: [t('Verdieping'), t('Van'), t('Naar'), t('CodeArchitect'), t('DeurType'), t('Maat'),
      { name: 'SlotBadge', choice: { choices: ['Cilinder', 'Badgelezer', 'Cilinder + badgelezer', 'Codeklavier', 'Geen slot', 'Onbekend'] } }, t('SleutelNr'), t('Gecontroleerd', 'boolean'),
      { name: 'Opmerking', text: { allowMultipleLines: true } }, t('OudeCode')],
    meters: [t('IstaUnit'), t('IstaCode'), t('IstaRuimte'), t('OudeStand', 'number'), t('NieuweStand', 'number'), t('Verbruik', 'number'), t('RuimteCode'), { name: 'Opmerking', text: { allowMultipleLines: true } }],
    orgs: [t('Kleur'), t('Soort')],
    settings: [t('Periode'), t('Verwarmingskost', 'number'), t('EenhedenTotaal', 'number'), t('HuurVerdelers', 'number'), t('AantalMeters', 'number')]
  };

  async function token() {
    try { return (await msal.acquireTokenSilent({ scopes: SCOPES, account })).accessToken; }
    catch (e) { if (IFR) return (await msal.acquireTokenPopup({ scopes: SCOPES, account })).accessToken; await msal.acquireTokenRedirect({ scopes: SCOPES, account }); }
  }
  async function g(path, opt = {}) {
    const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
    const r = await fetch(url, { ...opt, headers: { Authorization: 'Bearer ' + await token(), ...(opt.body && !(opt.body instanceof Blob) ? { 'Content-Type': 'application/json' } : {}), ...(opt.headers || {}) } });
    if (!r.ok) { let m = r.status + ''; try { m = (await r.json()).error.message; } catch (e) {} const err = new Error(m); err.status = r.status; throw err; }
    return r.status === 204 ? null : (r.headers.get('content-type') || '').includes('json') ? r.json() : r.blob();
  }
  const DK = { occ: { Van: 1, Tot: 1 }, costs: { Van: 1, Tot: 1 }, vo: { Datum: 1 }, rent: { Van: 1, Tot: 1 } };
  const ymd = v => v ? new Intl.DateTimeFormat('en-CA', { timeZone: 'Europe/Brussels' }).format(new Date(v)) : '';
  const toApp = (kind, f, id) => { const o = { _id: id }; for (const [a, s] of Object.entries(MAP[kind])) o[a] = DK[kind]?.[s] ? ymd(f[s]) : (f[s] ?? (typeof f[s] === 'boolean' ? false : '')); return o; };
  const toSp = (kind, o) => { const f = {}; for (const [a, s] of Object.entries(MAP[kind])) if (a in o) f[s] = DK[kind]?.[s] ? (o[a] ? o[a] + 'T12:00:00Z' : null) : (o[a] === '' && /M2|Verbruik|Stand|kost|Eenheden|Huur|Aantal|Bedrag|Lijnen|Nieuw|Gewijzigd|Verwijderd/.test(s) ? null : o[a]); return f; };

  const API = {
    async login() {
      await ready;
      if (IS_POPUP) return new Promise(() => {});          // dit is het aanmeld-pop-upvenster: niets doen, MSAL sluit het zelf
      const res = IFR ? null : await msal.handleRedirectPromise();
      account = res?.account || msal.getAllAccounts()[0] || null;
      if (!account && IFR) { try { account = (await msal.ssoSilent({ scopes: SCOPES })).account; } catch (e) { account = await signInButton(); } }
      if (!account) { await msal.loginRedirect({ scopes: SCOPES }); return null; }
      msal.setActiveAccount(account);
      const site = await g(`/sites/${C.siteHostname}:${C.sitePath}`); siteId = site.id;
      return { name: account.name, user: account.username };
    },
    async listId(kind) {
      if (listIds[kind]) return listIds[kind];
      const r = await g(`/sites/${siteId}/lists?$filter=displayName eq '${C.lists[kind]}'&$select=id`);
      return (listIds[kind] = r.value[0]?.id || null);
    },
    async load(kind) {
      const id = await this.listId(kind); if (!id) return null;
      let url = `/sites/${siteId}/lists/${id}/items?expand=fields&$top=999`, out = [];
      while (url) { const r = await g(url); out = out.concat(r.value.map(i => toApp(kind, i.fields, i.id))); url = r['@odata.nextLink']; }
      return out;
    },
    async canEdit(kind) {
      try { const id = await this.listId(kind); const r = await g(`/sites/${siteId}/lists/${id}?$select=id`); return !!r; } catch (e) { return false; }
    },
    async save(kind, obj) {
      const id = await this.listId(kind);
      if (obj._id) return g(`/sites/${siteId}/lists/${id}/items/${obj._id}/fields`, { method: 'PATCH', body: JSON.stringify(toSp(kind, obj)) });
      const r = await g(`/sites/${siteId}/lists/${id}/items`, { method: 'POST', body: JSON.stringify({ fields: toSp(kind, obj) }) }); obj._id = r.id; return r;
    },
    async remove(kind, id) { const lid = await this.listId(kind); return g(`/sites/${siteId}/lists/${lid}/items/${id}`, { method: 'DELETE' }); },
    async image(file) { return URL.createObjectURL(await g(`/sites/${siteId}/drive/root:/${C.planFolder}/${file}:/content`)); },
    // ---- one-time setup (setup/index.html)
    async createList(kind) {
      if (await this.listId(kind)) return 'bestaat al';
      const r = await g(`/sites/${siteId}/lists`, { method: 'POST', body: JSON.stringify({ displayName: C.lists[kind], columns: COLS[kind], list: { template: 'genericList' } }) });
      listIds[kind] = r.id; return 'aangemaakt';
    },
    async addMany(kind, rows, onStep) {
      const id = await this.listId(kind);
      for (let i = 0; i < rows.length; i += 20) {
        const reqs = rows.slice(i, i + 20).map((o, k) => ({ id: String(k + 1), method: 'POST', url: `/sites/${siteId}/lists/${id}/items`, headers: { 'Content-Type': 'application/json' }, body: { fields: toSp(kind, o) } }));
        const r = await g('/$batch', { method: 'POST', body: JSON.stringify({ requests: reqs }) });
        const bad = r.responses.filter(x => x.status >= 300); if (bad.length) throw new Error(kind + ': ' + (bad[0].body?.error?.message || bad[0].status));
        onStep && onStep(Math.min(i + 20, rows.length), rows.length);
      }
    },
    async upload(file, blob) { return g(`/sites/${siteId}/drive/root:/${C.planFolder}/${file}:/content`, { method: 'PUT', body: blob, headers: { 'Content-Type': 'image/jpeg' } }); },
    async uploadFile(path, blob) { return g(`/sites/${siteId}/drive/root:/${path.split('/').map(encodeURIComponent).join('/')}:/content`, { method: 'PUT', body: blob, headers: { 'Content-Type': blob.type || 'application/octet-stream' } }); },
    async batch(ops, onStep) {               // ops: [{kind, op:'add'|'upd'|'del', obj}]
      const out = [];
      for (let i = 0; i < ops.length; i += 20) {
        const part = ops.slice(i, i + 20), reqs = [];
        for (let k = 0; k < part.length; k++) { const { kind, op, obj } = part[k], lid = await this.listId(kind), base = `/sites/${siteId}/lists/${lid}/items`;
          reqs.push(op === 'add' ? { id: String(k + 1), method: 'POST', url: base, headers: { 'Content-Type': 'application/json' }, body: { fields: toSp(kind, obj) } }
            : op === 'upd' ? { id: String(k + 1), method: 'PATCH', url: `${base}/${obj._id}/fields`, headers: { 'Content-Type': 'application/json' }, body: toSp(kind, obj) }
            : { id: String(k + 1), method: 'DELETE', url: `${base}/${obj._id}` }); }
        const r = await g('/$batch', { method: 'POST', body: JSON.stringify({ requests: reqs }) });
        const bad = r.responses.filter(x => x.status >= 300); if (bad.length) throw new Error((bad[0].body?.error?.message || bad[0].status) + ` (${bad.length} fout)`);
        out.push(...r.responses); onStep && onStep(Math.min(i + 20, ops.length), ops.length);
      }
      return out;
    },
    async mailDraft({ to, subject, html, att }) {     // concept in Outlook van de aangemelde gebruiker (aparte toestemming Mail.ReadWrite)
      const sc = { scopes: ['Mail.ReadWrite'], account };
      let tk; try { tk = (await msal.acquireTokenSilent(sc)).accessToken; } catch (e) { tk = (await msal.acquireTokenPopup(sc)).accessToken; }
      const body = { subject, body: { contentType: 'HTML', content: html }, toRecipients: String(to || '').split(/[;,\s]+/).filter(Boolean).map(a => ({ emailAddress: { address: a } })),
        attachments: (att || []).map(a => ({ '@odata.type': '#microsoft.graph.fileAttachment', name: a.name, contentType: 'application/pdf', contentBytes: a.b64 })) };
      const r = await fetch('https://graph.microsoft.com/v1.0/me/messages', { method: 'POST', headers: { Authorization: 'Bearer ' + tk, 'Content-Type': 'application/json' }, body: JSON.stringify(body) });
      if (!r.ok) { let m = r.status + ''; try { m = (await r.json()).error.message; } catch (e) {} throw new Error(m); }
      return r.json();
    },
    logout() { IFR ? msal.logoutPopup({ account }).then(() => location.reload()) : msal.logoutRedirect({ account }); },
    get siteId() { return siteId; }
  };
  window.LK = API;
})();
