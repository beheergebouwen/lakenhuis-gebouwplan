// Data layer: Microsoft 365 login (MSAL) + SharePoint lists via Microsoft Graph.
(function () {
  const C = window.LK_CONFIG;
  const SCOPES = ['Sites.ReadWrite.All', 'User.Read'];
  const msal = new window.msal.PublicClientApplication({
    auth: { clientId: C.clientId, authority: 'https://login.microsoftonline.com/' + C.tenantId, redirectUri: location.origin + location.pathname.replace(/index\.html$/, '') },
    cache: { cacheLocation: 'localStorage' }
  });
  const ready = msal.initialize();
  let account = null, siteId = null; const listIds = {};

  // Field mapping: app field -> SharePoint internal column name
  const MAP = {
    rooms:   { code: 'Title', floor: 'Verdieping', name: 'Ruimte', use: 'Gebruik', zone: 'Zone', org: 'Organisatie', m2plan: 'M2Plan', src: 'Bron', m2meas: 'M2Gemeten', ok: 'Gecontroleerd', note: 'Opmerking', oldCode: 'OudeCode' },
    doors:   { code: 'Title', floor: 'Verdieping', fr: 'Van', to: 'Naar', arch: 'CodeArchitect', type: 'DeurType', size: 'Maat', lock: 'SlotBadge', key: 'SleutelNr', checked: 'Gecontroleerd', note: 'Opmerking', oldCode: 'OudeCode' },
    meters:  { nr: 'Title', unit: 'IstaUnit', istaCode: 'IstaCode', istaRoom: 'IstaRuimte', old: 'OudeStand', new: 'NieuweStand', units: 'Verbruik', room: 'RuimteCode', note: 'Opmerking' },
    orgs:    { name: 'Title', color: 'Kleur', type: 'Soort' },
    settings:{ key: 'Title', period: 'Periode', heatCost: 'Verwarmingskost', totalUnits: 'EenhedenTotaal', rentTotal: 'HuurVerdelers', meterCount: 'AantalMeters' }
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
    catch (e) { await msal.acquireTokenRedirect({ scopes: SCOPES, account }); }
  }
  async function g(path, opt = {}) {
    const url = path.startsWith('http') ? path : 'https://graph.microsoft.com/v1.0' + path;
    const r = await fetch(url, { ...opt, headers: { Authorization: 'Bearer ' + await token(), ...(opt.body && !(opt.body instanceof Blob) ? { 'Content-Type': 'application/json' } : {}), ...(opt.headers || {}) } });
    if (!r.ok) { let m = r.status + ''; try { m = (await r.json()).error.message; } catch (e) {} const err = new Error(m); err.status = r.status; throw err; }
    return r.status === 204 ? null : (r.headers.get('content-type') || '').includes('json') ? r.json() : r.blob();
  }
  const toApp = (kind, f, id) => { const o = { _id: id }; for (const [a, s] of Object.entries(MAP[kind])) o[a] = f[s] ?? (typeof f[s] === 'boolean' ? false : ''); return o; };
  const toSp = (kind, o) => { const f = {}; for (const [a, s] of Object.entries(MAP[kind])) if (a in o) f[s] = o[a] === '' && /M2|Verbruik|Stand|kost|Eenheden|Huur|Aantal/.test(s) ? null : o[a]; return f; };

  const API = {
    async login() {
      await ready;
      const res = await msal.handleRedirectPromise();
      account = res?.account || msal.getAllAccounts()[0] || null;
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
    logout() { msal.logoutRedirect({ account }); },
    get siteId() { return siteId; }
  };
  window.LK = API;
})();
