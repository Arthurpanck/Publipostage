// Pilotage d'un Grist local (Docker) : comptes de test, API, navigateur.
const http = require('node:http');
const fs = require('node:fs');
const path = require('node:path');
let chromium;
try { ({ chromium } = require('playwright')); } catch (e) { ({ chromium } = require('/opt/node-tools/node_modules/playwright')); }

const BASE = process.env.GRIST_URL || 'http://localhost:8484';
const OWNER = 'owner@test.fr', EDITOR = 'editor@test.fr', VIEWER = 'viewer@test.fr';
const sleep = ms => new Promise(r => setTimeout(r, ms));
const keys = {};
let browser;

// Sert le widget (racine du dépôt) sur un port libre, origine distincte de Grist
function serveWidget() {
  const root = path.resolve(process.env.WIDGET_ROOT || path.join(__dirname, '..', '..'));
  const types = { '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css' };
  const server = http.createServer((req, res) => {
    const file = path.join(root, decodeURIComponent(new URL(req.url, 'http://x').pathname));
    if (!file.startsWith(root) || !fs.existsSync(file) || fs.statSync(file).isDirectory()) { res.writeHead(404); return res.end(); }
    res.writeHead(200, { 'Content-Type': types[path.extname(file)] || 'application/octet-stream' });
    fs.createReadStream(file).pipe(res);
  });
  return new Promise(resolve => server.listen(0, '127.0.0.1', () => resolve({ url: `http://127.0.0.1:${server.address().port}/index.html`, close: () => server.close() })));
}

async function getBrowser() {
  return browser || (browser = await chromium.launch(process.env.CHROMIUM_PATH ? { executablePath: process.env.CHROMIUM_PATH } : {}));
}

// Navigateur neuf connecté (= déconnexion / reconnexion)
async function session(email) {
  const ctx = await (await getBrowser()).newContext({ viewport: { width: 1300, height: 850 } });
  const page = await ctx.newPage();
  await page.goto(`${BASE}/test/login?username=${encodeURIComponent(email)}&next=${encodeURIComponent(BASE + '/')}`);
  return { ctx, page, email };
}

async function apiKey(email) {
  if (keys[email]) return keys[email];
  const s = await session(email);
  const headers = { 'X-Requested-With': 'XMLHttpRequest' };
  const read = async r => (await r.text()).replace(/"/g, '');
  keys[email] = await read(await s.page.request.get(`${BASE}/api/profile/apiKey`, { headers }))
    || await read(await s.page.request.post(`${BASE}/api/profile/apiKey`, { headers }));
  await s.ctx.close();
  return keys[email];
}

async function api(email, method, route, body) {
  const headers = { Authorization: 'Bearer ' + await apiKey(email) };
  if (body && !(body instanceof FormData)) { headers['Content-Type'] = 'application/json'; body = JSON.stringify(body); }
  const r = await fetch(BASE + route, { method, headers, body });
  const text = await r.text();
  if (!r.ok) throw new Error(`${method} ${route} -> ${r.status} ${text.slice(0, 300)}`);
  try { return JSON.parse(text); } catch (e) { return text; }
}
const apply = (d, actions) => api(OWNER, 'POST', `/api/docs/${d.doc}/apply`, actions);
const sql = (doc, q) => api(OWNER, 'GET', `/api/docs/${doc}/sql?q=${encodeURIComponent(q)}`).then(r => r.records.map(x => x.fields));
const workspace = async () => (await api(OWNER, 'GET', '/api/orgs/current/workspaces'))[0].id;
const customView = (url, widgetOptions = null) => JSON.stringify({ mode: 'url', url, access: 'full', widgetOptions });

async function addWidget(d, url, view) {
  const res = await apply(d, [['CreateViewSection', d.table, view, 'custom', null, null]]);
  const { sectionRef, viewRef } = res.retValues[0];
  await apply(d, [['UpdateRecord', '_grist_Views_section', sectionRef, { options: JSON.stringify({ customView: customView(url) }) }]]);
  return { ...d, view: viewRef, sections: [sectionRef] };
}

// Document : table Personnes, page avec n widgets publipostage, un éditeur et un lecteur
async function createDoc(url, { widgets = 1 } = {}) {
  const doc = await api(OWNER, 'POST', `/api/workspaces/${await workspace()}/docs`, { name: 'publipostage-' + Date.now() });
  await api(OWNER, 'POST', `/api/docs/${doc}/tables`, { tables: [{ id: 'Personnes', columns: [{ id: 'Nom' }, { id: 'Ville' }] }] });
  await api(OWNER, 'POST', `/api/docs/${doc}/tables/Personnes/records`, { records: [{ fields: { Nom: 'Alice', Ville: 'Lyon' } }, { fields: { Nom: 'Bob', Ville: 'Paris' } }] });
  const [t] = await sql(doc, "select id, primaryViewId from _grist_Tables where tableId='Personnes'");
  let d = { doc, table: t.id, view: t.primaryViewId, sections: [] };
  for (let i = 0; i < widgets; i++) d.sections.push((await addWidget(d, url, d.view)).sections[0]);
  await api(OWNER, 'PATCH', `/api/docs/${doc}/access`, { delta: { users: { [EDITOR]: 'editors', [VIEWER]: 'viewers' } } });
  return d;
}

async function openDoc(s, d) {
  await s.page.goto(`${BASE}/doc/${d.doc}/p/${d.view}`);
  await s.page.waitForSelector('.test-gristdoc', { timeout: 30000 }).catch(() => {});
  await sleep(1500);
}

async function widgetFrame(s, n = 0) {
  for (let i = 0; i < 60; i++) {
    const frame = s.page.frames().filter(f => /127\.0\.0\.1:\d+\/(.*\/)?index\.html/.test(f.url()))[n];
    if (frame && await frame.evaluate(() => typeof state !== 'undefined').catch(() => false)) return frame;
    await sleep(500);
  }
  throw new Error('widget introuvable');
}

const widgetState = f => f.evaluate(() => ({
  name: state.templateName,
  toast: document.getElementById('toast').hidden ? '' : document.getElementById('toast').textContent,
  banner: document.getElementById('persistBanner')?.hidden === false ? document.getElementById('persistBanner').textContent : '',
  modal: !document.getElementById('modalOverlay').hidden,
}));

// Modèle chargé par le widget n (attend la fin du chargement)
async function templateIn(s, n = 0, wait = 8000) {
  const f = await widgetFrame(s, n);
  const t0 = Date.now();
  let st = await widgetState(f);
  while (!st.name && Date.now() - t0 < wait) { await sleep(300); st = await widgetState(f); }
  return st;
}

// Choisit un fichier dans le widget et attend la fin de l'enregistrement
async function upload(s, file, n = 0) {
  const f = await widgetFrame(s, n);
  await f.setInputFiles('#templateFile', { name: file.name, mimeType: 'application/octet-stream', buffer: file.buffer });
  const t0 = Date.now();
  let st = await widgetState(f);
  while (Date.now() - t0 < 20000 && !/enregistré|Lecture seule|Erreur|trop lourd|Enregistrer/.test(st.toast + st.banner)) { await sleep(300); st = await widgetState(f); }
  return st;
}

// Ouvre le document dans un navigateur neuf et lit le modèle du widget n
async function check(email, d, n = 0) {
  const s = await session(email);
  try { await openDoc(s, d); return await templateIn(s, n); } finally { await s.ctx.close(); }
}

async function savedOptions(d) {
  const rows = await sql(d.doc, `select id, options from _grist_Views_section where id in (${d.sections.join(',')})`);
  return rows.map(r => JSON.parse(JSON.parse(r.options || '{}').customView || '{}').widgetOptions || null);
}

module.exports = { BASE, OWNER, EDITOR, VIEWER, sleep, serveWidget, session, api, apply, sql, workspace, createDoc, addWidget, openDoc, widgetFrame, widgetState, templateIn, upload, check, savedOptions, close: () => browser && browser.close() };
