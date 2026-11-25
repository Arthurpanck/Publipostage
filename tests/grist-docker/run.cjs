// Le modèle reste-t-il attaché au widget ? Scénarios contre un vrai Grist.
// Usage : voir README.md. Filtre facultatif : node run.cjs "Migration"
const G = require('./grist.cjs');
const { loadApp, template, para } = require('../helpers.cjs');
const { OWNER, EDITOR, VIEWER } = G;

const app = loadApp();
const docx = (name, text, extra) => {
  const zip = new app.PizZip(template(app, para(text)));
  if (extra) zip.file('word/media/photo.bin', extra);
  return { name, buffer: Buffer.from(zip.generate({ type: 'uint8array', compression: 'STORE' })) };
};
const A = docx('modele-A.docx', 'Modèle A : {Nom} de {Ville}');
const B = docx('modele-B.docx', 'Modèle B : {Nom}');
const GROS = docx('modele-gros.docx', 'Gros modèle : {Nom}', require('node:crypto').randomBytes(20 * 1024 * 1024));

const purge = d => G.api(OWNER, 'POST', `/api/docs/${d.doc}/attachments/removeUnused?expiredonly=0`);
async function uploadAs(email, d, file, n = 0) {
  const s = await G.session(email);
  await G.openDoc(s, d);
  const st = await G.upload(s, file, n);
  const saveButton = await s.page.locator('.test-section-menu-small-btn-save').count();
  await s.ctx.close();
  return { ...st, saveButton };
}
const is = (st, file) => st.name === file.name;
const reserveStructure = d => G.apply(d, [['AddRecord', '_grist_ACLResources', -1, { tableId: '*', colIds: '*' }],
  ['AddRecord', '_grist_ACLRules', null, { resource: -1, aclFormula: 'user.Access != OWNER', permissionsText: '-S', memo: '' }]]);
const templateRows = d => G.sql(d.doc, 'select Widget, Modele, Empreinte, Fichier from Publipostage_Modeles_NE_PAS_SUPPRIMER').catch(() => null);

// Modèle à l'ancienne : pièce jointe téléversée + templateId dans les options
async function legacyDoc(url) {
  const d = await G.createDoc(url);
  const form = new FormData();
  form.append('upload', new Blob([A.buffer]), A.name);
  const [id] = await G.api(OWNER, 'POST', `/api/docs/${d.doc}/attachments`, form);
  const [row] = await G.sql(d.doc, `select options from _grist_Views_section where id=${d.sections[0]}`);
  const options = JSON.parse(row.options), view = JSON.parse(options.customView);
  view.widgetOptions = { templateId: id, templateName: A.name };
  options.customView = JSON.stringify(view);
  await G.apply(d, [['UpdateRecord', '_grist_Views_section', d.sections[0], { options: JSON.stringify(options) }]]);
  return d;
}

const SCENARIOS = {
  'Envoi puis rechargement (déconnexion / reconnexion)': async url => {
    const d = await G.createDoc(url); const up = await uploadAs(OWNER, d, A);
    const r = await G.check(OWNER, d);
    return [is(r, A) && up.saveButton === 0, `« ${up.toast} » · bouton Enregistrer de Grist allumé : ${up.saveButton} · rechargé : ${r.name}`];
  },
  'Envoi puis ouverture par un autre éditeur': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    const r = await G.check(EDITOR, d); return [is(r, A), r.name];
  },
  'Envoi puis ouverture par un lecteur': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    const r = await G.check(VIEWER, d); return [is(r, A), r.name];
  },
  'Envoi par un éditeur, ouverture par un lecteur': async url => {
    const d = await G.createDoc(url); await uploadAs(EDITOR, d, A);
    const r = await G.check(VIEWER, d); return [is(r, A), r.name];
  },
  'Ouverture immédiate après l\'envoi': async url => {
    const d = await G.createDoc(url);
    const s = await G.session(OWNER); await G.openDoc(s, d); await G.upload(s, A); await s.ctx.close();
    const r = await G.check(EDITOR, d); return [is(r, A), r.name];
  },
  'Règle d\'accès : colonne masquée aux lecteurs': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    await G.apply(d, [['AddRecord', '_grist_ACLResources', -1, { tableId: 'Personnes', colIds: 'Ville' }],
      ['AddRecord', '_grist_ACLRules', null, { resource: -1, aclFormula: "user.Access == 'viewers'", permissionsText: '-R', memo: '' }]]);
    const r = await G.check(VIEWER, d); return [is(r, A), r.name];
  },
  'Éditeur : envoi, puis ouverture par un lecteur et par le propriétaire': async url => {
    const d = await G.createDoc(url); const ed = await uploadAs(EDITOR, d, A);
    const v = await G.check(VIEWER, d), o = await G.check(OWNER, d);
    return [/enregistré/.test(ed.toast) && is(v, A) && is(o, A), `« ${ed.toast} » · lecteur : ${v.name} · propriétaire : ${o.name}`];
  },
  'Structure réservée : l\'éditeur est prévenu, puis enregistre une fois le widget ouvert par le propriétaire': async url => {
    const d = await G.createDoc(url); await reserveStructure(d);
    const before = await uploadAs(EDITOR, d, A);
    const noTable = !(await templateRows(d));
    await G.check(OWNER, d); // ouverture seule : prépare la table
    const ed = await uploadAs(EDITOR, d, B);
    const v = await G.check(VIEWER, d), [opts] = await G.savedOptions(d);
    return [/propriétaire doit ouvrir/.test(before.banner) && before.saveButton === 0 && noTable && /enregistré/.test(ed.toast) && !ed.banner && is(v, B),
      `1er envoi : « ${before.banner.slice(0, 60)}… » · après ouverture par le propriétaire : « ${ed.toast} » · lecteur : ${v.name} · options : ${JSON.stringify(opts)}`];
  },
  'Structure réservée : le propriétaire charge, l\'éditeur remplace': async url => {
    const d = await G.createDoc(url); await reserveStructure(d);
    await uploadAs(OWNER, d, A); const ed = await uploadAs(EDITOR, d, B);
    const r = await G.check(OWNER, d), e = await G.check(EDITOR, d);
    return [/enregistré/.test(ed.toast) && is(r, B) && is(e, B), `« ${ed.toast} » · propriétaire : ${r.name} · éditeur : ${e.name}`];
  },
  'Migration dans un document à structure réservée : faite à l\'ouverture par le propriétaire': async url => {
    const d = await legacyDoc(url); await reserveStructure(d);
    const ed = await G.check(EDITOR, d); const [still] = await G.savedOptions(d);
    const ow = await G.check(OWNER, d); const [after] = await G.savedOptions(d);
    await purge(d); const v = await G.check(VIEWER, d);
    return [is(ed, A) && !!still.templateId && is(ow, A) && !after.templateData && !after.templateId && after.templateName === A.name && is(v, A),
      `éditeur : ${ed.name} (options inchangées) · propriétaire : ${ow.name} → ${Object.keys(after).sort()} · lecteur après purge : ${v.name}`];
  },
  'Migration : modèle en base64 dans les options (version précédente)': async url => {
    const d = await G.createDoc(url);
    const [row] = await G.sql(d.doc, `select options from _grist_Views_section where id=${d.sections[0]}`);
    const options = JSON.parse(row.options), view = JSON.parse(options.customView);
    view.widgetOptions = { templateName: A.name, templateHash: 'ancien', templateData: A.buffer.toString('base64') };
    options.customView = JSON.stringify(view);
    await G.apply(d, [['UpdateRecord', '_grist_Views_section', d.sections[0], { options: JSON.stringify(options) }]]);
    const v0 = await G.check(VIEWER, d); const [still] = await G.savedOptions(d);
    const ed = await G.check(EDITOR, d); const [after] = await G.savedOptions(d);
    await purge(d); const v = await G.check(VIEWER, d);
    return [is(v0, A) && !!still.templateData && is(ed, A) && !after.templateData && is(v, A),
      `lecteur : ${v0.name} (rien d'écrit) · éditeur : ${ed.name} → ${Object.keys(after).sort()} · lecteur après purge : ${v.name}`];
  },
  'Purge des pièces jointes inutilisées (au-delà de 7 jours)': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    await purge(d); await G.api(OWNER, 'POST', `/api/docs/${d.doc}/force-reload`);
    const r = await G.check(OWNER, d); return [is(r, A), r.name];
  },
  'Copie du document': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    const doc = await G.api(OWNER, 'POST', `/api/docs/${d.doc}/copy`, { workspaceId: await G.workspace(), documentName: 'copie' });
    const r = await G.check(OWNER, { ...d, doc }); return [is(r, A), r.name];
  },
  'Copie du document comme modèle (sans données) : message clair': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
    const doc = await G.api(OWNER, 'POST', `/api/docs/${d.doc}/copy`, { workspaceId: await G.workspace(), documentName: 'copie', asTemplate: true });
    const s = await G.session(OWNER); await G.openDoc(s, { ...d, doc }); await G.sleep(2500);
    const st = await G.widgetState(await G.widgetFrame(s)); await s.ctx.close();
    return [!st.name && /« modele-A.docx » n'est plus disponible/.test(st.toast) && st.modal, `« ${st.toast.slice(0, 90)}… »`];
  },
  'Deux widgets neufs sur la même table': async url => {
    const d = await G.createDoc(url, { widgets: 2 });
    const s = await G.session(OWNER); await G.openDoc(s, d);
    await G.upload(s, A, 0); await G.upload(s, B, 1); await s.ctx.close();
    const s2 = await G.session(EDITOR); await G.openDoc(s2, d);
    const r0 = await G.templateIn(s2, 0), r1 = await G.templateIn(s2, 1); await s2.ctx.close();
    return [is(r0, A) && is(r1, B), `${r0.name} / ${r1.name}`];
  },
  'Deux widgets sur deux pages (même table)': async url => {
    const d = await G.createDoc(url), d2 = await G.addWidget(d, url, 0);
    await uploadAs(OWNER, d2, B); await uploadAs(OWNER, d, A);
    const r1 = await G.check(EDITOR, d), r2 = await G.check(EDITOR, d2);
    return [is(r1, A) && is(r2, B), `page 1 : ${r1.name} · page 2 : ${r2.name}`];
  },
  'Cinq widgets, cinq modèles différents (trois sur la même page, deux sur d\'autres pages)': async url => {
    const d = await G.createDoc(url, { widgets: 3 });
    const p2 = await G.addWidget(d, url, 0), p3 = await G.addWidget(d, url, 0);
    const files = [1, 2, 3, 4, 5].map(i => docx(`modele-${i}.docx`, `Texte du modèle ${i} pour {Nom}`));
    const s = await G.session(OWNER); await G.openDoc(s, d);
    for (let i = 0; i < 3; i++) await G.upload(s, files[i], i);
    await G.openDoc(s, p2); await G.upload(s, files[3]);
    await G.openDoc(s, p3); await G.upload(s, files[4]);
    await s.ctx.close();
    // Nom du modèle et texte de l'aperçu de chaque widget, pour un éditeur puis un lecteur
    const shown = async (v, n) => {
      const st = await G.templateIn(v, n); const f = await G.widgetFrame(v, n);
      for (let k = 0; k < 40 && !/Texte du modèle/.test(await f.evaluate(() => document.getElementById('preview-container').innerText)); k++) await G.sleep(250);
      const text = await f.evaluate(() => document.getElementById('preview-container').innerText);
      return `${st.name}:${(/modèle (\d)/.exec(text) || [])[1]}`;
    };
    const seen = [];
    for (const email of [EDITOR, VIEWER]) {
      const v = await G.session(email); await G.openDoc(v, d);
      for (let i = 0; i < 3; i++) seen.push(await shown(v, i));
      await G.openDoc(v, p2); seen.push(await shown(v, 0));
      await G.openDoc(v, p3); seen.push(await shown(v, 0));
      await v.ctx.close();
    }
    const rows = await templateRows(d);
    const expected = [...files, ...files].map((f, i) => `${f.name}:${(i % 5) + 1}`);
    return [JSON.stringify(seen) === JSON.stringify(expected) && rows.length === 5 && new Set(rows.map(r => r.Widget)).size === 5,
      `${seen.slice(0, 5).join(' / ')} (éditeur, idem lecteur) · lignes de la table : ${rows.length}`];
  },
  'Remplacement du modèle': async url => {
    const d = await G.createDoc(url); await uploadAs(OWNER, d, A); await uploadAs(OWNER, d, B);
    const r = await G.check(EDITOR, d); return [is(r, B), r.name];
  },
  'Deux utilisateurs ouverts : le changement arrive en direct': async url => {
    const d = await G.createDoc(url);
    const s1 = await G.session(OWNER), s2 = await G.session(EDITOR);
    await G.openDoc(s1, d); await G.openDoc(s2, d);
    await G.upload(s1, A); const a = await G.templateIn(s2);
    await G.upload(s1, B); await G.sleep(2000); const b = await G.widgetState(await G.widgetFrame(s2));
    await s1.ctx.close(); await s2.ctx.close();
    return [is(a, A) && is(b, B), `l'autre voit ${a.name} puis ${b.name}`];
  },
  'Annuler (Ctrl+Z) dans Grist rétablit le modèle précédent': async url => {
    const d = await G.createDoc(url);
    const s = await G.session(OWNER); await G.openDoc(s, d);
    await G.upload(s, A); await G.upload(s, B);
    await s.page.locator('.test-undo').first().dispatchEvent('click'); await G.sleep(2500);
    await s.ctx.close();
    const r = await G.check(OWNER, d); return [is(r, A), r.name];
  },
  'Gros modèle (20 Mo, pas de limite côté widget)': async url => {
    const d = await G.createDoc(url); const t0 = Date.now(); await uploadAs(OWNER, d, GROS);
    const r = await G.check(EDITOR, d); const [opts] = await G.savedOptions(d); const size = JSON.stringify(opts).length;
    return [is(r, GROS) && size < 500, `${r.name} · ${((Date.now() - t0) / 1000).toFixed(1)} s · options du widget : ${size} octets`];
  },
  'Lecteur qui choisit un modèle : session seulement, rien d\'écrit': async url => {
    const d = await G.createDoc(url); const up = await uploadAs(VIEWER, d, A);
    const r = await G.check(OWNER, d);
    return [/Lecture seule/.test(up.toast) && !r.name, `« ${up.toast} » · propriétaire : ${r.name}`];
  },
  'Migration : ancien modèle (pièce jointe) ouvert par un éditeur, puis purge': async url => {
    const d = await legacyDoc(url);
    const first = await G.check(EDITOR, d); const [saved] = await G.savedOptions(d);
    await purge(d);
    const r = await G.check(VIEWER, d);
    return [is(first, A) && is(r, A) && !saved.templateId, `1re ouverture : ${first.name} · options : ${Object.keys(saved).sort()} · après purge : ${r.name}`];
  },
  'Migration : ancien modèle ouvert d\'abord par un lecteur': async url => {
    const d = await legacyDoc(url); const r = await G.check(VIEWER, d); const [saved] = await G.savedOptions(d);
    return [is(r, A) && !!saved.templateId, `${r.name} · options inchangées : ${Object.keys(saved).sort()}`];
  },
  'Migration : pièce jointe déjà purgée → message clair': async url => {
    const d = await legacyDoc(url); await purge(d);
    const s = await G.session(OWNER); await G.openDoc(s, d);
    await G.sleep(3000); const st = await G.widgetState(await G.widgetFrame(s)); await s.ctx.close();
    return [!st.name && /n'est plus disponible/.test(st.toast) && st.modal, `« ${st.toast} »`];
  },
};
SCENARIOS['Export .grist puis import'] = async url => {
  const d = await G.createDoc(url); await uploadAs(OWNER, d, A);
  const s = await G.session(OWNER);
  const bin = await (await s.page.request.get(`${G.BASE}/api/docs/${d.doc}/download`)).body();
  await s.ctx.close();
  const form = new FormData(); form.append('upload', new Blob([bin]), 'export.grist');
  const res = await G.api(OWNER, 'POST', `/api/workspaces/${await G.workspace()}/import`, form);
  const r = await G.check(OWNER, { ...d, doc: res.id }); return [is(r, A), r.name];
};

(async () => {
  const only = process.argv[2] ? new RegExp(process.argv[2]) : null;
  const widget = await G.serveWidget();
  let failed = 0;
  for (const [name, fn] of Object.entries(SCENARIOS)) {
    if (only && !only.test(name)) continue;
    let ok, detail;
    try { [ok, detail] = await fn(widget.url); } catch (e) { ok = false; detail = String(e).slice(0, 300); }
    if (!ok) failed++;
    console.log(`${ok ? 'ok  ' : 'ÉCHEC'}  ${name}  —  ${detail}`);
  }
  await G.close(); widget.close();
  console.log(failed ? `\n${failed} scénario(s) en échec` : '\nTous les scénarios passent');
  process.exit(failed ? 1 : 0);
})();
