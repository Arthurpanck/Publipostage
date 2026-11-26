const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { loadApp, template, para } = require('./helpers.cjs');

const WIDGET = 'https://widget.test/publipostage/index.html';
const BASE = 'https://grist.test/api/docs/doc1';
const section = (id, widgetOptions, { url = WIDGET, tableRef = 1 } = {}) =>
    ({ id, tableRef, parentKey: 'custom', options: JSON.stringify({ customView: JSON.stringify({ mode: 'url', url, widgetOptions }) }) });
const STRUCTURE = 'Blocked by table structure access rules';
const tick = (ms = 30) => new Promise(r => setTimeout(r, ms));

// Contrôleur réel (main.js) dans un document Grist simulé : sections de widget
// (_grist_Views_section), table Publipostage_Modeles_NE_PAS_SUPPRIMER,
// pièces jointes et règles d'accès. Grist renvoie au widget (onOptions) les
// options de sa propre section quand elles changent. Une action refusée l'est
// en entier. `downloadDelay` ralentit les téléchargements ; `beforeApply` et
// `afterApply` encadrent chaque action (un autre widget agit entre-temps).
function controller({ sections = [section(7, null)], own = 7, readonly = false, structure = true,
    modeles = null, attachments = {}, rules = [], uploadStatus = 200, downloadDelay = 0, beforeApply = () => {}, afterApply = () => {} } = {}) {
    const events = [];
    const rows = new Map(sections.map(s => [s.id, { ...s }]));
    const widgetOptions = id => JSON.parse(JSON.parse(rows.get(id).options).customView).widgetOptions;
    const files = new Map(Object.entries(attachments).map(([id, bytes]) => [+id, Buffer.from(bytes)]));
    let nextFile = 100;
    const doc = { modeles: modeles && modeles.map((r, i) => ({ id: i + 1, ...r })), columns: modeles ? ['Widget', 'Modele', 'Empreinte', 'Fichier'] : null, rawTable: false };
    let onOptions = () => {};
    const fetch = async (url, init = {}) => {
        const m = /\/attachments(?:\/(\d+)\/download)?\?auth=/.exec(url);
        assert.ok(m, url);
        if (init.method === 'POST') {
            assert.equal(readonly, false, 'un lecteur ne téléverse rien');
            if (uploadStatus !== 200) return { ok: false, status: uploadStatus };
            const id = nextFile++;
            files.set(id, Buffer.from(await init.body.get('upload').arrayBuffer()));
            events.push(`téléversé:${id}`);
            return { ok: true, json: async () => [id] };
        }
        const bytes = files.get(+m[1]);
        events.push(`téléchargé:${m[1]}`);
        if (downloadDelay) await tick(downloadDelay);
        return bytes ? { ok: true, arrayBuffer: async () => new Uint8Array(bytes).buffer } : { ok: false, status: 404 };
    };
    const app = loadApp({
        location: new URL(`${WIDGET}?access=full&readonly=${readonly}`), setTimeout, fetch, FormData, Blob,
        grist: { ready() {}, onRecord() {}, onRecords() {}, onOptions(fn) { onOptions = fn; },
            getOptions: async () => widgetOptions(own),
            getSelectedTableId: async () => 'Donnees',
            docApi: {
                getAccessToken: async () => ({ baseUrl: BASE, token: 'jeton' }),
                fetchTable: async t => {
                    if (t === '_grist_Tables') return { id: [1, 2, ...(doc.columns ? [3] : [])], tableId: ['Donnees', 'Autre', ...(doc.columns ? ['Publipostage_Modeles_NE_PAS_SUPPRIMER'] : [])] };
                    if (t === '_grist_Tables_column') return { id: (doc.columns || []).map((_, i) => i + 1), parentId: (doc.columns || []).map(() => 3), colId: doc.columns || [] };
                    if (t === '_grist_ACLRules') return { id: rules.map((_, i) => i + 1), permissionsText: rules };
                    if (t === 'Publipostage_Modeles_NE_PAS_SUPPRIMER') {
                        if (!doc.columns) throw new Error('Table not found "Publipostage_Modeles_NE_PAS_SUPPRIMER"');
                        const col = c => doc.modeles.map(r => r[c] === undefined ? null : r[c]);
                        return { id: col('id'), Widget: col('Widget'), Modele: col('Modele'), Empreinte: col('Empreinte'), Fichier: col('Fichier') };
                    }
                    const list = [...rows.values()];
                    return { id: list.map(r => r.id), tableRef: list.map(r => r.tableRef), parentKey: list.map(r => r.parentKey), options: list.map(r => r.options) };
                },
                applyUserActions: async actions => {
                    assert.equal(readonly, false, 'un lecteur n\'écrit rien');
                    const touchesStructure = actions.some(a => a[1] === '_grist_Views_section' || /Table|Column/.test(a[0]));
                    if (!structure && touchesStructure) throw new Error(STRUCTURE);
                    await beforeApply(actions, { doc, rows, widgetOptions });
                    events.push(`action:${actions.map(a => a[0] + (a[0] === 'UpdateRecord' ? ':' + a[1] : '')).join('+')}`);
                    for (const [name, table, id, values] of actions) {
                        if (name === 'AddRawTable') { doc.columns = ['A', 'B', 'C']; doc.modeles = []; doc.rawTable = true; }
                        else if (name === 'RemoveColumn') doc.columns = doc.columns.filter(c => c !== id);
                        else if (name === 'AddColumn') doc.columns.push(id);
                        else if (name === 'ModifyColumn') (doc.descriptions = doc.descriptions || {})[id] = values.description;
                        else if (table === 'Publipostage_Modeles_NE_PAS_SUPPRIMER' && name === 'AddRecord') doc.modeles.push({ id: doc.modeles.length + 1, ...values });
                        else if (table === 'Publipostage_Modeles_NE_PAS_SUPPRIMER') Object.assign(doc.modeles.find(r => r.id === id), values);
                        else {
                            rows.get(id).options = values.options;
                            if (id === own) setTimeout(() => onOptions(widgetOptions(own)), 5);
                        }
                    }
                    afterApply(actions, { doc, rows, widgetOptions });
                },
            } },
        initUi() {}, uiToast: (message, type) => events.push({ message, type }),
        uiPreparePreview: async blob => () => events.push({ aperçu: blob }),
        uiOpenModal() { events.push('modale'); }, uiCloseModal() {}, uiSetTemplate: name => events.push({ modèle: name }), uiEnableActions() {}, uiShowPreview() {},
        uiPreviewEmptyText() {}, uiSetWarnings() {}, uiPersistBanner: text => text && events.push({ bandeau: text }),
        PDFLib: { PDFDocument: { load: async () => { throw new Error('PDF invalide'); } } },
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8'), app);
    app.readFileAsBuffer = async file => file.buffer;
    const start = () => onOptions(widgetOptions(own));
    const fileOf = row => new Uint8Array(files.get(row.Fichier[1]));
    return { app, events, saved: widgetOptions, rows, doc, files, fileOf, start, state: () => vm.runInContext('state', app) };
}
const docx = (app, text = '{{Titre}}') => template(app, para(text));
const toasts = events => events.filter(e => e.message).map(e => `${e.type}:${e.message}`);
const writes = events => events.filter(e => typeof e === 'string' && /^(action|téléversé)/.test(e));
const banner = events => (events.find(e => e.bandeau) || {}).bandeau || '';

test('un DOCX invalide est refusé et le modèle enregistré est conservé', async () => {
    const { app, events, saved } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    await app.handleTemplateUpload({ name: 'invalide.docx', buffer: template(app, para('{Titre')) });
    assert.deepEqual(writes(events), []);
    assert.deepEqual(saved(7), { templateId: 42, templateName: 'ancien.docx' });
    assert.equal(events.at(-1).type, 'error');
});
test('un PDF invalide est refusé', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'invalide.pdf', buffer: new Uint8Array() });
    assert.deepEqual(writes(events), []);
    assert.equal(events.at(-1).type, 'error');
});
test('le format binaire .doc est refusé avec une consigne de conversion', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'ancien.doc', buffer: new Uint8Array() });
    assert.match(events.at(-1).message, /Enregistrez le fichier en .docx/);
});
test('pas de limite de taille côté widget : un modèle de 20 Mo est enregistré', async () => {
    const { app, events, doc, fileOf, start } = controller();
    start(); await tick();
    const zip = new app.PizZip(docx(app));
    zip.file('word/media/photo.bin', new Uint8Array(20 * 1024 * 1024).fill(7));
    const buffer = zip.generate({ type: 'uint8array', compression: 'STORE' });
    await app.handleTemplateUpload({ name: 'lourd.docx', buffer });
    assert.equal(fileOf(doc.modeles[0]).length, buffer.length);
    assert.deepEqual(toasts(events).at(-1), 'success:Modèle enregistré dans le document');
});
test('fichier refusé par le serveur Grist (trop lourd) : message clair, modèle pour la session', async () => {
    const { app, events, state, start } = controller({ uploadStatus: 413 });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'lourd.docx', buffer: docx(app) });
    assert.match(toasts(events).at(-1), /^error:.*trop lourd pour ce serveur Grist/);
    assert.match(banner(events), /session seulement/);
    assert.equal(state().templateName, 'lourd.docx');
});

test('le modèle est rangé en pièce jointe dans une table sans page ; les options ne gardent que son nom', async () => {
    const { app, events, saved, doc, fileOf, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx', autre: 'conserver' })], attachments: { 42: [1] } });
    start(); await tick();
    const buffer = docx(app);
    await app.handleTemplateUpload({ name: 'MODELE.DOCX', buffer });
    assert.equal(doc.rawTable, true, 'AddRawTable : pas de nouvelle page');
    assert.deepEqual(doc.columns, ['Widget', 'Modele', 'Empreinte', 'Fichier']);
    for (const col of doc.columns) assert.match(doc.descriptions[col], /ne pas la supprimer/, `description de ${col}`);
    const row = doc.modeles.find(r => r.Modele === 'MODELE.DOCX');
    assert.equal(row.Widget, 7);
    assert.deepEqual(fileOf(row), buffer);
    const options = saved(7);
    assert.deepEqual(Object.keys(options).sort(), ['autre', 'templateHash', 'templateName', 'templateWidget']);
    assert.equal(options.templateHash, row.Empreinte);
    assert.equal(options.templateWidget, 7);
    assert.match(options.templateHash, /^[0-9a-f]{24}$/);
    assert.equal(events.filter(e => /^action:.*AddRecord.*_grist_Views_section/.test(e)).length >= 1, true,
        'fichier et options dans la même action : un Ctrl+Z annule les deux');
    assert.deepEqual(toasts(events).at(-1), 'success:Modèle enregistré dans le document');
    assert.equal(app.getTemplateType('MODELE.PDF'), 'pdf');
});

test('remplacement du modèle : la ligne du widget est mise à jour', async () => {
    const { app, doc, fileOf, saved, start } = controller();
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app, '{A}') });
    const b = docx(app, '{B}');
    await app.handleTemplateUpload({ name: 'b.docx', buffer: b });
    assert.equal(doc.modeles.length, 1);
    assert.equal(doc.modeles[0].Modele, 'b.docx');
    assert.deepEqual(fileOf(doc.modeles[0]), b);
    assert.equal(saved(7).templateName, 'b.docx');
});

test('deux widgets neufs sur la même table : seul le sien reçoit le modèle', async () => {
    const autre = section(5, null), autreTable = section(6, null, { tableRef: 2 }), autreUrl = section(8, null, { url: 'https://ailleurs.test/w.html' });
    const { app, saved, rows, doc, start } = controller({ sections: [autre, autreTable, autreUrl, section(7, null)] });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(saved(7).templateName, 'a.docx');
    assert.deepEqual(doc.modeles.map(r => r.Widget), [7]);
    assert.deepEqual(saved(5), { templateWidget: 5 }, "l'autre widget de la table reçoit son numéro");
    assert.equal(rows.get(6).options, autreTable.options, 'seul sur sa table : inchangé');
    assert.equal(rows.get(8).options, autreUrl.options, 'autre widget personnalisé : inchangé');
});

test('lecture seule : rien n\'est écrit, le modèle sert pour la session', async () => {
    const { app, events, state, start } = controller({ readonly: true });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(state().templateName, 'a.docx');
    assert.deepEqual(writes(events), []);
    assert.match(toasts(events).at(-1), /^error:Lecture seule/);
});

test('structure réservée aux propriétaires, table déjà prête : l\'éditeur enregistre son modèle', async () => {
    const { app, events, doc, fileOf, saved, start } = controller({ structure: false, modeles: [] });
    start(); await tick();
    const buffer = docx(app);
    await app.handleTemplateUpload({ name: 'a.docx', buffer });
    assert.equal(doc.modeles.length, 1);
    assert.equal(doc.modeles[0].Widget, 7);
    assert.deepEqual(fileOf(doc.modeles[0]), buffer);
    assert.equal(saved(7), null, 'options non modifiables : inchangées');
    assert.deepEqual(toasts(events).at(-1), 'success:Modèle enregistré dans le document');
    assert.equal(banner(events), '');
});

test('structure réservée aux propriétaires, table absente : session seulement, un propriétaire doit ouvrir le widget', async () => {
    const { app, events, doc, state, start } = controller({ structure: false });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(state().templateName, 'a.docx');
    assert.equal(doc.columns, null);
    assert.match(banner(events), /Un propriétaire doit ouvrir ce widget une fois/);
    assert.deepEqual(writes(events), [], 'le fichier n\'est pas envoyé pour rien');
});

test('ancien format ouvert par un éditeur qui ne peut pas le reprendre : utilisé, rien n\'est envoyé', async () => {
    const { state, events, start } = controller({ structure: false, sections: [section(7, { templateName: 'm.docx', templateHash: 'abc', templateData: 'UEsDBA==' })] });
    start(); await tick();
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(writes(events), []);
});

test('structure réservée : à l\'ouverture par un propriétaire, la table est préparée pour les éditeurs', async () => {
    const { doc, saved, start } = controller({ rules: ['-S'] });
    start(); await tick();
    assert.deepEqual(doc.columns, ['Widget', 'Modele', 'Empreinte', 'Fichier']);
    assert.equal(doc.rawTable, true);
    assert.equal(saved(7), null, 'seul sur sa table : pas besoin de numéro');
});

test('structure réservée, widgets neufs sur la même table : numérotés par le propriétaire, l\'éditeur enregistre', async () => {
    const sections = [section(5, null), section(7, null), section(9, null)];
    const owner = controller({ sections, rules: ['-S'] });
    owner.start(); await tick(100);
    assert.deepEqual(owner.saved(7), { templateWidget: 7 });
    assert.deepEqual(owner.saved(5), { templateWidget: 5 });
    assert.deepEqual(owner.saved(9), { templateWidget: 9 });
    // L'éditeur ouvre le document tel que le propriétaire l'a laissé
    const editor = controller({ structure: false, modeles: [], sections: [5, 7, 9].map(id => ({ id, tableRef: 1, parentKey: 'custom', options: owner.rows.get(id).options })) });
    editor.start(); await tick();
    await editor.app.handleTemplateUpload({ name: 'a.docx', buffer: docx(editor.app) });
    assert.deepEqual(editor.doc.modeles.map(r => [r.Widget, r.Modele]), [[7, 'a.docx']]);
    assert.deepEqual(toasts(editor.events).at(-1), 'success:Modèle enregistré dans le document');
});

test('document sans règle : ouvrir le widget ne crée rien', async () => {
    const { doc, events, start } = controller();
    start(); await tick();
    assert.equal(doc.columns, null);
    assert.deepEqual(writes(events), []);
    assert.equal(events.includes('modale'), true);
});

test('structure réservée et deux widgets neufs sur la même table : message clair', async () => {
    const { app, events, doc, start } = controller({ structure: false, modeles: [], sections: [section(5, null), section(7, null)] });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(doc.modeles.length, 0);
    assert.match(banner(events), /Demandez à un propriétaire de charger un modèle/);
});

test('au chargement, seul le modèle du widget est téléchargé, une seule fois', async () => {
    const mine = [80, 75, 3, 4, 1], other = [80, 75, 3, 4, 2];
    const { events, state, start } = controller({
        sections: [section(5, { templateName: 'x.docx', templateHash: 'hx', templateWidget: 5 }), section(7, { templateName: 'm.docx', templateHash: 'hm', templateWidget: 7 })],
        modeles: [{ Widget: 5, Modele: 'x.docx', Empreinte: 'hx', Fichier: ['L', 1] }, { Widget: 7, Modele: 'm.docx', Empreinte: 'hm', Fichier: ['L', 2] }],
        attachments: { 1: other, 2: mine } });
    start(); start(); await tick(); start(); await tick();
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(new Uint8Array(state().templateBuffer), new Uint8Array(mine));
    assert.deepEqual(events.filter(e => /^téléchargé/.test(e)), ['téléchargé:2']);
    assert.deepEqual(writes(events), []);
});

test('éditeur d\'un document à structure réservée : sa ligne l\'emporte sur des options restées anciennes', async () => {
    const { state, start } = controller({ structure: false,
        sections: [section(7, { templateName: 'ancien.docx', templateHash: 'h1', templateWidget: 7 })],
        modeles: [{ Widget: 7, Modele: 'nouveau.docx', Empreinte: 'h2', Fichier: ['L', 2] }], attachments: { 2: [9] } });
    start(); await tick();
    assert.equal(state().templateName, 'nouveau.docx');
});

test('widget dupliqué : il reprend le modèle désigné par ses options', async () => {
    const options = { templateName: 'm.docx', templateHash: 'hm', templateWidget: 5 };
    const { state, doc, saved, start } = controller({ sections: [section(5, options), section(7, options)],
        modeles: [{ Widget: 5, Modele: 'm.docx', Empreinte: 'hm', Fichier: ['L', 2] }], attachments: { 2: [9] } });
    start(); await tick(3000);
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(doc.modeles.map(r => r.Widget).sort(), [5, 7], 'puis une ligne propre, sans recopier le fichier');
    assert.deepEqual(doc.modeles.map(r => r.Fichier[1]), [2, 2]);
    assert.equal(saved(7).templateWidget, 7);
    assert.equal(saved(5).templateWidget, 5);
});

test('copie sans données (table vidée) : le widget dit quel modèle recharger', async () => {
    const { events, start } = controller({ sections: [section(7, { templateName: 'perdu.docx', templateHash: 'h', templateWidget: 7 })], modeles: [] });
    start(); await tick();
    assert.match(toasts(events).at(-1), /« perdu.docx » n'est plus disponible/);
    assert.equal(events.includes('modale'), true);
});

test('modèle en base64 dans les options (version précédente) : utilisé, puis rangé dans la table et retiré des options', async () => {
    const buffer = new Uint8Array([80, 75, 3, 4, 1, 2, 3]);
    const { state, saved, doc, fileOf, start } = controller({ sections: [section(7, { templateName: 'm.docx', templateHash: 'abc', templateData: Buffer.from(buffer).toString('base64') })] });
    start(); await tick(500);
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(new Uint8Array(state().templateBuffer), buffer);
    assert.deepEqual(saved(7), { templateName: 'm.docx', templateHash: 'abc', templateWidget: 7 });
    assert.deepEqual(fileOf(doc.modeles[0]), buffer);
});

test('modèle en base64 chez un lecteur : utilisé, rien n\'est écrit', async () => {
    const { state, events, start } = controller({ readonly: true, sections: [section(7, { templateName: 'm.docx', templateHash: 'abc', templateData: 'UEsDBA==' })] });
    start(); await tick();
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(writes(events), []);
});

test('ancien format (pièce jointe non rangée) : chargé, rangé dans la table et retiré des options', async () => {
    const { saved, state, doc, events, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })], attachments: { 42: [1, 2, 3] } });
    start(); await tick(500);
    assert.equal(state().templateName, 'ancien.docx');
    assert.equal(saved(7).templateId, undefined);
    assert.equal(saved(7).templateName, 'ancien.docx');
    assert.equal(doc.modeles[0].Modele, 'ancien.docx');
    assert.deepEqual([...doc.modeles[0].Fichier], ['L', 42], 'la pièce jointe existante est rangée telle quelle');
    assert.equal(events.some(e => /^téléversé/.test(e)), false, 'sans nouvel envoi');
});

test('ancien format chez un lecteur : téléchargé une seule fois malgré les notifications Grist', async () => {
    const { state, start, events } = controller({ readonly: true, sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })], attachments: { 42: [1] } });
    start(); start(); start();
    await tick();
    assert.deepEqual(events.filter(e => /^téléchargé/.test(e)), ['téléchargé:42']);
    assert.equal(state().templateName, 'ancien.docx');
    assert.deepEqual(writes(events), []);
});

test('ancien format dont la pièce jointe a été supprimée : message explicite', async () => {
    const { events, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    start();
    await tick();
    assert.match(toasts(events).at(-1), /« ancien.docx » n'est plus disponible/);
    assert.equal(events.includes('modale'), true);
});

test('un échec de lecture des tables liées bloque la génération au lieu de produire un document incomplet', async () => {
    const { app } = controller();
    app.addLinkedTables = async () => { throw new Error('Accès à Enfants indisponible'); };
    app.updateTemplateState(template(app, para('{Enfants.Nom}')), 'relations.docx', 'docx');
    await assert.rejects(app.dispatchGeneration({ id: 1, Titre: 'Parent' }, app.newBatch()), /Accès à Enfants indisponible/);
});

test('un aperçu terminé après un aperçu plus récent n’est pas affiché', async () => {
    const { app, events } = controller();
    vm.runInContext("state.currentRecord = { id: 1 }; state.templateType = 'docx';", app);
    const pending = [];
    app.dispatchGeneration = () => new Promise(resolve => pending.push(resolve));
    const ancien = app.updatePreview();
    const recent = app.updatePreview();
    pending[1]('récent');
    await recent;
    pending[0]('ancien');
    await ancien;
    assert.deepEqual(events.filter(e => e.aperçu).map(e => e.aperçu), ['récent']);
});

/* ── Opérations qui se croisent ─────────────────────────────────── */

test('modèle choisi pendant le chargement de l\'ancien : l\'ancien ne revient pas à l\'écran', async () => {
    // Éditeur d'un document à structure réservée : aucune notification ne viendrait corriger l'affichage
    const { app, events, state, doc, start } = controller({ downloadDelay: 60, structure: false,
        sections: [section(7, { templateName: 'ancien.docx', templateHash: 'h1', templateWidget: 7 })],
        modeles: [{ Widget: 7, Modele: 'ancien.docx', Empreinte: 'h1', Fichier: ['L', 2] }], attachments: { 2: [9] } });
    start(); await tick(10); // l'ancien modèle est en cours de téléchargement
    await app.handleTemplateUpload({ name: 'nouveau.docx', buffer: docx(app) });
    await tick(150);
    assert.deepEqual(events.filter(e => e.modèle).map(e => e.modèle), ['nouveau.docx']);
    assert.equal(state().templateName, 'nouveau.docx');
    assert.deepEqual(doc.modeles.map(r => r.Modele), ['nouveau.docx']);
});

test('modèle choisi pendant la reprise d\'un ancien format : une seule ligne, celle du nouveau', async () => {
    let slow = true;
    const { app, state, doc, saved, start } = controller({ modeles: [],
        sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })], attachments: { 42: [1, 2, 3] },
        // La reprise a lu la table (vide) et son écriture tarde
        beforeApply: async () => { if (slow) { slow = false; await tick(60); } } });
    start(); await tick(20);
    await app.handleTemplateUpload({ name: 'nouveau.docx', buffer: docx(app) });
    await tick(150);
    assert.equal(state().templateName, 'nouveau.docx');
    assert.deepEqual(doc.modeles.map(r => [r.Widget, r.Modele]), [[7, 'nouveau.docx']]);
    assert.deepEqual(Object.keys(saved(7)).sort(), ['templateHash', 'templateName', 'templateWidget']);
    assert.equal(saved(7).templateName, 'nouveau.docx');
});

test('table créée au même moment par un autre widget : l\'envoi réessaie et aboutit', async () => {
    let raced = false;
    const { app, events, doc, start } = controller({ beforeApply: (actions, { doc }) => {
        if (raced || actions[0][0] !== 'AddRawTable') return;
        raced = true; // l'autre widget vient de créer la table : cette action échoue en entier
        doc.columns = ['Widget', 'Modele', 'Empreinte', 'Fichier']; doc.modeles = [];
        throw new Error('KeyError: column A');
    } });
    start(); await tick();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.deepEqual(doc.modeles.map(r => [r.Widget, r.Modele]), [[7, 'a.docx']]);
    assert.equal(events.filter(e => /^téléversé/.test(e)).length, 1, 'le fichier n\'est envoyé qu\'une fois');
    assert.deepEqual(toasts(events).at(-1), 'success:Modèle enregistré dans le document');
});

test('options remises à zéro par Grist (widget re-sélectionné) : chaque widget retrouve son modèle', async () => {
    const { app, saved, state, events, start } = controller({ sections: [section(5, null), section(7, null), section(9, { templateWidget: 9 })],
        modeles: [{ Widget: 5, Modele: 'cinq.docx', Empreinte: 'h5', Fichier: ['L', 51] }, { Widget: 7, Modele: 'sept.docx', Empreinte: 'h7', Fichier: ['L', 71] }] });
    start(); await tick(600);
    assert.deepEqual(saved(5), { templateName: 'cinq.docx', templateHash: 'h5', templateWidget: 5 });
    assert.deepEqual(saved(7), { templateName: 'sept.docx', templateHash: 'h7', templateWidget: 7 });
    assert.deepEqual(saved(9), { templateWidget: 9 }, 'widget sans modèle : inchangé');
    assert.deepEqual(events.filter(e => /^téléchargé/.test(e)), ['téléchargé:71'], 'le widget lit son modèle, pas celui du voisin');
    assert.equal(events.includes('modale'), false);
});

test('options remises à zéro, lecteur : rien n\'est écrit ; après le passage d\'un éditeur, il retrouve son modèle', async () => {
    const modeles = [{ Widget: 5, Modele: 'cinq.docx', Empreinte: 'h5', Fichier: ['L', 51] }, { Widget: 7, Modele: 'sept.docx', Empreinte: 'h7', Fichier: ['L', 71] }];
    const attachments = { 71: [...docx(loadApp())] };
    const viewer = controller({ readonly: true, sections: [section(5, null), section(7, null)], modeles, attachments });
    viewer.start(); await tick(100);
    assert.equal(viewer.state().templateName, null);
    assert.deepEqual(writes(viewer.events), []);
    // Un éditeur ouvre l'autre widget (5) : les options des deux sont rétablies
    const editor = controller({ own: 5, sections: [section(5, null), section(7, null)], modeles, attachments });
    editor.start(); await tick(100);
    const after = controller({ readonly: true, modeles, attachments, sections: [5, 7].map(id => ({ id, tableRef: 1, parentKey: 'custom', options: editor.rows.get(id).options })) });
    after.start(); await tick(100);
    assert.equal(after.state().templateName, 'sept.docx');
});

test('colonne de la table supprimée : réparée à l\'ouverture du widget', async () => {
    const { doc, start } = controller({ modeles: [] });
    doc.columns = ['Widget', 'Modele', 'Fichier'];
    start(); await tick();
    assert.deepEqual([...doc.columns].sort(), ['Empreinte', 'Fichier', 'Modele', 'Widget']);
});

test('lecture des tables impossible à l\'ouverture : aucune erreur non traitée', async () => {
    const { app, events, start } = controller();
    const rejections = [];
    const onRejection = e => rejections.push(e);
    process.on('unhandledRejection', onRejection);
    app.grist.docApi.fetchTable = async () => { throw new Error('connexion perdue'); };
    start(); await tick(3300);
    process.off('unhandledRejection', onRejection);
    assert.deepEqual(rejections, []);
    assert.match(toasts(events).at(-1), /^error:Impossible de charger le modèle : connexion perdue/);
    assert.equal(events.includes('modale'), false, 'la lecture a échoué : rien ne dit que le modèle manque');
});

test('lecture interrompue à l\'ouverture : retentée, le modèle est chargé', async () => {
    const { app, state, events, start } = controller({ sections: [section(7, { templateName: 'sept.docx', templateHash: 'h7', templateWidget: 7 })],
        modeles: [{ Widget: 7, Modele: 'sept.docx', Empreinte: 'h7', Fichier: ['L', 71] }], attachments: { 71: [...docx(loadApp())] } });
    const fetchTable = app.grist.docApi.fetchTable;
    let failures = 2;
    app.grist.docApi.fetchTable = async t => {
        if (t === 'Publipostage_Modeles_NE_PAS_SUPPRIMER' && failures-- > 0) throw new Error('connexion perdue');
        return fetchTable(t);
    };
    start(); await tick(1300);
    assert.equal(state().templateName, 'sept.docx');
    assert.deepEqual(toasts(events), []);
});
