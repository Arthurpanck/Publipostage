const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { loadApp, template, para } = require('./helpers.cjs');

const WIDGET = 'https://widget.test/publipostage/index.html';
const section = (id, widgetOptions, { url = WIDGET, tableRef = 1 } = {}) =>
    ({ id, tableRef, parentKey: 'custom', options: JSON.stringify({ customView: JSON.stringify({ mode: 'url', url, widgetOptions }) }) });

// Contrôleur réel (main.js) dans un document Grist simulé : les options
// enregistrées vivent dans _grist_Views_section, et Grist renvoie au widget
// (onOptions) celles de sa propre section quand elles changent.
function controller({ sections = [section(7, null)], own = 7, readonly = false, denied = false } = {}) {
    if (denied === true) denied = 'No write access';
    const events = [];
    const rows = new Map(sections.map(s => [s.id, { ...s }]));
    const widgetOptions = id => JSON.parse(JSON.parse(rows.get(id).options).customView).widgetOptions;
    let onOptions = () => {};
    let unsaved = null;
    const app = loadApp({
        location: new URL(`${WIDGET}?access=full&readonly=${readonly}`), setTimeout, clearInterval, setInterval: () => 0,
        grist: { ready() {}, onRecord() {}, onRecords() {}, onOptions(fn) { onOptions = fn; },
            getOptions: async () => unsaved || widgetOptions(own), setOptions: async value => { events.push('setOptions'); unsaved = value; },
            getSelectedTableId: async () => 'Donnees',
            docApi: {
                fetchTable: async t => {
                    if (t === '_grist_Tables') return { id: [1, 2], tableId: ['Donnees', 'Autre'] };
                    const list = [...rows.values()];
                    return { id: list.map(r => r.id), tableRef: list.map(r => r.tableRef), parentKey: list.map(r => r.parentKey), options: list.map(r => r.options) };
                },
                applyUserActions: async actions => {
                    if (denied) throw new Error(denied);
                    for (const [, table, id, values] of actions) {
                        assert.equal(table, '_grist_Views_section');
                        rows.get(id).options = values.options;
                        events.push(`écrit:${id}`);
                        if (id === own) setTimeout(() => onOptions(widgetOptions(own)), 5);
                    }
                },
            } },
        initUi() {}, uiToast: (message, type) => events.push({ message, type }),
        uiOpenModal() { events.push('modale'); }, uiCloseModal() {}, uiSetTemplate() {}, uiEnableActions() {}, uiShowPreview() {},
        uiPreviewEmptyText() {}, uiSetWarnings() {}, uiPersistBanner: text => events.push({ bandeau: text }),
        PDFLib: { PDFDocument: { load: async () => { throw new Error('PDF invalide'); } } },
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8'), app);
    app.readFileAsBuffer = async file => file.buffer;
    const start = () => onOptions(widgetOptions(own));
    return { app, events, saved: widgetOptions, rows, start, state: () => vm.runInContext('state', app) };
}
const docx = app => template(app, para('{{Titre}}'));
const toasts = events => events.filter(e => e.message).map(e => `${e.type}:${e.message}`);

test('un DOCX invalide est refusé et le modèle enregistré est conservé', async () => {
    const { app, events, saved } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    await app.handleTemplateUpload({ name: 'invalide.docx', buffer: template(app, para('{Titre')) });
    assert.equal(events.some(e => typeof e === 'string' && e.startsWith('écrit')), false);
    assert.deepEqual(saved(7), { templateId: 42, templateName: 'ancien.docx' });
    assert.equal(events.at(-1).type, 'error');
});
test('un PDF invalide est refusé', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'invalide.pdf', buffer: new Uint8Array() });
    assert.equal(events.includes('écrit:7'), false);
    assert.equal(events.at(-1).type, 'error');
});
test('le format binaire .doc est refusé avec une consigne de conversion', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'ancien.doc', buffer: new Uint8Array() });
    assert.match(events.at(-1).message, /Enregistrez le fichier en .docx/);
});
test('un modèle de plus de 8 Mo est refusé', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'lourd.docx', buffer: new Uint8Array(8 * 1024 * 1024 + 1) });
    assert.equal(events.includes('écrit:7'), false);
    assert.match(events.at(-1).message, /trop lourd/);
});

test('le modèle est écrit dans le document sans passer par le bouton « Enregistrer »', async () => {
    const { app, events, saved, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx', autre: 'conserver' })] });
    start();
    const buffer = docx(app);
    await app.handleTemplateUpload({ name: 'MODELE.DOCX', buffer });
    const options = saved(7);
    assert.equal(options.templateName, 'MODELE.DOCX');
    assert.equal(options.autre, 'conserver');
    assert.equal(options.templateId, undefined, "plus de pièce jointe : rien que Grist puisse supprimer");
    assert.deepEqual(new Uint8Array(app.base64ToBuffer(options.templateData)), buffer);
    assert.match(options.templateHash, /^[0-9a-f]{24}$/);
    assert.equal(events.includes('setOptions'), false);
    assert.deepEqual(toasts(events).at(-1), 'success:Modèle enregistré dans le document');
    assert.equal(app.getTemplateType('MODELE.PDF'), 'pdf');
});

test('deux widgets neufs sur la même table : seul le sien reçoit le modèle', async () => {
    const autre = section(5, null), autreTable = section(6, null, { tableRef: 2 }), autreUrl = section(8, null, { url: 'https://ailleurs.test/w.html' });
    const { app, saved, rows, start, events } = controller({ sections: [autre, autreTable, autreUrl, section(7, null)] });
    start();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(saved(7).templateName, 'a.docx');
    assert.equal(rows.get(5).options, autre.options, "l'autre widget, sondé puis rétabli, retrouve ses options");
    assert.equal(saved(6), null);
    assert.equal(saved(8), null);
    assert.equal(events.includes('setOptions'), false);
});

test('les options enregistrées départagent deux widgets de la même table sans sonde', async () => {
    const { app, saved, events, start } = controller({ sections: [section(5, { templateName: 'x.docx', templateData: 'AA==' }), section(7, null)] });
    start();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(saved(7).templateName, 'a.docx');
    assert.deepEqual(events.filter(e => typeof e === 'string' && e.startsWith('écrit')), ['écrit:7']);
});

test('lecture seule : rien n\'est écrit, le modèle sert pour la session', async () => {
    const { app, events, state, start } = controller({ readonly: true });
    start();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(state().templateName, 'a.docx');
    assert.equal(events.some(e => e === 'setOptions' || /^écrit/.test(e)), false);
    assert.match(toasts(events).at(-1), /^error:Lecture seule/);
});

test('écriture refusée : options proposées à Grist et bandeau « Enregistrer »', async () => {
    const { app, events, start } = controller({ denied: true });
    start();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(events.includes('setOptions'), true);
    assert.match(events.find(e => e.bandeau).bandeau, /cliquez sur « Enregistrer »/);
});

test('structure réservée aux propriétaires : pas de bouton « Enregistrer » voué à l\'échec', async () => {
    const { app, events, state, start } = controller({ denied: 'Blocked by table structure access rules' });
    start();
    await app.handleTemplateUpload({ name: 'a.docx', buffer: docx(app) });
    assert.equal(state().templateName, 'a.docx');
    assert.equal(events.includes('setOptions'), false);
    assert.match(events.find(e => e.bandeau).bandeau, /réserve la modification de sa structure aux propriétaires/);
});

test('au chargement, le modèle intégré aux options est utilisé tel quel', async () => {
    const buffer = new Uint8Array([80, 75, 3, 4, 1, 2, 3]);
    const { app, state } = controller({ sections: [section(7, { templateName: 'm.docx', templateHash: 'abc', templateData: Buffer.from(buffer).toString('base64') })] });
    await app.applyOptions(await app.grist.getOptions(), true);
    assert.equal(state().templateName, 'm.docx');
    assert.deepEqual(new Uint8Array(state().templateBuffer), buffer);
});

test('ancien format (pièce jointe) : chargé puis intégré aux options', async () => {
    const { app, saved, state, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    const buffer = new Uint8Array([1, 2, 3]).buffer;
    app.downloadAttachmentFromGrist = async id => { assert.equal(id, 42); return buffer; };
    start();
    await new Promise(r => setTimeout(r, 50));
    assert.equal(state().templateName, 'ancien.docx');
    assert.equal(saved(7).templateId, undefined);
    assert.equal(saved(7).templateData, 'AQID');
});

test('ancien format chez un lecteur : téléchargé une seule fois malgré les notifications Grist', async () => {
    const { app, state, start, events } = controller({ readonly: true, sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    let downloads = 0;
    app.downloadAttachmentFromGrist = async () => { downloads++; return new Uint8Array([1]).buffer; };
    start(); start(); start();
    await new Promise(r => setTimeout(r, 30));
    assert.equal(downloads, 1);
    assert.equal(state().templateName, 'ancien.docx');
    assert.equal(events.some(e => /^écrit/.test(e)), false);
});

test('ancien format dont la pièce jointe a été supprimée : message explicite', async () => {
    const { app, events, start } = controller({ sections: [section(7, { templateId: 42, templateName: 'ancien.docx' })] });
    app.downloadAttachmentFromGrist = async () => { throw new Error('404'); };
    start();
    await new Promise(r => setTimeout(r, 20));
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
    const { app } = controller();
    const container = { replaceChildren(...nodes) { this.nodes = nodes; } };
    app.document = { getElementById: () => container, createElement: () => ({ childNodes: [] }) };
    app.docx = { renderAsync: async (blob, page) => { page.childNodes = [blob]; } };
    vm.runInContext("state.currentRecord = { id: 1 }; state.templateType = 'docx';", app);
    const pending = [];
    app.dispatchGeneration = () => new Promise(resolve => pending.push(resolve));
    const ancien = app.updatePreview();
    const recent = app.updatePreview();
    pending[1]('récent');
    await recent;
    pending[0]('ancien');
    await ancien;
    assert.deepEqual(container.nodes, ['récent']);
});
