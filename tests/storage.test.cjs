const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { loadApp, template, para } = require('./helpers.cjs');

function storage({ exists = false, wrongType = false, denied = false } = {}) {
    const rows = { id: [], Nom: [], Fichier: [] };
    const actions = [];
    const app = loadApp({ grist: { docApi: {
        async fetchTable(id) {
            if (id === '_grist_Tables') return { id: exists ? [8] : [], tableId: exists ? ['Publipostage_Modeles'] : [] };
            if (id === '_grist_Tables_column') return { id: [1,2], parentId: [8,8], colId: ['Nom','Fichier'], type: ['Text',wrongType ? 'Text' : 'Attachments'], isFormula: [false,false] };
            return rows;
        },
        async applyUserActions(batch) {
            if (denied) throw new Error('ACL_DENY');
            for (const a of batch) {
                actions.push(a);
                if (a[0] === 'AddTable') exists = true;
                else { rows.id.push(rows.id.length + 1); rows.Nom.push(a[3].Nom); rows.Fichier.push(a[3].Fichier); }
            }
            return { retValues: [{ table_id: 'Publipostage_Modeles' }] };
        },
    } } });
    vm.runInContext(fs.readFileSync(require.resolve('../template-storage.js'), 'utf8'), app);
    return { app, rows, actions };
}

test('DOCX et PDF restent référencés par de vraies cellules Attachments, sans doublon au nouvel appel', async () => {
    const { app, rows, actions } = storage();
    await app.retainTemplateAttachment(12, 'modele.docx');
    await app.retainTemplateAttachment(13, 'modele.pdf');
    await app.retainTemplateAttachment(12, 'modele.docx');
    assert.equal(actions.filter(a => a[0] === 'AddTable').length, 1);
    assert.equal(rows.id.length, 2);
    assert.deepEqual(JSON.parse(JSON.stringify(rows.Fichier)), [['L',12],['L',13]]);
});

test('une table homonyme incompatible ou un refus ACL ne produit pas une fausse réussite', async () => {
    await assert.rejects(storage({exists:true,wrongType:true}).app.retainTemplateAttachment(1,'a.docx'), /Attachments/);
    await assert.rejects(storage({denied:true}).app.retainTemplateAttachment(1,'a.docx'), /ACL_DENY/);
});

function controller() {
    const events = [];
    let onOptions;
    const app = loadApp({
        grist: { ready() {}, onRecord() {}, onRecords() {}, onOptions(fn) { onOptions = fn; } },
        initUi() {}, uiToast: (message,type) => events.push({message,type}),
        uiSetTemplateNotice: message => events.push({notice:message}), uiSetTemplate: name => events.push({name}),
        uiCloseModal() {}, uiOpenModal: () => events.push({modal:true}), uiEnableActions() {}, uiShowPreview() {}, uiPreviewEmptyText() {},
    });
    vm.runInContext(fs.readFileSync(require.resolve('../main.js'), 'utf8'), app);
    app.retainTemplateAttachment = async () => {};
    return { app, events, onOptions };
}

test('options initiales tardives : le modèle se charge sans délai fixe ni écriture au démarrage', async () => {
    const { app, events } = controller();
    app.downloadAttachmentFromGrist = async () => template(app, para('{A}'));
    await app.loadSavedTemplate({templateId:1,templateName:'persistant.docx'});
    assert.ok(events.some(e => e.name === 'persistant.docx'));
    assert.equal(events.some(e => e.modal), false);
});

test('un modèle inaccessible reste une erreur explicite et ne devient pas « aucun modèle »', async () => {
    const { app, events } = controller();
    app.downloadAttachmentFromGrist = async () => { throw new Error('HTTP 403'); };
    await app.loadSavedTemplate({templateId:1,templateName:'persistant.docx'});
    assert.ok(events.some(e => e.notice?.includes('HTTP 403')));
    assert.equal(events.some(e => e.modal), false);
});

test('un ancien téléchargement ne remplace pas le modèle choisi plus récemment', async () => {
    const { app, events } = controller();
    const pending = [];
    app.downloadAttachmentFromGrist = () => new Promise(resolve => pending.push(resolve));
    const old = app.loadSavedTemplate({templateId:1,templateName:'ancien.docx'});
    const recent = app.loadSavedTemplate({templateId:2,templateName:'recent.docx'});
    pending[1](template(app, para('{A}'))); await recent;
    pending[0](template(app, para('{A}'))); await old;
    assert.deepEqual(events.filter(e => e.name).map(e => e.name), ['recent.docx']);
});

test('suppression des options : le modèle précédent et ses actions sont désactivés', async () => {
    const { app } = controller();
    vm.runInContext("state.templateBuffer = new Uint8Array([1]); state.templateName = 'ancien.docx';", app);
    await app.loadSavedTemplate(null);
    assert.equal(vm.runInContext('state.templateBuffer', app), null);
});


test('un lecteur sans droit de migration peut toujours charger le modèle et reçoit une consigne explicite', async () => {
    const { app, events } = controller();
    app.downloadAttachmentFromGrist = async () => template(app, para('{A}'));
    app.retainTemplateAttachment = async () => { throw new Error('ACL_DENY'); };
    await app.loadSavedTemplate({templateId:1,templateName:'ancien.docx'});
    assert.ok(events.some(e => e.name === 'ancien.docx'));
    assert.ok(events.some(e => e.notice?.includes('ACL_DENY')));
});
