const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { loadApp, template, para } = require('./helpers.cjs');

function controller() {
    let options = { templateId: 42, templateName: 'ancien.docx', other: 'conserver' };
    const events = [];
    const app = loadApp({
        grist: { ready() {}, onRecord() {}, onRecords() {}, getOptions: async () => options,
            setOptions: async value => { events.push('save'); options = value; },
            getOption: async key => options[key], setOption: async (key, value) => { events.push('save'); options[key] = value; } },
        setTimeout() {}, initUi() {}, uiToast: (message, type) => events.push({message, type}),
        uiCloseModal() {}, uiSetTemplate() {}, uiEnableActions() {}, uiShowPreview() {}, uiPreviewEmptyText() {},
        PDFLib: { PDFDocument: { load: async () => { throw new Error('PDF invalide'); } } },
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8'), app);
    app.readFileAsBuffer = async file => file.buffer;
    app.uploadAttachmentToGrist = async () => { events.push('upload'); return 99; };
    return { app, events, options: () => options };
}
test('un DOCX invalide est refusé avant upload et conserve le modèle précédent', async () => {
    const { app, events, options } = controller();
    await app.handleTemplateUpload({ name: 'invalide.docx', buffer: template(app, para('{Titre')) });
    assert.equal(events.includes('upload'), false);
    assert.equal(events.includes('save'), false);
    assert.equal(options().templateId, 42);
    assert.equal(events.at(-1).type, 'error');
});
test('un PDF invalide est refusé avant upload', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'invalide.pdf', buffer: new Uint8Array() });
    assert.equal(events.includes('upload'), false);
    assert.equal(events.at(-1).type, 'error');
});
test('extension DOCX en majuscules et sauvegarde atomique des options', async () => {
    const { app, events, options } = controller();
    await app.handleTemplateUpload({ name: 'MODELE.DOCX', buffer: template(app, para('{{Titre}}')) });
    assert.deepEqual(events.filter(e => typeof e === 'string'), ['upload', 'save']);
    assert.equal(options().templateId, 99);
    assert.equal(options().templateName, 'MODELE.DOCX');
    assert.equal(options().other, 'conserver');
    assert.equal(app.getTemplateType('MODELE.PDF'), 'pdf');
});
test('le format binaire .doc est refusé avec une consigne de conversion', async () => {
    const { app, events } = controller();
    await app.handleTemplateUpload({ name: 'ancien.doc', buffer: new Uint8Array() });
    assert.equal(events.includes('upload'), false);
    assert.match(events.at(-1).message, /Enregistrez le fichier en .docx/);
});
