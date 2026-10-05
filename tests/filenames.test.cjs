const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { loadApp } = require('./helpers.cjs');

function controller() {
    const saved = [], files = [], messages = [];
    let onRecord, onRecords, ready, recordOptions, recordsOptions;
    const app = loadApp({
        grist: { ready(value) { ready = value; },
            onRecord(fn, options) { onRecord = fn; recordOptions = options; },
            onRecords(fn, options) { onRecords = fn; recordsOptions = options; } },
        initUi() {}, setTimeout() {}, uiEnableActions() {},
        uiToast: (...args) => messages.push(args), saveAs: (blob, name) => saved.push(name),
        JSZip: class { file(name) { files.push(name); } async generateAsync() { return new Blob(); } },
    });
    vm.runInContext(fs.readFileSync(require.resolve('../main.js'), 'utf8'), app);
    vm.runInContext("state.templateBuffer = {}; state.templateType = 'docx';", app);
    app.dispatchGeneration = async () => new Blob();
    app.refreshWarnings = () => {};
    app.updatePreview = async () => {};
    return { app, saved, files, messages, onRecord, onRecords, ready, recordOptions, recordsOptions };
}

test('noms portables : accents, chemins, extensions, réservés et valeurs vides', () => {
    const { buildExportFilename: name } = loadApp();
    assert.equal(name(' Réunion été.DOCX ', 'Doc_1', 'docx'), 'Réunion été.docx');
    assert.equal(name('../A/B:C?.pdf', 'Doc_1', 'pdf'), '_A_B_C_.pdf');
    assert.equal(name('CON', 'Doc_1', 'docx'), '_CON.docx');
    for (const value of [null, '', '  ', '...', ['E', 'error'], {}]) {
        assert.equal(name(value, 'Doc_1', 'pdf'), 'Doc_1.pdf');
    }
    assert.equal(name(0, 'Doc_1', 'docx'), '0.docx');
    assert.ok(name('é'.repeat(400), 'Doc_1', 'docx').length < 200);
});

test('mapping officiel facultatif : déclaration et conservation de toutes les colonnes', () => {
    const { ready, recordOptions, recordsOptions } = controller();
    assert.equal(ready.columns[0].name, 'filename');
    assert.equal(ready.columns[0].optional, true);
    assert.equal(recordOptions.includeColumns, 'normal');
    assert.equal(recordsOptions.includeColumns, 'normal');
});

test('export individuel DOCX/PDF : mapping, changement et retrait du choix', async () => {
    const { app, saved, onRecord } = controller();
    const row = {id: 1, Titre: 'Séance', Nom_cache: 'Caché'};
    await onRecord(row, null);
    await app.downloadSingle();
    await onRecord(row, {filename: 'Titre'});
    await app.downloadSingle();
    vm.runInContext("state.templateType = 'pdf';", app);
    await app.downloadSingle();
    await onRecord(row, {filename: 'Nom_cache'});
    await app.downloadSingle();
    await onRecord(row, {filename: null});
    await app.downloadSingle();
    assert.deepEqual(saved, ['Document_1.docx', 'Séance.docx', 'Séance.pdf', 'Caché.pdf', 'Document_1.pdf']);
});

test('ZIP : mapping onRecords, collisions, valeurs vides et retrait du choix', async () => {
    const { app, files, saved, onRecords } = controller();
    const rows = [
        {id:1,Titre:'Nom'}, {id:2,Titre:'nom'}, {id:3,Titre:'Nom (2)'},
        {id:4,Titre:'A/B'}, {id:5,Titre:'A:B'}, {id:6,Titre:''}, {id:'new'}
    ];
    await onRecords(rows, {filename: 'Titre'});
    await app.downloadBulk();
    assert.deepEqual(files, ['Nom.docx','nom (2).docx','Nom (2) (2).docx','A_B.docx','A_B (2).docx','Doc_6.docx']);
    assert.deepEqual(saved, ['Publipostage.zip']);
    files.length = 0;
    await onRecords([{id:1,Titre:'Nom'}], {filename: null});
    await app.downloadBulk();
    assert.deepEqual(files, ['Doc_1.docx']);
});
