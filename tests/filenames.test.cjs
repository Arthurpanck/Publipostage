const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { loadApp } = require('./helpers.cjs');

function controller() {
    const saved = [], files = [], messages = [];
    let onOptions;
    const options = { templateId: 42 };
    const app = loadApp({
        grist: { ready() {}, onRecord() {}, onRecords() {}, onOptions(fn) { onOptions = fn; },
            async setOption(key, value) { options[key] = value; },
            async getSelectedTableId() { return 'Donnees'; },
            docApi: { async fetchTable() { return { id: [1], Nom_cache: ['Caché'] }; } } },
        initUi() {}, setTimeout() {}, uiSetFilenameColumn() {},
        uiToast: (...args) => messages.push(args), saveAs: (blob, name) => saved.push(name),
        JSZip: class { file(name) { files.push(name); } async generateAsync() { return new Blob(); } },
    });
    vm.runInContext(fs.readFileSync(require.resolve('../main.js'), 'utf8'), app);
    vm.runInContext("state.templateBuffer = {}; state.templateType = 'docx';", app);
    app.dispatchGeneration = async () => new Blob();
    app.refreshWarnings = () => {};
    return { app, saved, files, messages, options, onOptions };
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

test('export individuel DOCX/PDF, colonne facultative et masquée', async () => {
    const { app, saved, onOptions } = controller();
    vm.runInContext("state.currentRecord = {id: 1, Titre: 'Séance'};", app);
    await app.downloadSingle();
    onOptions({ filenameColumn: 'TITRE' });
    await app.downloadSingle();
    vm.runInContext("state.templateType = 'pdf';", app);
    await app.downloadSingle();
    onOptions({ filenameColumn: 'nom_cache' });
    await app.downloadSingle();
    onOptions({ filenameColumn: 'colonne_supprimee' });
    await app.downloadSingle();
    assert.deepEqual(saved, ['Document_1.docx', 'Séance.docx', 'Séance.pdf', 'Caché.pdf', 'Document_1.pdf']);
});

test('ZIP : collisions de casse, de suffixe et de nettoyage, cellules vides', async () => {
    const { app, files, saved, onOptions } = controller();
    onOptions({ filenameColumn: 'Titre' });
    vm.runInContext(`state.allRecords = [
        {id:1,Titre:'Nom'}, {id:2,Titre:'nom'}, {id:3,Titre:'Nom (2)'},
        {id:4,Titre:'A/B'}, {id:5,Titre:'A:B'}, {id:6,Titre:''}, {id:'new'}
    ];`, app);
    await app.downloadBulk();
    assert.deepEqual(files, ['Nom.docx','nom (2).docx','Nom (2) (2).docx','A_B.docx','A_B (2).docx','Doc_6.docx']);
    assert.deepEqual(saved, ['Publipostage.zip']);
});

test('le choix persiste sans écraser le modèle et peut être désactivé', async () => {
    const { app, options } = controller();
    await app.saveFilenameColumn(' Titre ');
    assert.equal(options.filenameColumn, 'Titre');
    assert.equal(options.templateId, 42);
    await app.saveFilenameColumn('');
    assert.equal(options.filenameColumn, '');
});
