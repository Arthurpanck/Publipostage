const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const { loadApp, template, para, render, xmlText } = require('./helpers.cjs');

test('casse et accents : champs, conditions et variantes suffixées restent distincts', async () => {
    const app = loadApp();
    const buffer = template(app, para('{PRÉNOM}|{prenom}|{{PrEnOm}}|{FONCTION_2}|{#aCtIf}{NOM}{/ACTIF}|{^ViDe}Absent{/vide}'), para('{prénom}'), para('{PRENOM}'));
    const zip = await render(app, buffer, { Prenom: 'Élodie', Nom: 'DUPONT', Fonction_2: 'Adjointe', Actif: true, Vide: '' });
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Élodie|Élodie|Élodie|Adjointe|DUPONT|Absent');
    assert.equal(xmlText(zip.file('word/header1.xml').asText()), 'Élodie');
    assert.equal(xmlText(zip.file('word/footer1.xml').asText()), 'Élodie');
});

for (const mode of ['vue', 'reference', 'regroupement']) {
    test('tables et colonnes en casse mixte : ' + mode, async () => {
        const source = { id: [5], Service_commune_Fonction: ['Direction ÉDUCATION'], Commune: [1] };
        const tables = {
            _grist_Tables: { id: [10, 20], tableId: ['BDD_services', 'Communes'], summarySourceTable: [0, mode === 'regroupement' ? 10 : 0] },
            _grist_Tables_column: { id: [1], parentId: [10], colId: ['Commune'], type: ['Ref:Communes'] },
            BDD_services: source, Communes: { id: [1], group: [['L', 5]] },
        };
        const app = loadApp({ grist: {
            getSelectedTableId: async () => mode === 'vue' ? 'BDD_services' : 'Communes',
            fetchSelectedTable: async () => source,
            docApi: { fetchTable: async id => { assert.ok(tables[id], 'identifiant Grist réel conservé'); return tables[id]; } },
        } });
        const buffer = template(app, para('{BDD_SERVICES.service_COMMUNE_fonction}') + para('{#bDd_sErViCeS}{SERVICE_COMMUNE_FONCTION}{/BDD_services}'));
        const data = await app.addChildTablesData({}, 1, buffer);
        const zip = await render(app, buffer, data);
        assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Direction ÉDUCATIONDirection ÉDUCATION');
        assert.equal(app.getRelationsWarnings().length, 0);
    });
}

test('PDF : noms de champs et données suivent la même normalisation', async () => {
    let text, checked = false;
    class PDFTextField { getName() { return 'PRÉNOM'; } setText(value) { text = value; } }
    class PDFCheckBox { getName() { return 'ACTIF'; } check() { checked = true; } uncheck() { checked = false; } }
    const app = loadApp({ PDFLib: { PDFDocument: { load: async () => ({ getForm: () => ({ getFields: () => [new PDFTextField(), new PDFCheckBox()], flatten() {} }), save: async () => new Uint8Array() }) } } });
    vm.runInContext(fs.readFileSync('pdf-tools.js', 'utf8'), app);
    await app.generatePdfBlob({ Prenom: 'Élodie', Actif: true }, new Uint8Array());
    assert.equal(text, 'Élodie'); assert.equal(checked, true);
});
