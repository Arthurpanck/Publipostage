const { test } = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { loadApp, template, para, xmlText } = require('./helpers.cjs');

// Table Donnees telle que lue par fetchTable (RefList encodée)
const DONNEES = { id: [1, 2], manualSort: [1, 2], Titre: ['Visite', 'Rendez-vous A'], Participants: [25, 14],
    Lieu: [['L', 1, 3], ['L', 2]], Lieu_texte: ['Hall, Réfectoire', ''], Debut_texte: ['02/10 à 9h15', '08/09 à 16h45'] };

const TABLES = { Donnees: DONNEES,
    _grist_Tables: { id: [1, 2], tableId: ['Donnees', 'Lieux'] },
    _grist_Tables_column: { id: [1], parentId: [1], colId: ['Lieu'], type: ['RefList:Lieux'] } };

// Contrôleur réel (main.js) avec un Grist simulé
function controller(fetchTable = async (name) => TABLES[name]) {
    const app = loadApp({
        grist: { ready() {}, onRecord() {}, onRecords() {}, getSelectedTableId: async () => 'Donnees', docApi: { fetchTable } },
        setTimeout() {}, initUi() {}, uiToast() {}, uiSetWarnings() {}, uiEnableActions() {}, uiSetTemplate() {},
        uiShowPreview() {}, uiPreviewEmptyText() {}, uiCloseModal() {}, uiOpenModal() {},
    });
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../main.js'), 'utf8'), app);
    return app;
}

async function generate(app, body, record, viewRows) {
    app.__buffer = template(app, body);
    app.__rows = viewRows;
    vm.runInContext("state.templateBuffer = __buffer; state.templateType = 'docx'; state.allRecords = __rows;", app);
    const batch = app.newBatch();
    const blob = await app.dispatchGeneration(record, batch);
    const text = xmlText(new app.PizZip(await blob.arrayBuffer()).file('word/document.xml').asText());
    return { text, unknown: [...batch.unknownTags], notes: [...batch.notes] };
}

// Extrait du modèle test-publipostage-donnees-lieux.docx
const BODY = para('{Titre} | {{PARTICIPANTS}} | {Début_texte}')
    + para('{DONNEES.Titre};') + para('{{#DONNEES}}') + para('{TITRE} : {#LIEU_TEXTE}{lieu_texte}{/Lieu_texte}{^LIEU_TEXTE}aucun lieu{/lieu_texte}.') + para('{{/donnees}}');

test('Grist transmet les lignes sans leurs colonnes : elles sont complétées depuis la table', async () => {
    const { text, unknown, notes } = await generate(controller(), BODY, { id: 1 }, [{ id: 1 }, { id: 2 }]);
    assert.equal(text, 'Visite | 25 | 02/10 à 9h15Visite;Rendez-vous A;Visite : Hall, Réfectoire.Rendez-vous A : aucun lieu.');
    assert.deepEqual(unknown, []);
    assert.deepEqual(notes, []);
});

test('les valeurs reçues de Grist restent prioritaires sur celles relues', async () => {
    const { text } = await generate(controller(), para('{Titre}'), { id: 1, Titre: 'Reçu' }, [{ id: 1 }]);
    assert.equal(text, 'Reçu');
});

test('les listes relues sont décodées comme celles reçues par le widget', async () => {
    const app = controller();
    const record = await app.completeRow({ id: 1 }, app.newBatch());
    assert.deepEqual(record.Lieu, [1, 3]);
    assert.equal('manualSort' in app.rowData(record), false);
});

test('lecture de la table impossible : la ligne reçue est utilisée telle quelle', async () => {
    const app = controller(async () => { throw new Error('Accès refusé'); });
    const { text } = await generate(app, para('{Titre}'), { id: 1, Titre: 'Reçu' }, [{ id: 1, Titre: 'Reçu' }]);
    assert.equal(text, 'Reçu');
});
