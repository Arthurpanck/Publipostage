const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, batch, para, render, xmlText } = require('./helpers.cjs');

// Lignes reçues par onRecords : liaisons, filtres et tri déjà appliqués par Grist
const VIEW = [
    { id: 23, Service_commune_Prenom_Nom: 'Bob' },
    { id: 11, Service_commune_Prenom_Nom: 'Alice' },
];

// Table complète : la ligne 99 est exclue par le filtre du widget, et la
// colonne Telephone n'a pas été transmise au widget.
const TABLE = { id: [11, 23, 99], Service_commune_Prenom_Nom: ['Alice', 'Bob', 'Zoé'], Telephone: ['01', '02', '09'] };

function setup() {
    let tableIdCalls = 0;
    const app = loadApp({ grist: {
        getSelectedTableId: async () => { tableIdCalls++; return 'BDD_SERVICES'; },
        docApi: { fetchTable: async () => TABLE },
    } });
    return { app, tableIdCalls: () => tableIdCalls };
}

// Comme dispatchGeneration dans main.js, pour la ligne sélectionnée (Alice)
async function generate(app, buffer, viewRows = VIEW) {
    const record = viewRows.find(row => row.id === 11) || { id: 11, Service_commune_Prenom_Nom: 'Alice' };
    const lot = batch(viewRows);
    const data = await app.addLinkedTables(app.rowData(record), record, app.getReferencedTables(buffer), lot);
    const zip = await render(app, buffer, data);
    return { text: xmlText(zip.file('word/document.xml').asText()), notes: lot.notes };
}

test('la balise de la table courante suit les lignes du widget dans l’ordre Grist', async () => {
    const { app } = setup();
    const buf = template(app, para('Sélection : {Service_commune_Prenom_Nom}')
        + para('{BDD_SERVICES.Service_commune_Prenom_Nom};'));
    const { text, notes } = await generate(app, buf);
    assert.equal(text, 'Sélection : AliceBob;Alice;');
    assert.equal(notes.size, 0);
});

test('vue filtrée : tableau Word à doubles accolades et boucle manuelle', async () => {
    const { app } = setup();
    const buf = template(app, `<w:tbl><w:tr><w:tc>${para('{{BDD_SERVICES.Service_commune_Prenom_Nom}}')}</w:tc><w:tc>${para('Service')}</w:tc></w:tr></w:tbl>`
        + para('{#BDD_SERVICES}{Service_commune_Prenom_Nom};{/BDD_SERVICES}'));
    assert.equal((await generate(app, buf)).text, 'BobServiceAliceServiceBob;Alice;');
});

test('colonnes non transmises : complétées depuis la table, sans réintroduire les lignes filtrées', async () => {
    const { app } = setup();
    const buf = template(app, para('{BDD_SERVICES.Service_commune_Prenom_Nom} {BDD_SERVICES.Telephone};'));
    assert.equal((await generate(app, buf)).text, 'Bob 02;Alice 01;');
});

test('vue vide : la ligne de saisie n’est pas publipostée', async () => {
    const { app } = setup();
    const buf = template(app, para('{BDD_SERVICES.Service_commune_Prenom_Nom}')
        + para('{^BDD_SERVICES}Aucun service{/BDD_SERVICES}'));
    assert.equal((await generate(app, buf, [{ id: 'new', Service_commune_Prenom_Nom: '' }])).text, 'Aucun service');
});

test('aucune lecture des relations pour un modèle sans balise pointée', async () => {
    const { app, tableIdCalls } = setup();
    assert.equal((await generate(app, template(app, para('{Service_commune_Prenom_Nom}')))).text, 'Alice');
    assert.equal(tableIdCalls(), 0);
});
