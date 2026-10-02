const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, para, render, xmlText } = require('./helpers.cjs');

function setup() {
    let view = { id: [23, 11], Service_commune_Prenom_Nom: ['Bob', 'Alice'] };
    let attempts = 0;
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'BDD_SERVICES',
        fetchSelectedTable: async options => {
            assert.equal(options.includeColumns, 'normal');
            assert.equal(options.format, 'columns');
            attempts++;
            return JSON.parse(JSON.stringify(view));
        },
        docApi: { fetchTable: async () => { throw new Error('La table entière ne doit pas être lue'); } },
    } });
    return { app, attempts: () => attempts, setView: value => { view = value; } };
}

async function generate(app, buffer) {
    const data = await app.addChildTablesData({ Service_commune_Prenom_Nom: 'Alice' }, 11, buffer);
    const zip = await render(app, buffer, data);
    return xmlText(zip.file('word/document.xml').asText());
}

test('la balise de la table courante suit les lignes filtrées dans l’ordre Grist', async () => {
    const { app } = setup();
    const buf = template(app, para('Sélection : {Service_commune_Prenom_Nom}')
        + para('{BDD_SERVICES.Service_commune_Prenom_Nom};'));
    assert.equal(await generate(app, buf), 'Sélection : AliceBob;Alice;');
    assert.equal(app.getRelationsWarnings().length, 0);
});

test('vue filtrée : tableau Word à doubles accolades et boucle manuelle', async () => {
    const { app } = setup();
    const buf = template(app, `<w:tbl><w:tr><w:tc>${para('{{BDD_SERVICES.Service_commune_Prenom_Nom}}')}</w:tc><w:tc>${para('Service')}</w:tc></w:tr></w:tbl>`
        + para('{#BDD_SERVICES}{Service_commune_Prenom_Nom};{/BDD_SERVICES}'));
    assert.equal(await generate(app, buf), 'BobServiceAliceServiceBob;Alice;');
});

test('le cache de la vue est réutilisé dans un lot puis invalidé pour la nouvelle liaison', async () => {
    const { app, attempts, setView } = setup();
    const buf = template(app, para('{BDD_SERVICES.Service_commune_Prenom_Nom};'));
    await generate(app, buf);
    await generate(app, buf);
    assert.equal(attempts(), 1);
    setView({ id: [31], Service_commune_Prenom_Nom: ['Claire'] });
    app.clearRelationsCache();
    assert.equal(await generate(app, buf), 'Claire;');
    assert.equal(attempts(), 2);
});

test('vue vide : aucun ancien service ni ligne de saisie n’est conservé', async () => {
    const { app, setView } = setup();
    setView({ id: ['new'], Service_commune_Prenom_Nom: [''] });
    const buf = template(app, para('{BDD_SERVICES.Service_commune_Prenom_Nom}')
        + para('{^BDD_SERVICES}Aucun service{/BDD_SERVICES}'));
    assert.equal(await generate(app, buf), 'Aucun service');
});

test('un échec de lecture de la vue ne provoque pas de repli sur toute la table', async () => {
    const { app } = setup();
    app.grist.fetchSelectedTable = async () => { throw new Error('Vue indisponible'); };
    await assert.rejects(generate(app, template(app, para('{BDD_SERVICES.Service_commune_Prenom_Nom}'))), /Vue indisponible/);
});

test('la vue n’est pas relue pour un modèle contenant seulement une balise simple', async () => {
    const { app, attempts } = setup();
    assert.equal(await generate(app, template(app, para('{Service_commune_Prenom_Nom}'))), 'Alice');
    assert.equal(attempts(), 0);
});
