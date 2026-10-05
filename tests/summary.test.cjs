const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, para, render, xmlText } = require('./helpers.cjs');

// Forme des réponses docApi.fetchTable : group est une RefList encodée.
// Le nom de la table récapitulative est volontairement arbitraire : seul
// summarySourceTable fait foi, pas un suffixe de nom ni les valeurs groupées.
function setup() {
    const tables = {
        _grist_Tables: { id: [10, 20], tableId: ['Donnees', 'Recap'], summarySourceTable: [0, 10] },
        _grist_Tables_column: { id: [1], parentId: [20], colId: ['group'], type: ['RefList:Donnees'] },
        Recap: { id: [1, 2, 3], Categorie: ['A', 'B', 'Vide'], count: [2, 2, 0], group: [['L', 13, 11], ['L', 12, 13], ['L']] },
        Donnees: { id: [11, 12, 13], Titre: ['Visite', 'Repas', 'Accueil'], Participants: [25, 38, 26] },
    };
    const calls = [];
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'Recap',
        docApi: { fetchTable: async name => { calls.push(name); return JSON.parse(JSON.stringify(tables[name])); } },
    } });
    return { app, tables, calls };
}

async function generate(app, buffer, id) {
    const data = {};
    await app.completeParentData(data, id);
    await app.addChildTablesData(data, id, buffer);
    return render(app, buffer, data);
}

test('regroupement : plusieurs groupes, ligne partagée et ordre donné par group', async () => {
    const { app, calls } = setup();
    const buffer = template(app, para('{Categorie} {count}:') + para('{{Donnees.Titre}};'));
    for (const [id, expected] of [[1, 'A 2:Accueil;Visite;'], [2, 'B 2:Repas;Accueil;'], [3, 'Vide 0:']]) {
        const zip = await generate(app, buffer, id);
        assert.equal(xmlText(zip.file('word/document.xml').asText()), expected);
        assert.equal(app.getRelationsWarnings().length, 0);
    }
    assert.equal(calls.filter(t => t === 'Donnees').length, 1, 'un seul fetch des données pour le lot');
});

test('group masqué : les détails se chargent même sans complément préalable du parent', async () => {
    const { app } = setup();
    const data = await app.addChildTablesData({}, 1, template(app, para('{Donnees.Titre}')));
    assert.equal(data.donnees.map(r => r.titre).join(','), 'Accueil,Visite');
});

test('regroupement : tableau Word, en-tête, pied de page et boucle manuelle', async () => {
    const { app } = setup();
    const body = `<w:tbl><w:tr><w:tc>${para('{Donnees.Titre}')}</w:tc><w:tc>${para('{{Donnees.Participants}}')}</w:tc></w:tr></w:tbl>`
        + para('{#Donnees}{Titre};{/Donnees}');
    const buffer = template(app, body, para('{Donnees.Titre};'), para('{Categorie}'));
    const zip = await generate(app, buffer, 1);
    const xml = zip.file('word/document.xml').asText();
    assert.equal((xml.match(/<w:tr>/g) || []).length, 2);
    assert.equal(xmlText(xml), 'Accueil26Visite25Accueil;Visite;');
    assert.equal(xmlText(zip.file('word/header1.xml').asText()), 'Accueil;Visite;');
    assert.equal(xmlText(zip.file('word/footer1.xml').asText()), 'A');
});

test('une condition inversée seule distingue un groupe rempli d’un groupe vide', async () => {
    const { app } = setup();
    const buffer = template(app, para('{^Donnees}Aucune ligne{/Donnees}'));
    assert.equal(xmlText((await generate(app, buffer, 1)).file('word/document.xml').asText()), '');
    assert.equal(xmlText((await generate(app, buffer, 3)).file('word/document.xml').asText()), 'Aucune ligne');
});

test('actualisation : changement de contenu et d’appartenance après invalidation du cache', async () => {
    const { app, tables } = setup();
    const buffer = template(app, para('{Donnees.Titre};'));
    await generate(app, buffer, 1);
    tables.Donnees.Titre[1] = 'Repas modifié';
    tables.Recap.group[0] = ['L', 12];
    app.clearRelationsCache();
    assert.equal(xmlText((await generate(app, buffer, 1)).file('word/document.xml').asText()), 'Repas modifié;');
});

for (const group of [undefined, ['E', 'Erreur'], ['L', 'invalide'], ['L', 999]]) {
    test('group indisponible ou référence invalide bloque la génération : ' + JSON.stringify(group), async () => {
        const { app, tables } = setup();
        tables.Recap.group[0] = group;
        await assert.rejects(generate(app, template(app, para('{Donnees.Titre}')), 1), /indisponible|invalides|introuvable/);
    });
}

test('une colonne group ordinaire ne suffit pas à traiter la table comme un regroupement', async () => {
    const { app, tables, calls } = setup();
    tables._grist_Tables.summarySourceTable[1] = 0;
    const data = await app.addChildTablesData({}, 1, template(app, para('{Donnees.Titre}')));
    assert.equal(data.donnees, undefined);
    assert.equal(calls.includes('Donnees'), false);
    assert.match(app.getRelationsWarnings().join(' '), /donnees.*Recap/);
});
