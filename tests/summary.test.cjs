const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, batch, para, render, xmlText } = require('./helpers.cjs');

// Une table regroupée n'est pas un cas particulier : sa colonne group est une
// RefList vers la table source, portée par la ligne du widget. Grist la
// transmet décodée (includeColumns: 'normal'), même masquée dans le widget.
function setup() {
    const tables = {
        _grist_Tables: { id: [10, 20], tableId: ['Donnees', 'Recap'] },
        _grist_Tables_column: { id: [1], parentId: [20], colId: ['group'], type: ['RefList:Donnees'] },
        Donnees: { id: [11, 12, 13], Titre: ['Visite', 'Repas', 'Accueil'], Participants: [25, 38, 26] },
    };
    const records = [
        { id: 1, Categorie: 'A', count: 2, group: [13, 11] },
        { id: 2, Categorie: 'B', count: 2, group: [12, 13] },
        { id: 3, Categorie: 'Vide', count: 0, group: [] },
    ];
    const calls = [];
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'Recap',
        docApi: { fetchTable: async name => { calls.push(name); return JSON.parse(JSON.stringify(tables[name])); } },
    } });
    return { app, tables, records, calls };
}

// Comme dispatchGeneration dans main.js, pour une ligne du widget
async function generate(app, buffer, record, lot = batch()) {
    const data = app.rowData(record);
    await app.addLinkedTables(data, record, app.getReferencedTables(buffer), lot);
    return render(app, buffer, data);
}

test('regroupement : plusieurs groupes, ligne partagée et ordre donné par group', async () => {
    const { app, records, calls } = setup();
    const buffer = template(app, para('{Categorie} {count}:') + para('{{Donnees.Titre}};'));
    const lot = batch(records);
    for (const [record, expected] of [[records[0], 'A 2:Accueil;Visite;'], [records[1], 'B 2:Repas;Accueil;'], [records[2], 'Vide 0:']]) {
        const zip = await generate(app, buffer, record, lot);
        assert.equal(xmlText(zip.file('word/document.xml').asText()), expected);
    }
    assert.equal(lot.notes.size, 0);
    assert.equal(calls.filter(t => t === 'Donnees').length, 1, 'un seul fetch des données pour le lot');
    assert.equal(calls.includes('Recap'), false, 'la table regroupée n’est pas relue');
});

test('regroupement : tableau Word, en-tête, pied de page et boucle manuelle', async () => {
    const { app, records } = setup();
    const body = `<w:tbl><w:tr><w:tc>${para('{Donnees.Titre}')}</w:tc><w:tc>${para('{{Donnees.Participants}}')}</w:tc></w:tr></w:tbl>`
        + para('{#Donnees}{Titre};{/Donnees}');
    const buffer = template(app, body, para('{Donnees.Titre};'), para('{Categorie}'));
    const zip = await generate(app, buffer, records[0]);
    const xml = zip.file('word/document.xml').asText();
    assert.equal((xml.match(/<w:tr>/g) || []).length, 2);
    assert.equal(xmlText(xml), 'Accueil26Visite25Accueil;Visite;');
    assert.equal(xmlText(zip.file('word/header1.xml').asText()), 'Accueil;Visite;');
    assert.equal(xmlText(zip.file('word/footer1.xml').asText()), 'A');
});

test('une condition inversée seule distingue un groupe rempli d’un groupe vide', async () => {
    const { app, records } = setup();
    const buffer = template(app, para('{^Donnees}Aucune ligne{/Donnees}'));
    assert.equal(xmlText((await generate(app, buffer, records[0])).file('word/document.xml').asText()), '');
    assert.equal(xmlText((await generate(app, buffer, records[2])).file('word/document.xml').asText()), 'Aucune ligne');
});

test('actualisation : un nouveau lot voit le contenu et l’appartenance modifiés', async () => {
    const { app, tables, records } = setup();
    const buffer = template(app, para('{Donnees.Titre};'));
    await generate(app, buffer, records[0]);
    tables.Donnees.Titre[1] = 'Repas modifié';
    const zip = await generate(app, buffer, { ...records[0], group: [12] });
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Repas modifié;');
});

// undefined : colonne non transmise ; objet : erreur de formule décodée par Grist
for (const group of [undefined, { error: 'Erreur' }, ['invalide'], [999]]) {
    test('group indisponible ou référence invalide bloque la génération : ' + JSON.stringify(group), async () => {
        const { app, records } = setup();
        await assert.rejects(generate(app, template(app, para('{Donnees.Titre}')), { ...records[0], group }),
            /transmise|invalides|introuvable/);
    });
}
