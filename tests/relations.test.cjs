const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, batch, para, render, xmlText } = require('./helpers.cjs');

function setup({ list = false, ambiguous = false } = {}) {
    const calls = [];
    const tables = {
        _grist_Tables: { id: [1, 2], tableId: ['Parents', 'Enfants'] },
        _grist_Tables_column: {
            id: ambiguous ? [1, 2] : [1], parentId: ambiguous ? [2, 2] : [2],
            colId: ambiguous ? ['Parent', 'AutreParent'] : ['Parent'],
            type: ambiguous ? ['Ref:Parents', 'Ref:Parents'] : [list ? 'RefList:Parents' : 'Ref:Parents'],
        },
        Enfants: { id: [11, 12, 13, 14], Nom: ['Alice', 'Bob', 'Charlie', 'Sans parent'],
            Parent: list ? [['L', 1, 2], ['L', 2], ['L', 1], null] : [1, 2, 1, 0],
            AutreParent: [2, 1, 2, 0], manualSort: [1, 2, 3, 4] },
    };
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'Parents',
        docApi: { fetchTable: async name => { calls.push(name); return tables[name]; } },
    } });
    return { app, calls, tables };
}

// Comme dispatchGeneration dans main.js, pour une ligne du widget
async function generate(app, buffer, record, lot = batch()) {
    const data = await app.addLinkedTables(app.rowData(record), record, app.getReferencedTables(buffer), lot);
    return { data, text: xmlText((await render(app, buffer, data)).file('word/document.xml').asText()) };
}

for (const list of [false, true]) {
    test(`liens ${list ? 'RefList' : 'Ref'} : seuls les enfants du parent sélectionné sont générés`, async () => {
        const { app, calls } = setup({ list });
        const buffer = template(app, para('{{Enfants.Nom}}'));
        const expected = list ? ['AliceCharlie', 'AliceBob', ''] : ['AliceCharlie', 'Bob', ''];
        const lot = batch();
        for (const [i, id] of [1, 2, 99].entries()) {
            assert.equal((await generate(app, buffer, { id }, lot)).text, expected[i]);
        }
        assert.equal(calls.filter(name => name === 'Enfants').length, 1, 'un seul fetch par lot');
        await generate(app, buffer, { id: 1 });
        assert.equal(calls.filter(name => name === 'Enfants').length, 2, 'un nouveau lot relit la table');
    });
}

test('les colonnes techniques des tables liées ne sont pas publipostées', async () => {
    const { app } = setup();
    const { data } = await generate(app, template(app, para('{Enfants.Nom}')), { id: 1 });
    assert.deepEqual(Object.keys(data.Enfants[0]).sort(), ['AutreParent', 'Nom', 'Parent']);
});

test('modèle sans relation : aucun accès aux métadonnées', async () => {
    const { app, calls } = setup();
    await generate(app, template(app, para('{Nom}')), { id: 1, Nom: 'Parent' });
    assert.equal(calls.length, 0);
});

test('boucle explicite et condition parent conservent leur comportement', async () => {
    const { app } = setup();
    const lot = batch();
    const { text } = await generate(app, template(app, para('{#Enfants}{Nom};{/Enfants}') + para('{#Actif}Oui{/Actif}')), { id: 1, Actif: true }, lot);
    assert.equal(text, 'Alice;Charlie;Oui');
    assert.equal(lot.notes.size, 0);
});

test('table sans lien : diagnostic explicite, champ parent conservé', async () => {
    const { app } = setup();
    const lot = batch();
    const { data } = await generate(app, template(app, para('{Lieux.Nom} {Titre}')), { id: 1, Titre: 'Parent' }, lot);
    assert.equal(data.Titre, 'Parent');
    assert.match([...lot.notes].join(' '), /Lieux.*Parents/);
});

test('plusieurs colonnes de référence : avertir du lien choisi', async () => {
    const { app } = setup({ ambiguous: true });
    const lot = batch();
    const { data } = await generate(app, template(app, para('{Enfants.Nom}')), { id: 1 }, lot);
    assert.equal(data.Enfants.map(row => row.Nom).join(','), 'Alice,Charlie');
    assert.match([...lot.notes].join(' '), /Enfants.*Parent.*AutreParent/);
});

test('une lecture échouée fait échouer le lot ; le lot suivant relit la table', async () => {
    const { app, tables } = setup();
    let attempts = 0;
    app.grist.docApi.fetchTable = async name => {
        if (name === 'Enfants' && ++attempts === 1) throw new Error('Réseau indisponible');
        return tables[name];
    };
    const buffer = template(app, para('{Enfants.Nom}'));
    await assert.rejects(generate(app, buffer, { id: 1 }), /Impossible de lire la table liée "Enfants" : Réseau/);
    assert.equal((await generate(app, buffer, { id: 1 })).text, 'AliceCharlie');
    assert.equal(attempts, 2);
});

// Référence portée par la table du widget : Parents.Lieux -> Lieux
function setupForward(type) {
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'Parents',
        docApi: { fetchTable: async name => ({
            _grist_Tables: { id: [1, 3], tableId: ['Parents', 'Lieux'] },
            _grist_Tables_column: { id: [5], parentId: [1], colId: ['Lieux'], type: [type] },
            Lieux: { id: [7, 8, 9], Nom: ['Lyon', 'Bron', 'Vaulx'] },
        })[name] },
    } });
    return app;
}

test('référence portée par la ligne du widget : lignes désignées dans leur ordre', async () => {
    const app = setupForward('RefList:Lieux');
    const buffer = template(app, para('{Lieux.Nom};'));
    // Valeur décodée par Grist pour une RefList
    assert.equal((await generate(app, buffer, { id: 1, Lieux: [9, 7] })).text, 'Vaulx;Lyon;');
    assert.equal((await generate(app, buffer, { id: 1, Lieux: null })).text, '');
});

test('référence simple portée par la ligne du widget', async () => {
    const app = setupForward('Ref:Lieux');
    const buffer = template(app, para('{Lieux.Nom};'));
    assert.equal((await generate(app, buffer, { id: 1, Lieux: 8 })).text, 'Bron;');
    assert.equal((await generate(app, buffer, { id: 1, Lieux: 0 })).text, '');
});

test('référence à double sens : le lien de la table citée est utilisé sans alerte', async () => {
    const { app, tables } = setup();
    tables._grist_Tables_column = { id: [1, 2], parentId: [2, 1], colId: ['Parent', 'Enfants'],
        type: ['Ref:Parents', 'RefList:Enfants'] };
    const lot = batch();
    // L'ordre de la colonne du parent diffère : il ne doit pas être utilisé
    const { text } = await generate(app, template(app, para('{Enfants.Nom};')), { id: 1, Enfants: [13, 11] }, lot);
    assert.equal(text, 'Alice;Charlie;');
    assert.equal(lot.notes.size, 0);
});
