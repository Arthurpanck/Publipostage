const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, para, render, xmlText } = require('./helpers.cjs');

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
            AutreParent: [2, 1, 2, 0] },
    };
    const app = loadApp({ grist: {
        getSelectedTableId: async () => 'Parents',
        docApi: { fetchTable: async name => { calls.push(name); return tables[name]; } },
    } });
    return { app, calls, tables };
}

for (const list of [false, true]) {
    test(`liens ${list ? 'RefList' : 'Ref'} : seuls les enfants du parent sélectionné sont générés`, async () => {
        const { app, calls } = setup({ list });
        const buffer = template(app, para('{{Enfants.Nom}}'));
        const expected = list ? ['AliceCharlie', 'AliceBob', ''] : ['AliceCharlie', 'Bob', ''];
        for (const [i, parentId] of [1, 2, 99].entries()) {
            const data = await app.addChildTablesData({}, parentId, buffer);
            const zip = await render(app, buffer, data);
            assert.equal(xmlText(zip.file('word/document.xml').asText()), expected[i]);
        }
        assert.equal(calls.filter(name => name === 'Enfants').length, 1, 'un seul fetch par lot');
        app.clearRelationsCache();
        await app.addChildTablesData({}, 1, buffer);
        assert.equal(calls.filter(name => name === 'Enfants').length, 2);
    });
}

test('modèle sans relation : aucun accès aux métadonnées', async () => {
    const { app, calls } = setup();
    await app.addChildTablesData({ Nom: 'Parent' }, 1, template(app, para('{Nom}')));
    assert.equal(calls.length, 0);
});

test('boucle explicite et condition parent conservent leur comportement', async () => {
    const { app } = setup();
    const buffer = template(app, para('{#Enfants}{Nom};{/Enfants}') + para('{#Actif}Oui{/Actif}'));
    const data = await app.addChildTablesData({ Actif: true }, 1, buffer);
    assert.equal(xmlText((await render(app, buffer, data)).file('word/document.xml').asText()), 'Alice;Charlie;Oui');
    assert.equal(app.getRelationsWarnings().length, 0);
});

test('table sans lien : diagnostic explicite, champ parent conservé', async () => {
    const { app } = setup();
    const data = await app.addChildTablesData({ Titre: 'Parent' }, 1, template(app, para('{Lieux.Nom} {Titre}')));
    assert.equal(data.Titre, 'Parent');
    assert.match(app.getRelationsWarnings().join(' '), /lieux.*Parents/);
});

test('plusieurs colonnes de référence : avertir du lien choisi', async () => {
    const { app } = setup({ ambiguous: true });
    const data = await app.addChildTablesData({}, 1, template(app, para('{Enfants.Nom}')));
    assert.equal(data.enfants.map(row => row.nom).join(','), 'Alice,Charlie');
    assert.match(app.getRelationsWarnings().join(' '), /enfants.*Parent.*AutreParent/);
});

test('une lecture échouée peut être retentée sans recharger le widget', async () => {
    const { app } = setup();
    let attempts = 0;
    app.grist.docApi.fetchTable = async () => {
        if (++attempts === 1) throw new Error('Réseau indisponible');
        return { id: [1] };
    };
    await assert.rejects(app.fetchTableCached('Enfants'), /Réseau/);
    assert.equal((await app.fetchTableCached('Enfants')).id[0], 1);
    assert.equal(attempts, 2);
});
