const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, render, xmlText, para, run } = require('./helpers.cjs');

const DATA = { Titre: 'Annuaire', BDD_SERVICES: [
    { Fonction: 'Direction territoriale', Nom: 'Alice' },
    { Fonction: 'Service social', Nom: 'Bob' },
    { Fonction: 'DIRECTION TERRITORIALE', Nom: 'Chloé' },
    { Fonction: ['Service social', 'Accueil'], Nom: 'Dan' },
] };
const cell = (text, span) => `<w:tc>${span ? '<w:tcPr><w:gridSpan w:val="2"/></w:tcPr>' : ''}${para(text)}</w:tc>`;
const row = (...cells) => `<w:tr>${cells.join('')}</w:tr>`;
const table = (band) => '<w:tbl>' + row(cell(band, true)) + row(cell('Fonction'), cell('Nom'))
    + row(cell('{{BDD_SERVICES.Fonction}}'), cell('{{BDD_SERVICES.Nom}}')) + '</w:tbl>';
const rowsOf = (xml) => xml.split('</w:tbl>')[0].split('<w:tr>').slice(1).map((r) => [...r.matchAll(/<w:t[^>]*>([^<]*)<\/w:t>/g)].map((m) => m[1]).join('|'));

async function generate(body, data = DATA) {
    const app = loadApp();
    const unknown = new Set();
    const xml = (await render(app, template(app, body), JSON.parse(JSON.stringify(data)), unknown)).file('word/document.xml').asText();
    return { xml, text: xmlText(xml), unknown };
}

test('bandeau de tableau : lignes filtrées, valeur affichée avec sa mise en forme', async () => {
    const band = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>{FILTRE: bdd_services.FONCTION == “Direction Territoriale”}</w:t></w:r></w:p>';
    const body = '<w:tbl>' + `<w:tr><w:tc>${band}</w:tc></w:tr>` + row(cell('{BDD_SERVICES.Nom}'), cell('{BDD_SERVICES.Fonction}')) + '</w:tbl>';
    const { xml, unknown } = await generate(body);
    assert.deepEqual(rowsOf(xml), ['Direction Territoriale', 'Alice|Direction territoriale', 'Chloé|DIRECTION TERRITORIALE']);
    assert.match(xml, /<w:b\/><\/w:rPr><w:t>Direction Territoriale<\/w:t>/);
    assert.equal(unknown.size, 0);
});

test('un filtre posé dans un tableau s’arrête avec le tableau', async () => {
    const { xml } = await generate(table('{filtre: BDD_SERVICES.Fonction == "Accueil"}') + para('{BDD_SERVICES.Nom};'));
    assert.deepEqual(rowsOf(xml), ['Accueil', 'Fonction|Nom', 'Service social,Accueil|Dan']);
    assert.match(xmlText(xml), /Alice;Bob;Chloé;Dan;$/);
});

test('puces et paragraphes : portée jusqu’à {fin filtre}, ligne de fermeture retirée', async () => {
    const bullet = (text) => `<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>${run(text)}</w:p>`;
    const body = para('{filtre: BDD_SERVICES.Fonction == "Service social"}') + bullet('{BDD_SERVICES.Nom}')
        + para('{fin filtre}') + para('Tous : {#BDD_SERVICES}{Nom} {/BDD_SERVICES}');
    const { xml, text } = await generate(body);
    assert.equal(text, 'Service socialBobDanTous : Alice Bob Chloé Dan ');
    assert.equal((xml.match(/<w:numPr>/g) || []).length, 2, 'une puce par ligne retenue');
    assert.equal((xml.match(/<w:p>/g) || []).length + (xml.match(/<w:p /g) || []).length, 4, 'paragraphe {fin filtre} retiré');
});

test('filtres successifs sur la même table : le second remplace le premier', async () => {
    const body = para('{filtre: BDD_SERVICES.Fonction == "Accueil"}') + para('{BDD_SERVICES.Nom};')
        + para('{filtre: BDD_SERVICES.Fonction == "Direction territoriale"}') + para('{BDD_SERVICES.Nom};');
    assert.equal((await generate(body)).text, 'AccueilDan;Direction territorialeAlice;Chloé;');
});

test('différent de : rien d’affiché, paragraphe du filtre retiré', async () => {
    const body = para('{filtre: BDD_SERVICES.Fonction != "Direction territoriale"}') + para('{BDD_SERVICES.Nom};') + para('{/filtre}');
    const { xml, text } = await generate(body);
    assert.equal(text, 'Bob;Dan;');
    assert.equal((xml.match(/<w:p>/g) || []).length, 2);
});

test('boucle manuelle, condition inversée et doubles accolades suivent le filtre', async () => {
    const body = para('{{filtre: BDD_SERVICES.Fonction == "Inconnue"}}') + para('{#BDD_SERVICES}{Nom};{/BDD_SERVICES}')
        + para('{^BDD_SERVICES}Aucun service{/BDD_SERVICES}') + para('{fin filtre}') + para('{^BDD_SERVICES}Jamais{/BDD_SERVICES}');
    assert.equal((await generate(body)).text, 'InconnueAucun service');
});

test('valeurs brutes de Grist : RefList/ChoiceList encodées ["L", …] et nombres', async () => {
    const data = { Liens: [{ Tags: ['L', 'Rouge', 'Vert'], N: 2 }, { Tags: ['L', 'Bleu'], N: 3 }] };
    const body = para('{filtre: Liens.Tags == vert}') + para('{Liens.N};') + para('{filtre: Liens.N == "3"}') + para('{Liens.N};');
    assert.equal((await generate(body, data)).text, 'vert2;33;');
});

test('colonne ou table inconnue dans un filtre : signalée dans la pastille', async () => {
    const body = para('{filtre: BDD_SERVICES.Fonctoin == "x"}') + para('{BDD_SERVICES.Nom};') + para('{filtre: Lieux.Nom == "x"}');
    const { unknown } = await generate(body);
    assert.deepEqual([...unknown].sort(), ['bdd_services.fonctoin', 'lieux.nom']);
});

test('filtre mal écrit : refusé dès l’import du modèle, avec la forme attendue', () => {
    const app = loadApp();
    const buffer = template(app, para('{filtre: BDD_SERVICES.Fonction = "x"}'));
    assert.throws(() => app.compileDocxTemplate(buffer), /Filtre illisible.*Forme attendue : \{filtre: Table\.Colonne == "valeur"\}/);
});
