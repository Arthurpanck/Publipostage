const { test } = require('node:test');
const assert = require('node:assert/strict');
const { loadApp, template, render, xmlText, para, run } = require('./helpers.cjs');

test('balises simples, accents, valeurs nulles, zéro et faux', async () => {
    const app = loadApp();
    const unknown = new Set();
    const zip = await render(app, template(app, para('{Titre}|{Prénom}|{Vide}|{Zero}|{Faux}')), { Titre: 'Visite', Prenom: 'Élodie', Vide: null, Zero: 0, Faux: false }, unknown);
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Visite|Élodie||0|false');
    assert.equal(unknown.size, 0);
});
test('doubles accolades et balises simples peuvent coexister', async () => {
    const app = loadApp();
    const zip = await render(app, template(app, para('{{Titre}} {Participants} {{Prénom}}')), { Titre: 'Visite', Participants: 25, Prenom: 'Élodie' });
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Visite 25 Élodie');
});
test('les runs sans balises gardent exactement leur mise en forme', () => {
    const app = loadApp();
    const xml = '<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>GRAS</w:t></w:r>' + run(' normal') + '</w:p>';
    assert.equal(app.repairDocxXml(xml), xml);
});
test('balise répartie sur des runs sans déplacer le texte voisin', async () => {
    const app = loadApp();
    const body = '<w:p>' + run('Avant {{Ti') + '<w:r><w:rPr><w:b/></w:rPr><w:t>tre}} après</w:t></w:r></w:p>';
    const zip = await render(app, template(app, body), { Titre: 'Visite' });
    const xml = zip.file('word/document.xml').asText();
    assert.equal(xmlText(xml), 'Avant Visite après');
    assert.match(xml, /<w:b\s*\/>[\s\S]*?<w:t[^>]*> après<\/w:t>/);
});
test('en-têtes et pieds de page suivent les mêmes règles', async () => {
    const app = loadApp();
    const unknown = new Set();
    const zip = await render(app, template(app, para('{Prénom}'), para('{{Prénom}}'), para('{Prénom}')), { Prenom: 'Élodie' }, unknown);
    for (const name of ['document', 'header1', 'footer1']) assert.equal(xmlText(zip.file('word/' + name + '.xml').asText()), 'Élodie');
    assert.equal(unknown.size, 0);
});
for (const text of ['{{Titre}', '{Titre}}', '{{{Titre}}}', '{Titre', 'Titre}']) {
    test('rejette les accolades mal appariées : ' + text, () => {
        const app = loadApp();
        assert.throws(() => app.generateDocxBlob({ Titre: 'Visite' }, template(app, para(text))));
    });
}
test('une balise incomplète ne peut pas avaler le XML du paragraphe suivant', () => {
    const app = loadApp();
    const buffer = template(app, para('Texte {incomplet') + para('Suite {Titre}'));
    assert.throws(() => app.generateDocxBlob({ Titre: 'Visite' }, buffer));
});
test('boucles relationnelles simples et doubles dans un tableau', async () => {
    for (const double of [false, true]) {
        const app = loadApp();
        const tag = double ? '{{Enfants.Nom}}' : '{Enfants.Nom}';
        const body = `<w:tbl><w:tr><w:tc>${para(tag)}</w:tc><w:tc>${para('{Titre}')}</w:tc></w:tr></w:tbl>`;
        const zip = await render(app, template(app, body), { Titre: 'Parent', Enfants: [{Nom:'Alice'}, {Nom:'Bob'}] });
        const xml = zip.file('word/document.xml').asText();
        assert.equal((xml.match(/<w:tr>/g) || []).length, 2);
        assert.equal(xmlText(xml), 'AliceParentBobParent');
    }
});
test('boucles manuelles, conditions inversées et tableaux sans enfant', async () => {
    const app = loadApp();
    const zip = await render(app, template(app, para('{{#Enfants}}') + para('{{Nom}}') + para('{{/Enfants}}') + para('{^Enfants}Aucun{/Enfants}')), { Enfants: [] });
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Aucun');
});
test('les balises inconnues sont signalées, les entités XML préservées', async () => {
    const app = loadApp();
    const unknown = new Set();
    const zip = await render(app, template(app, para('A &amp; B {Inconnue} {Titre}')), { Titre: 'C & D <E>' }, unknown);
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'A & B  C & D <E>');
    assert.equal([...unknown].join(','), 'Inconnue');
});
test('une balise pointée non résolue est vidée et signalée telle qu’écrite', async () => {
    const app = loadApp();
    const unknown = new Set();
    const zip = await render(app, template(app, para('Avant {Inconnue.Nom} après') + para('{Titre}')), { Titre: 'T' }, unknown);
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'Avant  aprèsT');
    assert.equal([...unknown].join(','), 'Inconnue.Nom');
});
test('boucle imbriquée : un champ de la ligne reste accessible dans la boucle', async () => {
    const app = loadApp();
    const zip = await render(app, template(app, para('{#Enfants}{Nom} de {Titre};{/Enfants}')), { Titre: 'P', Enfants: [{ Nom: 'A' }, { Nom: 'B' }] });
    assert.equal(xmlText(zip.file('word/document.xml').asText()), 'A de P;B de P;');
});

test('une table citée seulement dans l’en-tête est détectée et répétée', async () => {
    const app = loadApp();
    const buffer = template(app, para('Corps'), para('{{Enfants.Nom}}'));
    assert.equal(app.getReferencedTables(buffer).join(','), 'Enfants');
    const zip = await render(app, buffer, { Enfants: [{Nom:'Alice'}, {Nom:'Bob'}] });
    assert.equal(xmlText(zip.file('word/header1.xml').asText()), 'AliceBob');
});
test('signets, correcteur Word et balises adjacentes restent compatibles', async () => {
    const app = loadApp();
    const body = '<w:p>' + run('{Ti') + '<w:bookmarkStart w:id="1" w:name="test"/><w:proofErr w:type="spellStart"/>'
        + run('tre}{Participants}') + '<w:bookmarkEnd w:id="1"/></w:p>';
    const zip = await render(app, template(app, body), { Titre: 'Visite', Participants: 25 });
    const xml = zip.file('word/document.xml').asText();
    assert.equal(xmlText(xml), 'Visite25');
    assert.match(xml, /w:bookmarkStart/);
    assert.match(xml, /w:bookmarkEnd/);
});
