const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { loadApp, template, xmlText, para, run } = require('./helpers.cjs');
const { RECORDS, gristMock } = require('./fixtures/rendus-grist.cjs');

// Comme dispatchGeneration dans main.js, avec un Grist simulé à partir d'un
// vrai document (tests/fixtures/rendus-grist.cjs).
async function generate(body, record = RECORDS[0], { metadata = true } = {}) {
    const fetched = [];
    const mock = gristMock(fetched);
    if (!metadata) {
        const fetchTable = mock.docApi.fetchTable;
        mock.docApi.fetchTable = async (t) => (t.startsWith('_grist') ? Promise.reject(new Error('refusé')) : fetchTable(t));
    }
    const app = loadApp({ grist: mock });
    Object.assign(app, { __rec: record, __buffer: template(app, body), __batch: { cache: new Map(), viewRows: RECORDS, unknownTags: new Set(), notes: new Set() } });
    const blob = await vm.runInContext(`(async () => {
        const data = rowData(__rec);
        await addLinkedTables(data, __rec, getReferencedTables(__buffer), __batch);
        return generateDocxBlob(data, __buffer, __batch.unknownTags, await loadRenderInfo(__batch));
    })()`, app);
    const zip = new app.PizZip(await blob.arrayBuffer());
    return { xml: zip.file('word/document.xml').asText(), unknown: [...app.__batch.unknownTags], fetched };
}

// Runs du document : [{ text, rPr }]
function runs(xml) {
    return [...xml.matchAll(/<w:r>(?:<w:rPr>([\s\S]*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/g)]
        .map(m => ({ rPr: m[1] || '', text: m[2] }));
}

test('booléen : case cochée ou vide suivie du nom de la colonne, sans toucher aux conditions', async () => {
    const { xml } = await generate(para('{Mutuelle}   {Teletravail} {#Mutuelle}oui{/Mutuelle}{^Teletravail}non{/Teletravail}'));
    assert.equal(xmlText(xml), '☑︎ Mutuelle   ☐︎ Télétravail ouinon');
    assert.match(xml, /w:ascii="Segoe UI Symbol"/);
});

test('choix : tag teinté et gras aux couleurs Grist, mise en forme de la balise conservée', async () => {
    const body = '<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Statut : {Statut} !</w:t></w:r></w:p>';
    const { xml } = await generate(body);
    const tag = runs(xml).find(r => r.text.includes('CDI'));
    assert.equal(tag.text, ' CDI ');
    assert.match(tag.rPr, /<w:b\/>/);
    assert.match(tag.rPr, /<w:sz w:val="28"\/>/);
    assert.match(tag.rPr, /<w:shd w:val="clear" w:color="auto" w:fill="[0-9A-F]{6}"\/>/);
    // Ordre du schéma Word : b avant color, sz avant shd
    assert.ok(tag.rPr.indexOf('<w:b/>') < tag.rPr.indexOf('<w:color') && tag.rPr.indexOf('<w:sz') < tag.rPr.indexOf('<w:shd'));
    assert.equal(xmlText(xml), 'Statut :  CDI  !');
});

test('choix multiples et références multiples : un tag par valeur', async () => {
    const { xml } = await generate(para('{Competences}|{Collegues}'));
    assert.equal(xmlText(xml), ' Sourcing   Entretiens   Anglais | Julien Petit   Inès Garnier ');
    const fills = runs(xml).filter(r => /w:shd/.test(r.rPr)).map(r => /w:fill="(\w+)"/.exec(r.rPr)[1]);
    assert.equal(new Set(fills.slice(0, 3)).size, 3); // trois couleurs Grist différentes
    assert.deepEqual(fills.slice(3), ['EEF1F4', 'EEF1F4']); // Personnes.Nom est un texte : gris
});

test('référence : valeur affichée de la fiche liée, toujours en gris', async () => {
    const { xml } = await generate(para('{Equipe}/{Pole_equipe}'));
    assert.equal(xmlText(xml), ' Talents / RH ');
    const fills = runs(xml).filter(r => /w:shd/.test(r.rPr)).map(r => /w:fill="(\w+)"/.exec(r.rPr)[1]);
    assert.deepEqual(fills, ['EEF1F4', 'EEF1F4']); // même si la colonne affichée est un choix coloré
});

test('valeurs vides : aucun tag', async () => {
    const { xml } = await generate(para('[{Competences}][{Collegues}]'), RECORDS[2]);
    assert.equal(xmlText(xml), '[][]');
});

test('étapes seules dans leur paragraphe : tableau dans l’ordre des choix', async () => {
    const { xml } = await generate(para('Avant') + para('{Etape|etapes}') + para('Après'));
    assert.equal((xml.match(/<w:tbl>/g) || []).length, 1);
    assert.equal(xmlText(xml), 'Avant●●◉○○OffreSignatureAccueilFormationAutonomeAprès');
    assert.match(xml, /<w:b\/><w:bCs\/><w:color w:val="1F1F1F"\/><w:sz w:val="17"\/><w:szCs w:val="17"\/><\/w:rPr><w:t xml:space="preserve">Accueil/);
});

test('étapes dans une phrase : version en ligne', async () => {
    const { xml } = await generate(para('Parcours : {Etape|etapes}.'), RECORDS[1]);
    assert.equal((xml.match(/<w:tbl>/g) || []).length, 0);
    assert.equal(xmlText(xml), 'Parcours : ◉ Offre  ›  ○ Signature  ›  ○ Accueil  ›  ○ Formation  ›  ○ Autonome.');
});

test('étapes dans une cellule de tableau : la cellule garde un paragraphe final', async () => {
    const { xml } = await generate(`<w:tbl><w:tr><w:tc>${para('{Etape|etapes}')}</w:tc></w:tr></w:tbl>`);
    assert.match(xml, /<\/w:tbl><w:p\/><\/w:tc>/);
});

test('barre : valeur sur la barre et objectif, valeurs formatées comme dans Grist', async () => {
    const { xml } = await generate(para('{Salaire|barre:Objectif}'));
    const text = xmlText(xml);
    assert.match(text, /^ 2 450,00 € +  sur 4 000,00 €$/);
    const bar = runs(xml).filter(r => /w:bdr/.test(r.rPr));
    assert.equal(bar.length, 2);
    const width = bar[0].text.length + bar[1].text.length;
    assert.equal(bar[0].text.length, Math.round(0.6125 * width)); // remplissage 2450 / 4000
    assert.match(bar[0].rPr, /w:fill="DCEFEE"/);
});

test('barre à trois valeurs : pourcentage de la colonne à gauche', async () => {
    const { xml } = await generate(para('{Salaire|barre:Objectif:Taux}'), RECORDS[1]);
    assert.match(xmlText(xml), /^105 %   2 100,00 € +  sur 2 000,00 €$/);
    const bar = runs(xml).filter(r => /w:bdr/.test(r.rPr));
    assert.equal(bar.length, 1); // dépasse l'objectif : barre pleine, en rouge
    assert.match(bar[0].rPr, /w:fill="FCE1DD"/);
    assert.match(runs(xml)[0].rPr, /<w:b\/>[\s\S]*w:val="B42318"/);
});

test('barre sur des entiers, balise écrite librement', async () => {
    const { xml, unknown } = await generate(para('{ Conges_pris | Barre : Conges_acquis }'));
    assert.match(xmlText(xml), /^ 18 +  sur 25$/);
    assert.deepEqual(unknown, []);
});

test('barre : colonne objectif introuvable signalée, valeur brute insérée', async () => {
    const { xml, unknown } = await generate(para('{Salaire|barre:Inconnue}'));
    assert.equal(xmlText(xml), '2450');
    assert.deepEqual(unknown, ['inconnue']);
});

test('mise en forme inconnue ou mal écrite : modèle refusé dès l’import', () => {
    const app = loadApp();
    for (const tag of ['{Statut|badge}', '{Salaire|barre}', '{Etape|etapes:Ordre}', '{#Statut|etapes}{/Statut}']) {
        assert.throws(() => app.compileDocxTemplate(template(app, para(tag))), /Erreur Template Word/, tag);
    }
});

test('tables liées : les colonnes typées de la table liée sont mises en forme', async () => {
    const { xml } = await generate(para('{Equipes.Nom} {Equipes.Pole}'));
    assert.equal(xmlText(xml), 'Talents  RH ');
});

test('sans métadonnées Grist : valeurs brutes, la barre reste disponible', async () => {
    const { xml } = await generate(para('{Mutuelle} {Statut} ') + para('{Salaire|barre:Objectif}'), RECORDS[0], { metadata: false });
    const text = xmlText(xml);
    assert.ok(text.startsWith('true CDI '));
    assert.match(text, /2450 +  sur 4000$/);
});

test('le texte voisin de la balise garde sa mise en forme', async () => {
    const body = '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">a {Mutuelle} b</w:t></w:r>' + run(' c') + '</w:p>';
    const { xml } = await generate(body);
    const r = runs(xml);
    assert.deepEqual(r.map(x => x.text), ['a ', '☑︎', ' Mutuelle', ' b', ' c']);
    assert.ok(r.slice(0, 4).every(x => /<w:i\/>/.test(x.rPr)));
});

// Canvas simulé : la voie « images » (navigateur) produit des PNG insérés en ligne
function fakeDocument(calls) {
    const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    const ctx2d = new Proxy({}, { get: (t, k) => (k in t ? t[k] : k === 'measureText' ? (s) => ({ width: String(s).length * 7 }) : () => {}), set: (t, k, v) => { t[k] = v; return true; } });
    return { createElement: () => ({ getContext: () => ctx2d, toDataURL: () => { calls.push(1); return PNG; } }) };
}
async function generateDrawn(body, record = RECORDS[0]) {
    const calls = [];
    const app = loadApp({ grist: gristMock(), document: fakeDocument(calls), atob });
    Object.assign(app, { __rec: record, __buffer: template(app, body), __batch: { cache: new Map(), viewRows: RECORDS, unknownTags: new Set(), notes: new Set() } });
    const blob = await vm.runInContext(`(async () => {
        const data = rowData(__rec);
        await addLinkedTables(data, __rec, getReferencedTables(__buffer), __batch);
        return generateDocxBlob(data, __buffer, __batch.unknownTags, await loadRenderInfo(__batch));
    })()`, app);
    return { zip: new app.PizZip(await blob.arrayBuffer()), calls };
}

test('navigateur : étiquettes, case, barre et étapes dessinées en images', async () => {
    const body = para('{Statut} {Competences} {Mutuelle} {Salaire|barre:Objectif:Taux}') + para('{Etape|etapes}');
    const { zip, calls } = await generateDrawn(body);
    const xml = zip.file('word/document.xml').asText();
    // 1 + 3 étiquettes, 1 case, 1 barre, 1 frise
    assert.equal((xml.match(/<w:drawing>/g) || []).length, 7);
    assert.equal(calls.length, 7);
    assert.match(xml, /descr="CDI"/);
    assert.match(xml, /descr="Case cochée"/);
    assert.match(xml, /descr="61 % 2 450,00 € sur 4 000,00 €"/);
    assert.match(xml, /descr="Offre &gt; Signature &gt; \[Accueil\] &gt; Formation &gt; Autonome"/);
    assert.equal((xml.match(/<w:tbl>/g) || []).length, 0);
    // Image identique (même PNG) : un seul média, une seule relation
    assert.deepEqual(Object.keys(zip.files).filter(f => f.startsWith('word/media/')), ['word/media/rendu_1.png']);
    const rels = zip.file('word/_rels/document.xml.rels').asText();
    assert.equal((rels.match(/<Relationship /g) || []).length, 1);
    assert.match(zip.file('[Content_Types].xml').asText(), /Extension="png" ContentType="image\/png"/);
    // Les étiquettes sont abaissées pour être centrées sur le texte
    assert.match(xml, /<w:position w:val="-\d+"\/><\/w:rPr><w:drawing>/);
});

test('navigateur : identifiants d’images uniques dans le document', async () => {
    const { zip } = await generateDrawn(para('{Competences} {Collegues}'));
    const ids = [...zip.file('word/document.xml').asText().matchAll(/<wp:docPr id="(\d+)"/g)].map(m => m[1]);
    assert.equal(ids.length, 5);
    assert.equal(new Set(ids).size, 5);
});
