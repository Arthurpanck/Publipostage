const { test } = require('node:test');
const assert = require('node:assert/strict');
const vm = require('node:vm');
const { loadApp, template, xmlText, para, run } = require('./helpers.cjs');
const { RECORDS, gristMock, fetchMock } = require('./fixtures/rendus-grist.cjs');

// Canvas simulé : chaque dessin produit un « PNG » qui dépend de ce qui est
// écrit et des couleurs, pour vérifier les images sans navigateur.
function fakeDocument(drawings) {
    const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==';
    return {
        createElement: () => {
            const ops = [];
            const ctx = new Proxy({}, {
                get: (t, k) => {
                    if (k in t) return t[k];
                    if (k === 'measureText') return (s) => ({ width: String(s).length * 7 });
                    if (k === 'fillText') return (s) => ops.push(`texte:${s}`);
                    return () => {};
                },
                set: (t, k, v) => { t[k] = v; if (/Style$|^font$/.test(k)) ops.push(`${k}:${v}`); return true; },
            });
            return {
                width: 0, height: 0, getContext: () => ctx,
                toDataURL: () => { drawings.push(ops); return 'data:image/png;base64,' + Buffer.concat([Buffer.from(PNG, 'base64'), Buffer.from(ops.join('|'))]).toString('base64'); },
            };
        },
    };
}

// Comme dispatchGeneration dans main.js, avec un Grist simulé à partir d'un
// vrai document (tests/fixtures/rendus-grist.cjs).
async function generate(body, record = RECORDS[0], { metadata = true, canvas = true } = {}) {
    const mock = gristMock(), drawings = [], urls = [];
    if (!metadata) {
        const fetchTable = mock.docApi.fetchTable;
        mock.docApi.fetchTable = async (t) => (t.startsWith('_grist') ? Promise.reject(new Error('refusé')) : fetchTable(t));
    }
    const app = loadApp({ grist: mock, atob, fetch: fetchMock(urls), ...(canvas ? { document: fakeDocument(drawings) } : {}) });
    Object.assign(app, { __rec: record, __buffer: template(app, body), __batch: { cache: new Map(), viewRows: RECORDS, unknownTags: new Set(), notes: new Set() } });
    const blob = await vm.runInContext(`(async () => {
        const data = rowData(__rec);
        await addLinkedTables(data, __rec, getReferencedTables(__buffer), __batch);
        const info = await loadRenderInfo(__batch);
        await loadAttachments(__batch, data, info);
        return generateDocxBlob(data, __buffer, __batch.unknownTags, info);
    })()`, app);
    const zip = new app.PizZip(await blob.arrayBuffer());
    const xml = zip.file('word/document.xml').asText();
    return { zip, xml, drawings, urls, unknown: [...app.__batch.unknownTags] };
}

// Descriptions des images, dans l'ordre du document
const alts = xml => [...xml.matchAll(/<wp:docPr [^>]*descr="([^"]*)"/g)].map(m => m[1]);
const images = xml => (xml.match(/<w:drawing>/g) || []).length;
const drawn = (drawings, text) => drawings.find(ops => ops.includes(`texte:${text}`));

test('booléen : case dessinée suivie du nom de la colonne, conditions intactes', async () => {
    const { xml } = await generate(para('{Mutuelle}   {Teletravail} {#Mutuelle}oui{/Mutuelle}{^Teletravail}non{/Teletravail}'));
    assert.deepEqual(alts(xml), ['Case cochée', 'Case non cochée']);
    assert.equal(xmlText(xml), ' Mutuelle    Télétravail ouinon');
});

test('choix : pilule teintée aux couleurs Grist, centrée sur la ligne', async () => {
    const body = '<w:p><w:r><w:rPr><w:sz w:val="28"/></w:rPr><w:t xml:space="preserve">Statut : {Statut} !</w:t></w:r></w:p>';
    const { xml, drawings } = await generate(body);
    assert.deepEqual(alts(xml), ['CDI']);
    assert.equal(xmlText(xml), 'Statut :  !');
    // Couleur de l'option CDI (#2E9E5B) : fond à 84 % de blanc, texte à 40 % de noir
    assert.ok(drawn(drawings, 'CDI').includes('fillStyle:#DEEFE5'));
    assert.ok(drawn(drawings, 'CDI').includes('fillStyle:#1C5F37'));
    // Taille de la balise conservée, image abaissée pour être centrée
    assert.match(xml, /<w:rPr><w:position w:val="-\d+"\/><w:sz w:val="28"\/><\/w:rPr><w:drawing>/);
});

test('choix clair dans Grist : la couleur sert de fond telle quelle, texte foncé de même teinte', () => {
    const app = loadApp();
    const pink = vm.runInContext('tagColors("#FED6FB")', app); // couleur du document de démonstration
    assert.equal(pink.bg, 'FED6FB');
    const [r, g, b] = [0, 2, 4].map(i => parseInt(pink.fg.slice(i, i + 2), 16));
    assert.ok(r > g && b > g && r + g + b < 300, pink.fg); // rose foncé
    assert.equal(vm.runInContext('tagColors("#2E9E5B").bg', app), 'DEEFE5'); // couleur soutenue : teintée
});

test('choix multiples, références : une pilule par valeur, références en gris', async () => {
    const { xml, drawings } = await generate(para('{Competences}|{Equipe}|{Pole_equipe}|{Collegues}'));
    assert.deepEqual(alts(xml), ['Sourcing', 'Entretiens', 'Anglais', 'Talents', 'RH', 'Julien Petit', 'Inès Garnier']);
    for (const name of ['Talents', 'RH', 'Julien Petit']) assert.ok(drawn(drawings, name).includes('fillStyle:#ECEEF1'), name);
});

test('valeurs vides : aucune image', async () => {
    const { xml } = await generate(para('[{Competences}][{Collegues}][{Photo}]'), RECORDS[2]);
    assert.equal(images(xml), 0);
    assert.equal(xmlText(xml), '[][][]');
});

test('étapes seules dans leur paragraphe : frise sur la largeur de la page', async () => {
    const { xml, drawings } = await generate(para('Avant') + '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>' + run('{Etape|etapes}') + '</w:p>' + para('Après'));
    assert.deepEqual(alts(xml), ['Offre &gt; Signature &gt; [Accueil] &gt; Formation &gt; Autonome']);
    assert.match(xml, /<w:p><w:pPr><w:jc w:val="center"\/><\/w:pPr><w:r><w:drawing>/); // paragraphe d'origine conservé
    assert.match(xml, /<wp:extent cx="6120130"/); // 9638 twips = 642,5 px
    for (const label of ['Offre', 'Accueil', 'Autonome']) assert.ok(drawn(drawings, label));
});

test('étapes dans une phrase : frise compacte, en ligne', async () => {
    const { xml } = await generate(para('Parcours : {Etape|etapes}.'), RECORDS[1]);
    assert.deepEqual(alts(xml), ['[Offre] &gt; Signature &gt; Accueil &gt; Formation &gt; Autonome']);
    assert.equal(xmlText(xml), 'Parcours : .');
    assert.match(xml, /<w:position w:val="-\d+"\/><\/w:rPr><w:drawing>/);
});

test('étapes dans une cellule : largeur de la cellule', async () => {
    const cell = '<w:tc><w:tcPr><w:tcW w:type="dxa" w:w="3216"/></w:tcPr>' + para('{Etape|etapes}') + '</w:tc>';
    const { xml } = await generate(`<w:tbl><w:tr>${cell}</w:tr></w:tbl>`);
    assert.match(xml, /<wp:extent cx="1905000"/); // (3216 - 216) / 15 = 200 px
});

test('barre : valeur sur la barre, objectif, valeurs formatées comme dans Grist', async () => {
    const { xml, drawings } = await generate(para('{Salaire|barre:Objectif}'));
    assert.deepEqual(alts(xml), ['2 450,00 € sur 4 000,00 €']);
    const ops = drawn(drawings, '2 450,00 €');
    assert.ok(ops.includes('texte:sur 4 000,00 €'));
    assert.ok(ops.includes('fillStyle:#DCEFEE'));
});

test('barre à trois valeurs au-delà de l’objectif : pourcentage et barre en rouge', async () => {
    const { xml, drawings } = await generate(para('{Salaire|barre:Objectif:Taux}'), RECORDS[1]);
    assert.deepEqual(alts(xml), ['105 % 2 100,00 € sur 2 000,00 €']);
    const ops = drawn(drawings, '105 %');
    assert.ok(ops.includes('fillStyle:#B42318') && ops.includes('fillStyle:#FCE1DD'));
});

test('barre sur des entiers, balise écrite librement', async () => {
    const { xml, unknown } = await generate(para('{ Conges_pris | Barre : Conges_acquis }'));
    assert.deepEqual(alts(xml), ['18 sur 25']);
    assert.deepEqual(unknown, []);
});

test('barre : colonne objectif introuvable signalée, valeur brute insérée', async () => {
    const { xml, unknown } = await generate(para('{Salaire|barre:Inconnue}'));
    assert.equal(xmlText(xml), '2450');
    assert.deepEqual(unknown, ['inconnue']);
});

test('pièces jointes : image réduite à 9 cm de haut ou à la largeur, autre fichier en étiquette', async () => {
    const { xml, zip, urls } = await generate(para('{Photo}'));
    assert.deepEqual(alts(xml), ['portrait.png', 'contrat.pdf']);
    assert.deepEqual(urls, ['https://grist.test/api/docs/demo/attachments/1/download?auth=jeton']); // le PDF n'est pas téléchargé
    // 800 x 600 réduit à 340 px de haut (9 cm) : 453 x 340 px
    assert.match(xml, /<wp:extent cx="4314825" cy="3238500"\/>/);
    assert.ok(zip.file('word/media/rendu_1.png'));
    assert.match(zip.file('[Content_Types].xml').asText(), /Extension="png"/);
});

test('mise en forme inconnue ou mal écrite : modèle refusé dès l’import', () => {
    const app = loadApp();
    for (const tag of ['{Statut|badge}', '{Salaire|barre}', '{Etape|etapes:Ordre}', '{#Statut|etapes}{/Statut}']) {
        assert.throws(() => app.compileDocxTemplate(template(app, para(tag))), /Erreur Template Word/, tag);
    }
});

test('tables liées : les colonnes typées de la table liée sont mises en forme', async () => {
    const { xml } = await generate(para('{Equipes.Nom} {Equipes.Pole}'));
    assert.equal(xmlText(xml), 'Talents ');
    assert.deepEqual(alts(xml), ['RH']);
});

test('sans métadonnées Grist : valeurs brutes', async () => {
    const { xml } = await generate(para('{Mutuelle} {Statut}'), RECORDS[0], { metadata: false });
    assert.equal(xmlText(xml), 'true CDI');
    assert.equal(images(xml), 0);
});

test('sans canvas : valeurs insérées en texte', async () => {
    const { xml } = await generate(para('{Mutuelle} {Competences} {Salaire|barre:Objectif} {Etape|etapes}'), RECORDS[0], { canvas: false });
    assert.equal(xmlText(xml), 'true Sourcing, Entretiens, Anglais 2450 Accueil');
});

test('le texte voisin de la balise garde sa mise en forme', async () => {
    const body = '<w:p><w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">a {Mutuelle} b</w:t></w:r>' + run(' c') + '</w:p>';
    const { xml } = await generate(body);
    const texts = [...xml.matchAll(/<w:r>(?:<w:rPr>([\s\S]*?)<\/w:rPr>)?<w:t[^>]*>([^<]*)<\/w:t><\/w:r>/g)];
    assert.deepEqual(texts.map(m => m[2]), ['a ', ' Mutuelle', ' b', ' c']);
    assert.ok(texts.slice(0, 3).every(m => /<w:i\/>/.test(m[1])));
});

test('images : un média par image identique, identifiants uniques', async () => {
    const { xml, zip } = await generate(para('{Statut} {Statut} {Competences}'));
    assert.equal(images(xml), 5);
    assert.equal(Object.keys(zip.files).filter(f => f.startsWith('word/media/')).length, 4); // CDI dessiné deux fois à l'identique
    const ids = [...xml.matchAll(/<wp:docPr id="(\d+)"/g)].map(m => m[1]);
    assert.equal(new Set(ids).size, 5);
    const rels = zip.file('word/_rels/document.xml.rels').asText();
    assert.equal((rels.match(/<Relationship /g) || []).length, 4);
});

test('pièces jointes dans une cellule : réduites à la largeur de la cellule', async () => {
    const cell = '<w:tc><w:tcPr><w:tcW w:w="2616" w:type="dxa"/></w:tcPr>' + para('{Photo}') + '</w:tc>';
    const { xml } = await generate(`<w:tbl><w:tr>${cell}</w:tr></w:tbl>`);
    assert.match(xml, /<wp:extent cx="1524000" cy="1143000"\/>/); // (2616 - 216) / 15 = 160 px de large
});

test('colonne du même nom que la table liée : pilules et fiches liées ensemble', async () => {
    const { xml } = await generate(para('{Personnes}') + para('{Personnes.Nom} : {Personnes.Fonction}'));
    assert.deepEqual(alts(xml), ['Julien Petit', 'Inès Garnier']);
    assert.equal(xmlText(xml), ' Julien Petit : Gestionnaire de paieInès Garnier : Juriste social'); // espace entre les deux pilules
});

test('pièces jointes d’une table liée', async () => {
    const { xml, urls } = await generate(para('{Equipes.Nom} {Equipes.Logo}'));
    assert.deepEqual(alts(xml), ['portrait.png']);
    assert.equal(urls.length, 1);
});

test('relations auto-fermantes des vrais modèles Word : l’image est bien reliée', async () => {
    const drawings = [];
    const app = loadApp({ grist: gristMock(), atob, fetch: fetchMock(), document: fakeDocument(drawings) });
    const zip = new app.PizZip(template(app, para('Corps'), '', para('{Statut}')));
    zip.file('word/_rels/footer1.xml.rels', '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>');
    Object.assign(app, { __rec: RECORDS[0], __buffer: zip.generate({ type: 'uint8array' }), __batch: { cache: new Map(), viewRows: RECORDS, unknownTags: new Set(), notes: new Set() } });
    const blob = await vm.runInContext(`(async () => {
        const data = rowData(__rec);
        return generateDocxBlob(data, __buffer, __batch.unknownTags, await loadRenderInfo(__batch));
    })()`, app);
    const out = new app.PizZip(await blob.arrayBuffer());
    const rels = out.file('word/_rels/footer1.xml.rels').asText();
    const rId = /r:embed="(\w+)"/.exec(out.file('word/footer1.xml').asText())[1];
    assert.match(rels, new RegExp(`<Relationship Id="${rId}" [^>]*Target="media/rendu_1.png"/></Relationships>`));
});

test('référence reçue développée par Grist ({ id, … }) : fiche liée et étiquette', async () => {
    const record = { ...RECORDS[0], Equipe: { id: 1, Nom: 'Talents', Pole: 'RH' } };
    const { xml } = await generate(para('{Equipe} · {Equipes.Pole}'), record);
    assert.deepEqual(alts(xml), ['Talents', 'RH']);
});

test('pièces jointes : tailles petit, moyen, grand', async () => {
    const { xml } = await generate(para('{Photo|petit}') + para('{Photo|moyen}') + para('{Photo|grand}'));
    const heights = [...xml.matchAll(/<wp:extent cx="\d+" cy="(\d+)"\/>/g)].map(m => Math.round(+m[1] / 9525))
        .filter((_, i) => i % 2 === 0); // la pilule du PDF suit chaque image
    assert.deepEqual(heights, [76, 151, 340]);
});

test('taille d’image sur une colonne qui n’est pas une pièce jointe : signalée, valeur brute', async () => {
    const { xml, unknown } = await generate(para('{Statut|petit}'));
    assert.equal(xmlText(xml), 'CDI');
    assert.deepEqual(unknown, ['statut|petit (taille réservée aux colonnes Pièce jointe)']);
    const app = loadApp();
    assert.throws(() => app.compileDocxTemplate(template(app, para('{Photo|petit:3}'))), /Erreur Template Word/);
});

// Ce que Grist transmet réellement au widget (expandRefs, par défaut) : la valeur
// affichée des références, pas leur id (app/client/components/WidgetFrame.ts).
const AS_GRIST_SENDS = { ...RECORDS[0], Equipe: 'Talents', Pole_equipe: 'RH', Collegues: ['Julien Petit', 'Inès Garnier'], Personnes: ['Julien Petit', 'Inès Garnier'] };

test('références reçues sous leur valeur affichée : étiquettes et fiches liées', async () => {
    const body = para('{Equipe} {Collegues}') + para('{Equipes.Pole} · {Equipes.Logo|petit}') + para('{Personnes.Nom} : {Personnes.Fonction}');
    const { xml, unknown } = await generate(body, AS_GRIST_SENDS);
    assert.deepEqual(alts(xml), ['Talents', 'Julien Petit', 'Inès Garnier', 'RH', 'portrait.png']);
    assert.match(xmlText(xml), /Julien Petit : Gestionnaire de paieInès Garnier : Juriste social$/);
    assert.deepEqual(unknown, []);
});

test('colonne du même nom que la table liée, reçue sous sa valeur affichée', async () => {
    const { xml } = await generate(para('{Personnes}') + para('{Personnes.Nom}'), AS_GRIST_SENDS);
    assert.deepEqual(alts(xml), ['Julien Petit', 'Inès Garnier']);
});

// Modèle avec styles.xml et thème, comme un vrai document Word
async function generateStyled(body, record = RECORDS[0]) {
    const drawings = [];
    const app = loadApp({ grist: gristMock(), atob, fetch: fetchMock(), document: fakeDocument(drawings) });
    const zip = new app.PizZip(template(app, body));
    zip.file('word/styles.xml', `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
        <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
        <w:style w:type="paragraph" w:styleId="Titre"><w:basedOn w:val="Normal"/><w:rPr><w:rFonts w:ascii="Georgia" w:hAnsi="Georgia"/><w:sz w:val="36"/></w:rPr></w:style>
        <w:style w:type="character" w:styleId="Petit"><w:rPr><w:sz w:val="16"/></w:rPr></w:style></w:styles>`);
    zip.file('word/theme/theme1.xml', '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements><a:fontScheme><a:majorFont><a:latin typeface="Aptos Display"/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/></a:minorFont></a:fontScheme></a:themeElements></a:theme>');
    Object.assign(app, { __rec: record, __buffer: zip.generate({ type: 'uint8array' }), __batch: { cache: new Map(), viewRows: RECORDS, unknownTags: new Set(), notes: new Set() } });
    const blob = await vm.runInContext('(async () => generateDocxBlob(rowData(__rec), __buffer, __batch.unknownTags, await loadRenderInfo(__batch)))()', app);
    return { xml: new app.PizZip(await blob.arrayBuffer()).file('word/document.xml').asText(), drawings };
}
const fontOf = (drawings, text) => drawn(drawings, text).find(op => op.startsWith('font:600'));

test('étiquette : police et taille de la balise (thème, style de paragraphe, style de caractère)', async () => {
    const body = para('{Statut}')
        + '<w:p><w:pPr><w:pStyle w:val="Titre"/></w:pPr>' + run('{Competences}') + '</w:p>'
        + '<w:p><w:r><w:rPr><w:rStyle w:val="Petit"/></w:rPr><w:t>{Equipe}</w:t></w:r></w:p>';
    const { drawings } = await generateStyled(body);
    assert.match(fontOf(drawings, 'CDI'), /^font:600 14\.66\d*px "Aptos", /); // défauts : thème « Corps », 11 pt
    assert.match(fontOf(drawings, 'Sourcing'), /^font:600 24px "Georgia", /); // style de paragraphe : 18 pt
    assert.match(fontOf(drawings, 'Talents'), /^font:600 10\.66\d*px "Aptos", /); // style de caractère : 8 pt
});

test('étiquette trop large pour sa cellule : passe sur plusieurs lignes', async () => {
    const cell = '<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr>' + para('{Statut}') + '</w:tc>';
    const { drawings } = await generate(`<w:tbl><w:tr>${cell}</w:tr></w:tbl>`, RECORDS[2]); // « Période d'essai »
    const ops = drawings.find(o => o.some(op => op.startsWith('texte:Période')));
    assert.deepEqual(ops.filter(op => op.startsWith('texte:')), ['texte:Période', "texte:d'essai"]);
});
