/**
 * MODULE DOCX
 * Version : 2.0
 * maintainer : dieux.alexandre@gmail.com
 * Gestion des fichiers Word (.docx)
 *
 * Syntaxe dans le modèle :
 *   - Champ de la ligne : {Nom} ou {{Nom}}, identifiant de colonne Grist
 *     (accents et caractères spéciaux tolérés, comme sanitizeKey).
 *   - Lignes d'une table liée : {Membres.Nom_Membre} (voir data-tools.js).
 *     Dans une ligne de tableau : la ligne est répétée pour chaque ligne liée.
 *     Hors tableau : le paragraphe est répété pour chaque ligne liée.
 *   - Boucle manuelle docxtemplater : {#Membres}{Nom_Membre}{/Membres}.
 *
 * Limites connues :
 *   - Une seule table liée par ligne de tableau ou paragraphe répété.
 */

// Balise pointée {Table.Colonne} (accents tolérés côté colonne)
const DOTTED_TAG_SRC = '\\{([A-Za-z0-9_]+)\\.([A-Za-z0-9_À-ÿ]+)\\}';

// Modèle analysé une seule fois : XML réparé de chaque partie texte et tables
// citées. Chaque génération repart de ce résultat.
let parsedTemplate = { buffer: null };

function parseTemplate(buffer) {
    if (parsedTemplate.buffer !== buffer) {
        const zip = new PizZip(buffer);
        const parts = {};
        const tables = new Set();
        for (const file of docxTextFiles(zip)) {
            const xml = repairDocxXml(file.asText());
            parts[file.name] = xml;
            dottedTablesIn(xml).forEach(table => tables.add(table));
            // boucles explicites {#Table}...{/Table} écrites à la main dans le modèle
            for (const m of xml.matchAll(/\{[#^]([A-Za-z0-9_]+)\}/g)) {
                tables.add(m[1]);
            }
        }
        parsedTemplate = { buffer: buffer, parts: parts, tables: [...tables] };
    }
    return parsedTemplate;
}

// Tables citées dans le modèle par {Table.Colonne} ou {#Table}
function getReferencedTables(buffer) {
    return parseTemplate(buffer).tables;
}

// Même compilation à l'import et à la génération : un modèle invalide ne doit
// pas remplacer le dernier modèle utilisable dans les options Grist.
function compileDocxTemplate(buffer, data, nullGetter) {
    try {
        const zip = new PizZip(buffer);
        for (const [name, xml] of Object.entries(parseTemplate(buffer).parts)) {
            zip.file(name, transformDottedXml(xml, data));
        }
        return new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            // Une balise désigne la colonne de même identifiant nettoyé
            parser: (tag) => {
                const key = sanitizeKey(tag);
                return { get: (scope) => (tag === '.' || !scope) ? scope : scope[key] };
            },
            nullGetter: nullGetter,
        });
    } catch (error) {
        handleDocxError(error);
    }
}

// Génération du docx. Les balises sans colonne correspondante (faute de frappe,
// libellé au lieu de l'identifiant de colonne...) sont ajoutées à unknownTags.
function generateDocxBlob(data, buffer, unknownTags = new Set()) {
    // Clés connues : champs de la ligne + colonnes des tables liées,
    // pour distinguer "cellule vide" (normal) de "balise sans colonne" (erreur)
    const connues = new Set(Object.keys(data));
    for (const valeur of Object.values(data)) {
        if (Array.isArray(valeur) && valeur.length > 0 && typeof valeur[0] === 'object' && valeur[0] !== null) {
            Object.keys(valeur[0]).forEach((col) => connues.add(col));
        }
    }

    const doc = compileDocxTemplate(buffer, data, (part) => {
        if (!part.module && part.value && !connues.has(sanitizeKey(part.value))) {
            unknownTags.add(part.value);
        }
        return "";
    });

    try {
        // Ajout de la données custom
        doc.render(data);
    } catch (error) {
        handleDocxError(error);
    }

    return doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
}

// Toutes les parties Word contenant du texte (pas les propriétés/métadonnées).
function docxTextFiles(zip) {
    return zip.file(/^word\/(?:document|header[0-9]*|footer[0-9]*|footnotes|endnotes)\.xml$/);
}

/**
 * Recolle uniquement le TEXTE des balises, sans fusionner les runs Word.
 * Le texte voisin conserve donc sa mise en forme, les signets et le XML.
 * Une balise ne traverse jamais un paragraphe, une tabulation ou un saut.
 * Les doubles accolades équilibrées sont un alias de la syntaxe historique.
 * Les accolades mal appariées restent intactes pour le diagnostic du moteur.
 */
function repairDocxXml(xml) {
    return xml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[ />])[\s\S])*?<\/w:p>/g, (paragraph) => {
        const nodes = [];
        let text = '';
        const parts = /<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>|<w:(?:tab|br|cr|drawing|object)\b[^>]*\/?>/g;
        let part;
        while ((part = parts.exec(paragraph)) !== null) {
            if (part[1] === undefined) {
                text += '\0';
                continue;
            }
            nodes.push({ start: part.index + part[0].indexOf('>') + 1,
                offset: text.length, length: part[1].length, text: part[1] });
            text += part[1];
        }
        const tags = [...text.matchAll(/\{\{([^{}\0]+)\}\}|\{([^{}\0]+)\}/g)];
        // De droite à gauche : les offsets des balises précédentes restent valides.
        for (const tag of tags.reverse()) {
            const start = tag.index, end = start + tag[0].length;
            if (text[start - 1] === '{' || text[end] === '}') continue;
            const replacement = '{' + (tag[1] || tag[2]).trim() + '}';
            for (const node of nodes) {
                const from = Math.max(start - node.offset, 0);
                const to = Math.min(end - node.offset, node.length);
                if (from >= to) continue;
                node.text = node.text.slice(0, from)
                    + (start >= node.offset ? replacement : '') + node.text.slice(to);
            }
        }
        for (const node of nodes.reverse()) {
            paragraph = paragraph.slice(0, node.start) + node.text
                + paragraph.slice(node.start + node.length);
        }
        return paragraph;
    });
}

// Tables citées via {Table.Col} dans un fragment XML
function dottedTablesIn(fragment) {
    const found = new Set();
    const tagRe = new RegExp(DOTTED_TAG_SRC, 'g');
    let m;
    while ((m = tagRe.exec(fragment)) !== null) {
        found.add(m[1]);
    }
    return [...found];
}

// Traduction de syntaxe au moment de générer :
// - {Membres.Col} dans une ligne de tableau -> {#Membres}{Col}...{/Membres}
//   enroulant toute la ligne, pour que docxtemplater répète la ligne.
// - {Membres.Col} ailleurs (hors tableau, cellule unique, tableau imbriqué) ->
//   le paragraphe entier est répété (marqueurs + paragraphLoop).
// SEULES les tables réellement résolues dans `data` sont transformées : une
// balise pointée non résolue reste une balise simple, vide et signalée comme
// sans correspondance, au lieu de faire disparaître silencieusement la ligne
// ou le paragraphe (et les balises simples comme {EEE} qui s'y trouvent).
function transformDottedXml(xml, data) {

    const estResolue = (t) => data && Array.isArray(data[t]);
    const enBoucle = (fragment, table) => fragment.replace(
        new RegExp(`\\{${table}\\.([A-Za-z0-9_À-ÿ]+)\\}`, 'g'),
        '{$1}'
    );

    // 1) Lignes de tableau LES PLUS INTERNES (le garde-fou (?!<\/?w:tr[ >])
    //    empêche d'enjamber un tableau imbriqué) d'au moins 2 cellules : la
    //    ligne est répétée. Avec 1 seule cellule, docxtemplater répéterait le
    //    contenu bout à bout au lieu de la ligne -> passe 2.
    xml = xml.replace(/<w:tr[ >](?:(?!<\/?w:tr[ >])[\s\S])*?<\/w:tr>/g, (row) => {
        const tablesInRow = dottedTablesIn(row).filter(estResolue);
        if (tablesInRow.length === 0) {
            return row; // ligne normale, on ne touche pas
        }
        const cellules = (row.match(/<w:tc[ >]/g) || []).length;
        if (cellules < 2) {
            return row; // cellule unique : répétition de paragraphe (passe 2)
        }
        if (tablesInRow.length > 1) {
            console.warn("Plusieurs tables enfants dans la même ligne, seule la première est répétée :", tablesInRow);
        }
        const table = tablesInRow[0];

        // {Table.Col} -> {Col} puis enrouler la ligne : {#table} dans le 1er
        // nœud texte, {/table} dans le dernier.
        // (?:\s[^>]*)? cible uniquement <w:t>, pas <w:tr>/<w:tc>/<w:tbl>...
        return enBoucle(row, table)
            .replace(/(<w:t(?:\s[^>]*)?>)/, `$1{#${table}}`)
            .replace(/(<\/w:t>)(?![\s\S]*<\/w:t>)/, `{/${table}}$1`);
    });

    // 2) Balises pointées restantes : le paragraphe est répété pour chaque
    //    enfant (valable hors tableau comme dans une cellule). Le garde-fou
    //    (?!<w:p[ />]) évite d'enjamber un autre paragraphe, et l'ouverture
    //    <w:p(?:\s...)?> exclut les paragraphes vides auto-fermés <w:p/>.
    xml = xml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[ />])[\s\S])*?<\/w:p>/g, (para) => {
        const tablesInPara = dottedTablesIn(para).filter(estResolue);
        if (tablesInPara.length === 0) {
            return para;
        }
        if (tablesInPara.length > 1) {
            console.warn("Plusieurs tables enfants dans le même paragraphe, seule la première est répétée :", tablesInPara);
        }
        const table = tablesInPara[0];

        // paragraphes marqueurs : avec paragraphLoop, docxtemplater les retire
        // et répète le paragraphe central pour chaque enregistrement enfant
        return `<w:p><w:r><w:t>{#${table}}</w:t></w:r></w:p>`
            + enBoucle(para, table)
            + `<w:p><w:r><w:t>{/${table}}</w:t></w:r></w:p>`;
    });

    return xml;
}

// Affichage des erreurs
function handleDocxError(error) {
    if (error.properties && error.properties.errors instanceof Array) {
        const errorMessages = error.properties.errors.map(function (err) {
            return err.properties.explanation;
        }).join("\n");
        throw new Error("Erreur Template Word :\n" + errorMessages);
    } else {
        throw error;
    }
}
