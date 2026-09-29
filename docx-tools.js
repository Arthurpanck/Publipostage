/**
 * MODULE DOCX
 * Version : 1.5
 * maintainer : dieux.alexandre@gmail.com
 * Gestion des fichiers Word (.docx)
 */

// Balises du dernier rendu n'ayant trouvé AUCUNE colonne correspondante
// (faute de frappe, libellé au lieu de l'identifiant de colonne...).
// Affichées dans la barre de statut pour diagnostiquer les "champs vides".
let lastUnknownTags = [];
let lastKnownKeys = [];

function getUnknownTags() {
    return lastUnknownTags;
}

function getKnownKeys() {
    return lastKnownKeys;
}

// Même compilation à l'import et à la génération : un modèle invalide ne doit
// pas remplacer le dernier modèle utilisable dans les options Grist.
function compileDocxTemplate(buffer, data, nullGetter) {
    try {
        const zip = new PizZip(buffer);
        // Résoudre les balises pointées avant de nettoyer les identifiants.
        transformDottedLoops(zip, data);
        sanitizeDocxXml(zip);
        const options = { paragraphLoop: true, linebreaks: true };
        if (nullGetter) options.nullGetter = nullGetter;
        return new window.docxtemplater(zip, options);
    } catch (error) {
        handleDocxError(error);
    }
}

// Génération du docx
function generateDocxBlob(data, buffer) {
    lastUnknownTags = [];
    lastKnownKeys = [];

    // Clés connues : champs du parent + colonnes des tables enfants,
    // pour distinguer "cellule vide" (normal) de "balise sans colonne" (erreur)
    const connues = new Set(Object.keys(data || {}));
    for (const cle of Object.keys(data || {})) {
        const valeur = data[cle];
        if (Array.isArray(valeur) && valeur.length > 0 && typeof valeur[0] === 'object' && valeur[0] !== null) {
            Object.keys(valeur[0]).forEach((col) => connues.add(col));
        }
    }
    const inconnues = new Set();

    // Une cellule vide est normale ; une balise sans colonne est signalée.
    const doc = compileDocxTemplate(buffer, data, (part) => {
        if (!part.module && part.value && !connues.has(part.value)) {
            inconnues.add(part.value);
        }
        return "";
    });

    try {
        // Ajout de la données custom
        doc.render(data);
    } catch (error) {
        handleDocxError(error);
    }
    lastUnknownTags = [...inconnues];
    lastKnownKeys = [...connues].filter((k) => !Array.isArray(data && data[k]));

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

// Le nettoyage est limité aux nœuds texte : jamais au XML entre deux balises.
function sanitizeDocxXml(zip) {
    for (const file of docxTextFiles(zip)) {
        const xml = repairDocxXml(file.asText()).replace(
            /(<w:t(?:\s[^>]*)?>)([^<]*)(<\/w:t>)/g,
            (_, open, text, close) => open + text.replace(/\{([^{}]+)\}/g, (match, key) => {
                if (/^[#/^@=]/.test(key) || key === '.') return match;
                return `{${sanitizeKey(key)}}`;
            }) + close
        );
        zip.file(file.name, xml);
    }
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

function sanitizeKey(keytoSanitize) {
    if (!keytoSanitize) {
        return "";
    }
    let sanitize = keytoSanitize.toString();
    sanitize = sanitize.normalize('NFKD');
    sanitize = sanitize.replace(/[\u0300-\u036f]/g, "");
    sanitize = sanitize.replace(/[^a-zA-Z0-9_]+/g, "_");
    sanitize = sanitize.replace(/^_+/, "");
    return sanitize;
}