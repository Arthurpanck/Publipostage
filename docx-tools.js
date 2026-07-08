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

function getUnknownTags() {
    return lastUnknownTags;
}

// Génération du docx
function generateDocxBlob(data, buffer) {
    const zip = new PizZip(buffer);
    lastUnknownTags = [];

    // Traduction des boucles {Table.Colonne} AVANT la sanitisation des clés
    // (sinon le "." serait remplacé par "_" et la traduction échouerait).
    // `data` sert à ne transformer que les tables réellement résolues.
    try {
        transformDottedLoops(zip, data);
    } catch (e) {
        console.warn("ATTENTION : La transformation des boucles relationnelles a échoué", e);
    }

    try {
        sanitizeDocxXml(zip);
    } catch (e) {
        console.warn("ATTENTION : Le nettoyage automatique du XML a échoué", e);
    }

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

    let doc;
    try {
        doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
            // Par défaut docxtemplater écrit le texte "undefined" quand une
            // balise n'a pas de valeur. On rend une chaîne vide à la place,
            // et on mémorise les balises sans colonne pour les signaler.
            nullGetter: function (part) {
                if (!part.module && part.value && !connues.has(part.value)) {
                    inconnues.add(part.value);
                }
                return "";
            },
        });
    } catch(error) {
        handleDocxError(error);
    }

    try {
        // Ajout de la données custom
        doc.render(data);
    } catch (error) {
        handleDocxError(error);
    }
    lastUnknownTags = [...inconnues];

    return doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    });
}

/**
 * Réparation du XML Word : retire les marqueurs du correcteur et recolle les
 * balises {..} que Word a éclatées sur plusieurs "runs". Word fragmente une
 * balise dès qu'on l'édite (runs avec propriétés <w:rPr>, rsid), qu'un mot est
 * souligné par le correcteur (<w:proofErr>) ou que le curseur y a laissé un
 * signet (_GoBack). Sans réparation, la balise n'est pas reconnue et la
 * sanitisation avale le XML entre "{" et le "}" suivant, cassant le document.
 */
function repairDocxXml(xml) {
    // Suppression des balises de correction orthographique et de grammaire
    xml = xml.replace(/<w:proofErr[^>]*\/>/g, "");
    xml = xml.replace(/<w:lang[^>]*\/>/g, "");
    xml = xml.replace(/<w:noProof[^>]*\/>/g, "");
    // reconstruction des balises cassées entre deux runs simples adjacents
    xml = xml.replace(/<\/w:t><\/w:r><w:r[^>]*><w:t[^>]*>/g, "");
    // reconstruction d'une balise "{" restée ouverte en fin de run : on fusionne
    // le run suivant (avec ses éventuels <w:rPr> et signets intercalés) jusqu'à
    // ce que la balise soit recollée, en gardant la mise en forme du 1er run
    // (?:\s[^>]*)? cible uniquement <w:t>/<w:t attr...>, jamais <w:tab/> ni
    // d'autres éléments, pour ne pas produire de XML imbriqué invalide
    const splitTag = /(\{[^{}<]*)<\/w:t><\/w:r>(?:<w:bookmark(?:Start|End)[^>]*\/>)*<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t(?:\s[^>]*)?>/g;
    let avant;
    do {
        avant = xml;
        xml = xml.replace(splitTag, "$1");
    } while (xml !== avant);
    return xml;
}

// Nettoyage des docxs
function sanitizeDocxXml(zip) {
    const xmlFile = "word/document.xml";
    if (!zip.file(xmlFile)) {
        return;
    }

    let xml = repairDocxXml(zip.file(xmlFile).asText());

    /**
     * Fonction de sanitation comme pour les ID Grist
     */
    xml = xml.replace(/\{(.*?)\}/g, (match, key) => {
        if (key.startsWith('#') || key.startsWith('/')) {
            return match; // marqueurs de boucle {#Table}/{/Table}, à préserver
        }
        return `{${sanitizeKey(key)}}`;
    });

    zip.file(xmlFile, xml);
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