/**
 * MODULE DOCX
 * Version : 1.5
 * maintainer : dieux.alexandre@gmail.com
 * Gestion des fichiers Word (.docx)
 */

// Génération du docx
function generateDocxBlob(data, buffer) {
    const zip = new PizZip(buffer);

    // Traduction des boucles {Table.Colonne} AVANT la sanitisation des clés
    // (sinon le "." serait remplacé par "_" et la traduction échouerait)
    try {
        transformDottedLoops(zip);
    } catch (e) {
        console.warn("ATTENTION : La transformation des boucles relationnelles a échoué", e);
    }

    try {
        sanitizeDocxXml(zip);
    } catch (e) {
        console.warn("ATTENTION : Le nettoyage automatique du XML a échoué", e);
    }

    let doc;
    try {
        doc = new window.docxtemplater(zip, {
            paragraphLoop: true,
            linebreaks: true,
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
    const splitTag = /(\{[^{}<]*)<\/w:t><\/w:r>(?:<w:bookmark(?:Start|End)[^>]*\/>)*<w:r\b[^>]*>(?:<w:rPr>[\s\S]*?<\/w:rPr>)?<w:t[^>]*>/g;
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