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

// Nettoyage des docxs
function sanitizeDocxXml(zip) {
    const xmlFile = "word/document.xml";
    if (!zip.file(xmlFile)) {
        return;
    }

    let xml = zip.file(xmlFile).asText();

    // Suppression des balises de correction orthographique et de grammaire
    xml = xml.replace(/<w:proofErr[^>]*\/>/g, "");
    xml = xml.replace(/<w:lang[^>]*\/>/g, "");
    xml = xml.replace(/<w:noProof[^>]*\/>/g, "");
    // reconstruction des balises cassées
    xml = xml.replace(/<\/w:t><\/w:r><w:r[^>]*><w:t[^>]*>/g, "");

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