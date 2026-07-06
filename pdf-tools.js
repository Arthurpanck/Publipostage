/**
 * MODULE PDF
 * Version : 1.5
 * maintainer : dieux.alexandre@gmail.com
 * Gestion des fichiers PDF
 */

async function generatePdfBlob(data, buffer) {
    const { PDFDocument } = PDFLib;
    const pdfDoc = await PDFDocument.load(buffer);
    const form = pdfDoc.getForm();

    const fieldMap = {};
    const allFields = form.getFields();

    allFields.forEach(field => {
        let realName = field.getName();

        // Décodage Hexadécimal pour les accents PDF (par exemple #C3#A9 -> é)
        try {
            if (realName.indexOf('#') !== -1) {
                realName = decodeURIComponent(realName.replace(/#/g, '%'));
            }
        } catch (e) {}

        const cleanName = sanitizeKey(realName);
        fieldMap[cleanName] = field;
    });

    // Remplissage
    for (const [key, value] of Object.entries(data)) {
        const field = fieldMap[key];
        const textValue = value === null || value === undefined ? '' : String(value);

        if (field) {
            try {
                if (field.constructor.name === 'PDFTextField') {
                    field.setText(textValue);
                } else if (field.constructor.name === 'PDFCheckBox') {
                    if (value === true || value === 'true' || value === 1) field.check();
                    else field.uncheck();
                }
            } catch (e) {
                console.error(`Erreur champ PDF "${key}"`, e);
            }
        }
    }

    form.flatten();
    const pdfBytes = await pdfDoc.save();
    return new Blob([pdfBytes], { type: "application/pdf" });
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