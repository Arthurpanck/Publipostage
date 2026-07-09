/**
 * MAIN CONTROLLER
 * Version : 2.0
 * Objet : événements Grist, orchestration de la génération (docx/pdf) et
 * pilotage de la vue (ui.js). Ne manipule pas le DOM directement.
 */

let state = {
    currentRecord: null,
    allRecords: [],
    templateBuffer: null,
    templateType: null,
    templateName: null
};

grist.ready({
    requiredAccess: 'full'
});

initUi({
    onPickFile: handleTemplateUpload,
    onDownloadLine: downloadSingle,
    onDownloadZip: downloadBulk,
});

// Lancement différé du chargement du template car plante parfois si pas de timeout
setTimeout(() => {
    loadSavedTemplate();
}, 500);

grist.onRecord(async (record) => {
    state.currentRecord = record;
    clearRelationsCache(); // les données liées ont pu changer
    updateActionsState();

    // MAJ de la preview si on clique sur un enregistrement
    if (state.currentRecord && state.templateBuffer) {
        await updatePreview();
    }
});

grist.onRecords((records) => {
    state.allRecords = records;
    clearRelationsCache(); // les données liées ont pu changer
    updateActionsState();
});

// Upload du Template (fichier choisi via la modale ou la roue crantée)
async function handleTemplateUpload(file) {
    // Validation extension
    let type = null;
    if (file.name.endsWith('.docx') || file.name.endsWith('.doc')) {
        type = 'docx';
    } else if (file.name.endsWith('.pdf')) {
        type = 'pdf';
    } else {
        uiToast("Format non supporté. Utilisez .docx ou .pdf", "error");
        return;
    }

    uiToast("Upload et sauvegarde du template...", "normal");

    try {
        const buffer = await readFileAsBuffer(file);

        const attachmentId = await uploadAttachmentToGrist(file);
        if (!attachmentId) {
            throw new Error("ID de fichier invalide reçu.");
        }

        // Sauvegarde des informations du template uploadé
        await grist.setOption('templateId', attachmentId);
        await grist.setOption('templateName', file.name);

        // Vérification si le fichier est bien sauvegardé
        const checkId = await grist.getOption('templateId');
        if (checkId != attachmentId) {
            uiToast("Attention, cliquer sur 'Enregistrer' en haut de la page !", "error");
            return;
        }

        updateTemplateState(buffer, file.name, type);
        uiCloseModal();
        uiToast("Modèle sauvegardé", "success");

    } catch (err) {
        console.error(err);
        uiToast("Erreur lors de la sauvegarde : " + err.message, "error");
    }
}

// Export de la ligne sélectionnée
async function downloadSingle() {
    if (!state.currentRecord || !state.templateBuffer) {
        return;
    }
    uiToast("Génération du document...", "normal");
    try {
        const blob = await dispatchGeneration(state.currentRecord);
        saveAs(blob, `Document_${state.currentRecord.id || 'export'}.${state.templateType}`);
        refreshWarnings();
        uiToast("Téléchargement terminé", "success");
    } catch (error) {
        console.error(error);
        uiToast("Erreur : " + error.message, "error");
    }
}

// Export en masse (ZIP)
async function downloadBulk() {
    if (!state.allRecords.length || !state.templateBuffer) {
        return;
    }
    uiToast(`Génération du ZIP (${state.allRecords.length} fichiers)...`, "normal");
    try {
        const zip = new JSZip();
        for (const row of state.allRecords) {
            if (row.id === 'new') {
                continue;
            }
            const fileName = `Doc_${row.id}.${state.templateType}`;
            const docBlob = await dispatchGeneration(row);
            zip.file(fileName, docBlob);
        }
        const content = await zip.generateAsync({type: "blob"});
        saveAs(content, "Publipostage.zip");
        refreshWarnings();
        uiToast("ZIP créé avec succès", "success");
    } catch (error) {
        console.error(error);
        uiToast("Erreur ZIP : " + error.message, "error");
    }
}

// --- LOGIQUE MÉTIER : nettoyage des clés + dispatch selon le type de modèle ---
async function dispatchGeneration(rawData) {
    // retrait des metadatas Grist
    const cleanData = {};
    for (const key in rawData) {
        if (!key.startsWith('__') && key !== 'id') {
            const cleanKey = sanitizeKey(key);
            cleanData[cleanKey] = rawData[key];
        }
    }

    // Ajout des colonnes masquées du widget (non reçues via grist.onRecord) :
    // la ligne complète est relue dans la table pour que TOUTES les colonnes
    // soient publipostables, même celles ajoutées après la création du widget.
    try {
        await completeParentData(cleanData, rawData.id);
    } catch (e) {
        console.warn("Complément des colonnes indisponible", e);
    }

    if (state.templateType === 'docx') {
        // Ajout des tables enfants liées si le modèle contient des balises
        // {Table.Colonne} (voir relations-tools.js). En cas d'échec, le
        // publipostage de base fonctionne toujours.
        try {
            await addChildTablesData(cleanData, rawData.id, state.templateBuffer);
        } catch (e) {
            console.warn("Publipostage relationnel indisponible", e);
        }
        return generateDocxBlob(cleanData, state.templateBuffer);
    } else if (state.templateType === 'pdf') {
        return await generatePdfBlob(cleanData, state.templateBuffer);
    }
}

// Prévisualisation du document pour la ligne sélectionnée
async function updatePreview() {
    const container = document.getElementById('preview-container');
    if (!container) {
        return;
    }

    try {
        const blob = await dispatchGeneration(state.currentRecord);
        if (state.templateType === 'docx') {
            await docx.renderAsync(blob, container, null, { className: "docx_viewer", inWrapper: true, ignoreWidth: false });
        } else if (state.templateType === 'pdf') {
            const pdfUrl = URL.createObjectURL(blob);
            container.innerHTML = `<iframe src="${pdfUrl}"></iframe>`;
        }
        uiShowPreview(true);
        refreshWarnings();
    } catch (e) {
        console.error("Erreur Preview :", e);
        uiShowPreview(false);
        uiPreviewEmptyText("Erreur de chargement de l'aperçu");
        uiToast("Erreur d'aperçu : " + e.message, "error");
    }
}

// Alimente la pastille "N balises sans correspondance" de l'en-tête
function refreshWarnings() {
    const tags = (typeof getUnknownTags === 'function') ? getUnknownTags() : [];
    const notes = (typeof getRelationsWarnings === 'function') ? getRelationsWarnings() : [];
    uiSetWarnings(tags, notes);
}

function sanitizeKey(keytoSanitize) {
    if (!keytoSanitize) {
        return "";
    }
    let sanitize = keytoSanitize.toString();
    sanitize = sanitize.normalize('NFKD');
    sanitize = sanitize.replace(/[\u0300-\u036f]/g, ""); // Supprime les accents
    sanitize = sanitize.replace(/[^a-zA-Z0-9_]+/g, "_"); // Remplace caractères spéciaux
    sanitize = sanitize.replace(/^_+/, ""); // Supprime _ au début
    return sanitize;
}

function updateActionsState() {
    const ready = state.templateBuffer !== null;
    uiEnableActions(ready && state.currentRecord !== null, ready && state.allRecords.length > 0);
}

// Récupère le template sauvegardé au chargement de la page.
// Modèle configuré -> on le charge sans modale ; sinon la modale
// d'instructions s'affiche automatiquement (premier chargement).
async function loadSavedTemplate() {
    try {
        const templateId = await grist.getOption('templateId');
        const templateName = await grist.getOption('templateName');

        if (templateId && templateName) {
            uiToast("Récupération du modèle...", "normal");
            let type = 'docx';
            if (templateName.endsWith('.pdf')) {
                type = 'pdf';
            }

            const buffer = await downloadAttachmentFromGrist(templateId);
            updateTemplateState(buffer, templateName, type);
            uiToast("Modèle chargé : " + templateName, "success");
        } else {
            uiOpenModal();
        }
    } catch (e) {
        console.warn("Erreur chargement", e);
        uiOpenModal();
    }
}

// Envoie le fichier à Grist (method POST)
async function uploadAttachmentToGrist(file) {
    const tokenInfo = await grist.docApi.getAccessToken({
        readOnly: false
    });

    const formData = new FormData();
    formData.set('upload', file, file.name);
    const response = await fetch(`${tokenInfo.baseUrl}/attachments?auth=${tokenInfo.token}`, {
        method: 'POST',
        body: formData,
        headers: {
            'X-Requested-With': 'XMLHttpRequest'
        }
    });
    if (!response.ok) {
        throw new Error('Echec Upload');
    }
    const ids = await response.json();
    return ids[0];
}

// Récupère le fichier de Grist (method GET)
async function downloadAttachmentFromGrist(attachmentId) {
    const tokenInfo = await grist.docApi.getAccessToken({ readOnly: true });
    const url = `${tokenInfo.baseUrl}/attachments/${attachmentId}/download?auth=${tokenInfo.token}`;

    const response = await fetch(url);
    if (!response.ok) {
        throw new Error('Impossible de récupérer le template');
    }
    return await response.arrayBuffer();
}

function readFileAsBuffer(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => resolve(e.target.result);
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

function updateTemplateState(buffer, name, type) {
    state.templateBuffer = buffer;
    state.templateType = type;
    state.templateName = name;

    clearRelationsCache();
    uiSetTemplate(name);
    updateActionsState();

    // Rafraichissement de l'aperçu
    if (state.currentRecord) {
        updatePreview();
    } else {
        uiShowPreview(false);
        uiPreviewEmptyText("Sélectionner une ligne pour voir l'aperçu");
    }
}
