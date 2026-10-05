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

// Toutes les colonnes de la table, y compris celles masquées dans le widget :
// elles restent publipostables sans avoir à les afficher.
const RECORD_OPTIONS = { includeColumns: 'normal' };

grist.onRecord(async (record) => {
    state.currentRecord = record;
    updateActionsState();

    // MAJ de la preview si on clique sur un enregistrement
    if (state.currentRecord && state.templateBuffer) {
        await updatePreview();
    }
}, RECORD_OPTIONS);

grist.onRecords(async (records) => {
    state.allRecords = records;
    updateActionsState();
    // Une liaison ou un filtre peut changer les lignes sans déplacer le curseur.
    if (state.currentRecord && state.templateBuffer) await updatePreview();
}, RECORD_OPTIONS);

function getTemplateType(name) {
    const extension = name.split('.').pop().toLowerCase();
    return extension === 'docx' || extension === 'pdf' ? extension : null;
}

// Upload du Template (fichier choisi via la modale ou la roue crantée)
async function handleTemplateUpload(file) {
    const type = getTemplateType(file.name);
    if (!type) {
        uiToast(/\.doc$/i.test(file.name)
            ? "Le format .doc n'est pas pris en charge. Enregistrez le fichier en .docx."
            : "Format non supporté. Utilisez .docx ou .pdf", "error");
        return;
    }

    uiToast("Upload et sauvegarde du template...", "normal");

    try {
        const buffer = await readFileAsBuffer(file);
        // Valider avant tout upload ou changement des options persistées.
        if (type === 'docx') compileDocxTemplate(buffer);
        else await PDFLib.PDFDocument.load(buffer);

        const attachmentId = await uploadAttachmentToGrist(file);
        if (!attachmentId) {
            throw new Error("ID de fichier invalide reçu.");
        }

        // Sauvegarde des informations du template uploadé
        await grist.setOptions({
            ...await grist.getOptions(),
            templateId: attachmentId,
            templateName: file.name,
        });

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
    const batch = newBatch();
    try {
        const blob = await dispatchGeneration(state.currentRecord, batch);
        saveAs(blob, `Document_${state.currentRecord.id || 'export'}.${state.templateType}`);
        showWarnings(batch);
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
    const batch = newBatch();
    uiToast(`Génération du ZIP (${batch.viewRows.length} fichiers)...`, "normal");
    try {
        const zip = new JSZip();
        for (const row of batch.viewRows) {
            if (row.id === 'new') {
                continue;
            }
            const fileName = `Doc_${row.id}.${state.templateType}`;
            const docBlob = await dispatchGeneration(row, batch);
            zip.file(fileName, docBlob);
        }
        const content = await zip.generateAsync({type: "blob"});
        saveAs(content, "Publipostage.zip");
        showWarnings(batch);
        uiToast("ZIP créé avec succès", "success");
    } catch (error) {
        console.error(error);
        uiToast("Erreur ZIP : " + error.message, "error");
    }
}

// Un lot = un aperçu, un document ou un ZIP. Les tables lues, les lignes du
// widget et les alertes lui appartiennent : rien à invalider entre deux lots,
// et un événement Grist reçu pendant un ZIP ne modifie pas le lot en cours.
function newBatch() {
    return { cache: new Map(), viewRows: state.allRecords, unknownTags: new Set(), notes: new Set() };
}

// --- LOGIQUE MÉTIER : nettoyage des clés + dispatch selon le type de modèle ---
async function dispatchGeneration(record, batch) {
    const data = rowData(record); // retrait des métadonnées Grist

    if (state.templateType === 'docx') {
        // Ajout des tables liées si le modèle contient des balises
        // {Table.Colonne} (voir data-tools.js)
        await addLinkedTables(data, record, getReferencedTables(state.templateBuffer), batch);
        return generateDocxBlob(data, state.templateBuffer, batch.unknownTags);
    } else if (state.templateType === 'pdf') {
        return await generatePdfBlob(data, state.templateBuffer);
    }
}

// Prévisualisation du document pour la ligne sélectionnée
let previewRun = 0;
async function updatePreview() {
    const container = document.getElementById('preview-container');
    if (!container) {
        return;
    }

    // onRecord et onRecords peuvent relancer l'aperçu coup sur coup : le rendu
    // se fait hors du DOM et seul l'aperçu le plus récent est affiché.
    const run = ++previewRun;
    const batch = newBatch();
    try {
        const blob = await dispatchGeneration(state.currentRecord, batch);
        const page = document.createElement('div');
        if (state.templateType === 'docx') {
            await docx.renderAsync(blob, page, null, { className: "docx_viewer", inWrapper: true, ignoreWidth: false });
        } else if (state.templateType === 'pdf') {
            const pdfUrl = URL.createObjectURL(blob);
            page.innerHTML = `<iframe src="${pdfUrl}"></iframe>`;
        }
        if (run !== previewRun) {
            return;
        }
        container.replaceChildren(...page.childNodes);
        uiShowPreview(true);
        showWarnings(batch);
    } catch (e) {
        if (run !== previewRun) {
            return;
        }
        console.error("Erreur Preview :", e);
        uiShowPreview(false);
        uiPreviewEmptyText("Erreur de chargement de l'aperçu");
        uiToast("Erreur d'aperçu : " + e.message, "error");
    }
}

// Alimente la pastille d'alerte de l'en-tête avec les alertes du lot
function showWarnings(batch) {
    uiSetWarnings([...batch.unknownTags], [...batch.notes]);
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
            const type = getTemplateType(templateName);
            if (!type) throw new Error('Format du modèle sauvegardé non supporté.');

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
