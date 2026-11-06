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
    templateName: null,
    templateHash: null,
    filenameColumn: ''
};

grist.ready({
    requiredAccess: 'full',
    columns: [{ name: 'filename', title: 'Nom du document', optional: true,
        description: 'Colonne utilisée pour nommer le document et chaque fichier du ZIP. Vide : nom automatique.' }]
});

initUi({
    onPickFile: handleTemplateUpload,
    onDownloadLine: downloadSingle,
    onDownloadZip: downloadBulk,
});

// Le modèle vit dans les options du widget (voir template-store.js). Grist
// les envoie au démarrage puis à chaque modification, y compris quand un autre
// utilisateur change le modèle.
let optionsSeen = false;
grist.onOptions(async (options) => {
    noteOptions(options);
    const first = !optionsSeen;
    optionsSeen = true;
    await applyOptions(options, first);
});

grist.onRecord(async (record, mappings) => {
    state.filenameColumn = mappings?.filename || '';
    state.currentRecord = record;
    updateActionsState();

    // MAJ de la preview si on clique sur un enregistrement
    if (state.currentRecord && state.templateBuffer) {
        await updatePreview();
    }
}, { includeColumns: 'normal' });

grist.onRecords(async (records, mappings) => {
    state.filenameColumn = mappings?.filename || '';
    state.allRecords = records;
    updateActionsState();
    // Une liaison ou un filtre peut changer les lignes sans déplacer le curseur.
    if (state.currentRecord && state.templateBuffer) await updatePreview();
}, { includeColumns: 'normal' });

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

    uiToast("Enregistrement du modèle...", "normal");

    try {
        const buffer = await readFileAsBuffer(file);
        if (buffer.byteLength > TEMPLATE_MAX_BYTES) {
            uiToast(`Modèle trop lourd (${Math.ceil(buffer.byteLength / 1048576)} Mo) : 8 Mo au plus. Réduisez les images du document.`, "error");
            return;
        }
        // Valider avant de remplacer le modèle enregistré.
        if (type === 'docx') compileDocxTemplate(buffer);
        else await PDFLib.PDFDocument.load(buffer);

        const options = await templateOptions(await grist.getOptions(), buffer, file.name);
        updateTemplateState(buffer, file.name, type, options.templateHash);
        uiCloseModal();

        const status = await persistOptions(options);
        if (status === 'saved') {
            uiPersistBanner(null);
            uiToast("Modèle enregistré dans le document", "success");
        } else if (status === 'readonly') {
            uiToast("Lecture seule : modèle utilisé pour cette session uniquement", "error");
        } else if (status === 'owners') {
            uiPersistBanner("Modèle utilisé pour cette session seulement : ce document réserve la modification de sa structure aux propriétaires. Demandez à un propriétaire de charger le modèle, ou d'autoriser les éditeurs à modifier la structure (règles d'accès).");
        } else {
            watchPendingTemplate(options.templateHash);
        }
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
        const row = state.currentRecord;
        const fileName = await exportFilename(row, state.filenameColumn, state.templateType, 'Document');
        const blob = await dispatchGeneration(row, batch);
        saveAs(blob, fileName);
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
        const usedNames = new Set();
        const column = state.filenameColumn;
        const type = state.templateType;
        for (const row of batch.viewRows) {
            if (row.id === 'new') {
                continue;
            }
            const fileName = uniqueExportFilename(await exportFilename(row, column, type, 'Doc'), usedNames);
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

// --- LOGIQUE MÉTIER : dispatch selon le type de modèle ---
// La ligne est complétée des colonnes que Grist n'aurait pas transmises
// (voir completeRow) ; les générateurs normalisent les clés.
async function dispatchGeneration(received, batch) {
    const record = await completeRow(received, batch);
    const data = rowData(record); // retrait des métadonnées Grist

    if (state.templateType === 'docx') {
        // Ajout des tables liées si le modèle contient des balises
        // {Table.Colonne} (voir relations-tools.js)
        await addLinkedTables(data, record, getReferencedTables(state.templateBuffer), batch);
        // Types et couleurs des colonnes Grist (voir render-tools.js)
        const renderInfo = await loadRenderInfo(batch);
        await loadAttachments(batch, data, renderInfo);
        return generateDocxBlob(data, state.templateBuffer, batch.unknownTags, renderInfo);
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

// Charge le modèle décrit par les options. Sans modèle, la modale
// d'instructions s'ouvre au premier chargement.
async function applyOptions(options, first) {
    const saved = templateFromOptions(options);
    if (!saved) {
        if (first && !state.templateBuffer) uiOpenModal();
        return;
    }
    if (saved.hash && saved.hash === state.templateHash) return;
    if (saved.legacyId && saved.legacyId === legacyLoaded) return;
    const type = getTemplateType(saved.name);
    if (!type) {
        uiToast('Format du modèle enregistré non supporté.', 'error');
        return;
    }
    if (saved.buffer) {
        updateTemplateState(saved.buffer, saved.name, type, saved.hash);
        uiPersistBanner(null);
        return;
    }
    await loadLegacyTemplate(options, saved, type);
}

// Ancien format : le modèle était une pièce jointe que Grist finit par
// supprimer. Il est intégré aux options dès qu'un éditeur ouvre le widget ;
// pour un lecteur, il est chargé une fois par session.
let legacyLoaded = null;
async function loadLegacyTemplate(options, saved, type) {
    legacyLoaded = saved.legacyId;
    let buffer;
    try {
        uiToast("Récupération du modèle...", "normal");
        buffer = await downloadAttachmentFromGrist(saved.legacyId);
    } catch (e) {
        console.warn("Erreur chargement", e);
        uiToast(`Le modèle « ${saved.name} » n'est plus disponible dans Grist (pièce jointe supprimée ou inaccessible) : chargez-le à nouveau.`, "error");
        uiOpenModal();
        return;
    }
    const upgraded = await templateOptions(options, buffer, saved.name);
    updateTemplateState(buffer, saved.name, type, upgraded.templateHash);
    uiToast("Modèle chargé : " + saved.name, "success");
    if (!isReadonly()) await persistOptions(upgraded, { fallback: false });
}

// Écriture directe impossible : le bandeau rappelle de cliquer sur
// « Enregistrer » tant que le modèle n'est pas dans le document.
let pendingWatch = null;
function watchPendingTemplate(hash) {
    uiPersistBanner("Modèle pas encore enregistré dans le document : cliquez sur « Enregistrer » (Save) en haut du widget, sinon il sera perdu au prochain chargement.");
    clearInterval(pendingWatch);
    pendingWatch = setInterval(async () => {
        if (hash !== state.templateHash) return clearInterval(pendingWatch);
        if (await isTemplateSaved(hash).catch(() => false)) {
            clearInterval(pendingWatch);
            uiPersistBanner(null);
            uiToast("Modèle enregistré dans le document", "success");
        }
    }, 3000);
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

function updateTemplateState(buffer, name, type, hash = null) {
    state.templateBuffer = buffer;
    state.templateType = type;
    state.templateName = name;
    state.templateHash = hash;

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

// Le mapping officiel donne l’identifiant réel de la colonne choisie dans Grist.
function exportFilename(row, column, type, prefix) {
    return buildExportFilename(column ? row[column] : undefined, `${prefix}_${row.id || 'export'}`, type);
}
