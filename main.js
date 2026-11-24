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

// Le modèle est rangé dans la table Publipostage_Modeles ; les options du
// widget le désignent (voir template-store.js). Grist les envoie au démarrage
// puis à chaque modification, y compris quand un autre utilisateur change le
// modèle. Les notifications sont traitées l'une après l'autre.
let optionsSeen = false;
let optionsChain = Promise.resolve();
grist.onOptions((options) => {
    noteOptions(options);
    const first = !optionsSeen;
    optionsSeen = true;
    optionsChain = optionsChain.then(() => applyOptions(options, first)).catch(e => console.error(e));
    if (first) prepareTableForEditors();
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

        const hash = await templateHash(buffer);
        updateTemplateState(buffer, file.name, type, hash);
        uiCloseModal();

        const status = await storeTemplate({ name: file.name, hash, buffer });
        if (status === 'saved' || status === 'row') {
            uiPersistBanner(null);
            uiToast("Modèle enregistré dans le document", "success");
        } else if (status === 'readonly') {
            uiToast("Lecture seule : modèle utilisé pour cette session uniquement", "error");
        } else if (status === 'owners') {
            uiPersistBanner("Modèle utilisé pour cette session seulement : ce document réserve la modification de sa structure aux propriétaires. Un propriétaire doit ouvrir ce widget une fois ; ensuite, les éditeurs pourront enregistrer leurs modèles.");
        } else if (status === 'section') {
            uiPersistBanner("Modèle utilisé pour cette session seulement : plusieurs widgets Publipostage de cette table n'ont pas encore de modèle, et ce document réserve la modification de sa structure aux propriétaires. Demandez à un propriétaire de charger un premier modèle dans ce widget.");
        }
    } catch (err) {
        console.error(err);
        uiToast("Erreur lors de la sauvegarde : " + err.message, "error");
        if (state.templateName === file.name) {
            uiPersistBanner("Modèle utilisé pour cette session seulement : son enregistrement dans le document a échoué.");
        }
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
        return await generatePdfBlob(data, state.templateBuffer, await loadRenderInfo(batch));
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

// Charge le modèle désigné par les options. Sans modèle, la modale
// d'instructions s'ouvre au premier chargement.
let missingShown = null;
async function applyOptions(options, first) {
    let template;
    try {
        template = await readTemplate(options, hash => !!hash && hash === state.templateHash);
    } catch (e) {
        console.warn("Erreur chargement", e);
        uiToast("Impossible de charger le modèle : " + e.message, "error");
        return;
    }
    if (!template) {
        if (first && !state.templateBuffer) uiOpenModal();
        return;
    }
    if (template.missing) {
        if (missingShown === template.name || state.templateBuffer) return;
        missingShown = template.name;
        uiToast(`Le modèle « ${template.name} » n'est plus disponible dans ce document (copie sans données, table ${TEMPLATE_TABLE} supprimée…) : chargez-le à nouveau.`, "error");
        uiOpenModal();
        return;
    }
    if (!template.buffer || template.hash === state.templateHash) return;
    const type = getTemplateType(template.name);
    if (!type) {
        uiToast('Format du modèle enregistré non supporté.', 'error');
        return;
    }
    updateTemplateState(template.buffer, template.name, type, template.hash);
    uiPersistBanner(null);
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
