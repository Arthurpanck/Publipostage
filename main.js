/**
 * MAIN CONTROLLER
 * Version : 2.0
 * Objet : événements Grist, orchestration de la génération (docx/pdf) et
 * pilotage de la vue (ui.js). Ne manipule pas le DOM.
 */

const state = {
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

// Le modèle est rangé dans la table Publipostage_Modeles_NE_PAS_SUPPRIMER ;
// les options du widget le désignent (voir template-store.js). Grist les
// envoie au démarrage puis à chaque modification, y compris quand un autre
// utilisateur change le modèle. Les notifications sont traitées l'une après
// l'autre. Avant la première lecture, les options des widgets du document
// sont remises d'aplomb (healBeforeReading), ce qui peut modifier les nôtres :
// la lecture utilise les dernières reçues.
let optionsSeen = false;
let optionsChain = Promise.resolve();
grist.onOptions((options) => {
    noteOptions(options);
    const first = !optionsSeen;
    optionsSeen = true;
    const apply = first
        ? healBeforeReading().then(() => applyOptions(storeOptions, true))
        : optionsChain.then(() => applyOptions(options, false));
    optionsChain = apply.catch(e => console.error(e));
    if (first) prepareTemplateTable();
});

// Ligne sélectionnée (onRecord) ou lignes du widget (onRecords) : une liaison
// ou un filtre peut changer les lignes sans déplacer le curseur.
function receiveRows(change, mappings) {
    state.filenameColumn = mappings?.filename || '';
    Object.assign(state, change);
    updateActionsState();
    if (state.currentRecord && state.templateBuffer) return updatePreview();
}
grist.onRecord((record, mappings) => receiveRows({ currentRecord: record }, mappings), { includeColumns: 'normal' });
grist.onRecords((records, mappings) => receiveRows({ allRecords: records }, mappings), { includeColumns: 'normal' });

function getTemplateType(name) {
    const extension = name.split('.').pop().toLowerCase();
    return extension === 'docx' || extension === 'pdf' ? extension : null;
}

// Tant que l'accès complet n'a pas été accordé au widget (il vient d'être
// ajouté ou re-sélectionné), Grist refuse toute lecture et toute écriture.
const NO_ACCESS = "Ce widget a besoin de l'accès complet au document pour lire et enregistrer son modèle : ouvrez le panneau de configuration du widget (colonne de droite de Grist) et cliquez sur « Accepter ».";
const isAccessDenied = e => /access not granted/i.test(String(e && e.message));

// Upload du Template (fichier choisi via la modale ou la roue crantée).
// Tant qu'il est en cours, les notifications d'options sont ignorées
// (applyOptions) : le modèle choisi ne doit pas être remplacé par l'ancien.
let uploading = 0;
async function handleTemplateUpload(file) {
    const type = getTemplateType(file.name);
    if (!type) {
        uiToast(/\.doc$/i.test(file.name)
            ? "Le format .doc n'est pas pris en charge. Enregistrez le fichier en .docx."
            : "Format non supporté. Utilisez .docx ou .pdf", "error");
        return;
    }

    uiToast("Enregistrement du modèle...", "normal");

    uploading++;
    try {
        const buffer = await readFileAsBuffer(file);
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
            uiPersistBanner("Modèle utilisé pour cette session seulement : plusieurs widgets Publipostage de cette table ont les mêmes réglages (widgets neufs ou dupliqués), et ce document réserve la modification de sa structure aux propriétaires. Demandez à un propriétaire de charger un modèle dans ce widget.");
        }
    } catch (err) {
        console.error(err);
        if (isAccessDenied(err)) {
            uiToast("Modèle non enregistré : accès au document non accordé", "error");
            uiPersistBanner(NO_ACCESS);
            return;
        }
        uiToast("Erreur lors de la sauvegarde : " + err.message, "error");
        if (state.templateName === file.name) {
            uiPersistBanner("Modèle utilisé pour cette session seulement : son enregistrement dans le document a échoué.");
        }
    } finally {
        uploading--;
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

// Prévisualisation du document pour la ligne sélectionnée. onRecord et
// onRecords peuvent la relancer coup sur coup : le rendu est préparé hors de
// la page (uiPreparePreview) et seul l'aperçu le plus récent est affiché.
let previewRun = 0;
async function updatePreview() {
    const run = ++previewRun;
    const batch = newBatch();
    try {
        const blob = await dispatchGeneration(state.currentRecord, batch);
        const show = await uiPreparePreview(blob, state.templateType);
        if (run !== previewRun) {
            return;
        }
        show();
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
// d'instructions s'ouvre au premier chargement. Un modèle choisi par
// l'utilisateur pendant la lecture garde la priorité.
let missingShown = null;
async function applyOptions(options, first) {
    if (uploading) return;
    const loaded = state.templateHash;
    let template;
    try {
        template = await readTemplateRetrying(options, hash => !!hash && hash === loaded);
    } catch (e) {
        console.warn("Erreur chargement", e);
        if (isAccessDenied(e)) uiPersistBanner(NO_ACCESS);
        else uiToast("Impossible de charger le modèle : " + e.message, "error");
        return;
    }
    if (uploading || state.templateHash !== loaded) return;
    if (!template) {
        if (first && !state.templateBuffer) uiOpenModal();
        return;
    }
    if (template.missing) {
        if (missingShown === template.name || state.templateBuffer) return;
        missingShown = template.name;
        uiToast(`Le modèle « ${template.name} » n'est plus disponible dans ce document (copie sans données, table ${TEMPLATE_TABLE} supprimée ou masquée par une règle d'accès…) : chargez-le à nouveau.`, "error");
        uiOpenModal();
        return;
    }
    if (!template.buffer) return; // déjà chargé
    const type = getTemplateType(template.name);
    if (!type) {
        uiToast('Format du modèle enregistré non supporté.', 'error');
        return;
    }
    updateTemplateState(template.buffer, template.name, type, template.hash);
    uiPersistBanner(null);
}

// Une lecture qui échoue (réseau, Grist occupé) est retentée avant d'être
// signalée : le modèle est toujours dans le document.
async function readTemplateRetrying(options, skip) {
    for (let attempt = 0; ; attempt++) {
        try {
            return await readTemplate(options, skip);
        } catch (e) {
            if (attempt === 2 || isAccessDenied(e)) throw e;
            await new Promise(resolve => setTimeout(resolve, 1000 * (attempt + 1)));
        }
    }
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
