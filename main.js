/**
 * MAIN CONTROLLER
 * Version : 1.5
 * maintainer : dieux.alexandre@gmail.com
 * Objet : Gestion de l'interface, des événements Grist et de l'orchestration des événements
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

// Lancement différé du chargement du template car plante parfois si pas de timeout
setTimeout(() => {
    loadSavedTemplate();
}, 500);

grist.onRecord(async (record) => {
    state.currentRecord = record;
    updateUiState();

    // MAJ de la preview si on clique sur un enregistrement
    if (state.currentRecord && state.templateBuffer) {
        await updatePreview();
    }
});

grist.onRecords((records) => {
    state.allRecords = records;
    updateUiState();
});

document.getElementById('btnToggleUpload').addEventListener('click', function() {
    toggleUploadSection(true);
});

function toggleUploadSection(showUpload) {
    const header = document.getElementById('templateHeader');
    const container = document.getElementById('uploadContainer');
    const headerName = document.getElementById('headerFileName');

    if (showUpload) {
        header.style.display = 'none';
        container.classList.remove('hidden');
    } else {
        header.style.display = 'flex';
        container.classList.add('hidden');

        if(state.templateName) {
            headerName.textContent = state.templateName;
        }
    }
}

// Upload du Template
document.getElementById('templateFile').addEventListener('change', async function(e) {
    const file = e.target.files[0];
    if (!file) {
        return;
    }

    // Validation extension
    let type = null;
    if (file.name.endsWith('.docx') || file.name.endsWith('.doc')) {
        type = 'docx';
    } else if (file.name.endsWith('.pdf')) {
        type = 'pdf';
    } else {
        setStatus("Format non supporté. Utilisez .docx ou .pdf", "error");
        return;
    }

    setStatus("Upload et sauvegarde du template...", "normal");

    // Gestion du template
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
            setStatus("Attention, cliquer sur 'Enregistrer' en haut de la page !", "error");
            return;
        }

        updateTemplateState(buffer, file.name, type);
        setStatus("Template sauvegardé", "success");

    } catch (err) {
        console.error(err);
        setStatus("Erreur lors de la sauvegarde : " + err.message, "error");
    }
});

// Single export
document.getElementById('btnSingle').addEventListener('click', async () => {
    if (!state.currentRecord || !state.templateBuffer) {
        return;
    }
    setStatus("Génération du document", "normal");
    try {
        const blob = await dispatchGeneration(state.currentRecord);
        saveAs(blob, `Document_${state.currentRecord.id || 'export'}.${state.templateType}`);
        setStatus("Téléchargement terminé", "success");
    } catch (error) {
        console.error(error);
        setStatus("Erreur: " + error.message, "error");
    }
});

// Mass export
document.getElementById('btnBulk').addEventListener('click', async () => {
    if (!state.allRecords.length || !state.templateBuffer) {
        return;
    }
    setStatus(`Génération du ZIP (${state.allRecords.length} fichiers)`, "normal");
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
        setStatus("ZIP créé avec succès", "success");
    } catch (error) {
        console.error(error);
        setStatus("Erreur ZIP: " + error.message, "error");
    }
});

// --- LOGIQUE MÉTIER - Gestion des clés - dispatcher en fonction du type de template - prévisualisation - sanitizer ---
async function dispatchGeneration(rawData) {
    // retrait des metadatas Grist
    const cleanData = {};
    for (const key in rawData) {
        if (!key.startsWith('__') && key !== 'id') {
            const cleanKey = sanitizeKey(key);
            cleanData[cleanKey] = rawData[key];
        }
    }

    if (state.templateType === 'docx') {
        return generateDocxBlob(cleanData, state.templateBuffer);
    } else if (state.templateType === 'pdf') {
        return await generatePdfBlob(cleanData, state.templateBuffer);
    }
}

// Prévisualisation documents
async function updatePreview() {
    const container = document.getElementById('preview-container');
    if(!container) {
        return;
    }

    try {
        const blob = await dispatchGeneration(state.currentRecord);
        if (state.templateType === 'docx') {
            docx.renderAsync(blob, container, null, { className: "docx_viewer", inWrapper: true, ignoreWidth: false });
        } else if (state.templateType === 'pdf') {
            const pdfUrl = URL.createObjectURL(blob);
            container.innerHTML = `<iframe src="${pdfUrl}" width="100%" height="100%" style="border:none;"></iframe>`;
        }
    } catch (e) {
        console.error("Erreur Preview :", e);
        container.innerHTML = `<div style="color:red; padding:20px">Erreur de chargement de l'aperçu : ${e.message}</div>`;
    }
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

function updateUiState() {
    const ready = state.templateBuffer !== null;
    const hasRecord = state.currentRecord !== null;
    const hasRecords = state.allRecords.length > 0;

    const btnSingle = document.getElementById('btnSingle');
    const btnBulk = document.getElementById('btnBulk');
    if(btnSingle) {
        btnSingle.disabled = !ready || !hasRecord;
    }
    if(btnBulk) {
        btnBulk.disabled = !ready || !hasRecords;
    }
}

function setStatus(msg, type) {
    const statusElement = document.getElementById('status');
    if(statusElement) {
        statusElement.textContent = msg;
        statusElement.className = 'status ' + (type || '');
    }
}

// Récupère le template sauvegardé au chargement de la page
async function loadSavedTemplate() {
    try {
        const templateId = await grist.getOption('templateId');
        const templateName = await grist.getOption('templateName');

        if (templateId && templateName) {
            setStatus("Récupération du template", "normal");
            let type = 'docx';
            if (templateName.endsWith('.pdf')) {
                type = 'pdf';
            }

            const buffer = await downloadAttachmentFromGrist(templateId);
            updateTemplateState(buffer, templateName, type);
            setStatus("Template chargé : " + templateName, "success");

            toggleUploadSection(false);
        } else {
            setStatus("Aucun template configuré.", "normal");

            toggleUploadSection(true);
        }
    } catch (e) {
        console.warn("Erreur chargement", e);
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

    // Mise à jour de l'affichage (affichage du nom de fichier sauvegardé
    const nameEl = document.getElementById('fileName');
    if(nameEl) {
        nameEl.textContent = name + " (Sauvegardé)";
    }

    updateUiState();

    // Rafraichissement de l'aperçu
    if (state.currentRecord) {
        updatePreview();
    }

    toggleUploadSection(false);
}