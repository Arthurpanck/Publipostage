/**
 * MODULE UI (vue)
 * Version : 2.0
 * Interface issue du handoff design "ergonomie et layout" :
 *  - modale d'instructions avec zone de dépôt du modèle ;
 *  - en-tête persistant 54px : titre | pastille d'alerte | actions | modèle+roue ;
 *  - responsive en CSS (paliers dans style.css) : sans titre, libellés courts
 *    ("Ligne"/"Zip"), puis actions repliées dans le menu de la roue crantée ;
 *  - prévisualisation centrale permanente ;
 *  - toast discret pour les états transitoires (génération, erreurs).
 * Ce module ne connaît pas Grist : main.js (contrôleur) lui injecte des
 * handlers via initUi() et pilote son état via les fonctions ui*.
 */

const uiState = {
    hasTemplate: false,
    templateName: '',
    showModal: false,
    showGearMenu: false,
    showWarnMenu: false,
    unmatchedTags: [],   // balises sans colonne correspondante
    warningNotes: [],    // avertissements complémentaires (tables non liées...)
    canDownloadLine: false,
    canDownloadZip: false,
};

let uiHandlers = { onPickFile: null, onDownloadLine: null, onDownloadZip: null };
let uiToastTimer = null;

const $ = (id) => document.getElementById(id);

/* ── Initialisation (appelée par main.js) ───────────────────────── */
function initUi(handlers) {
    uiHandlers = handlers;

    // Sélection de fichier (input caché, déclenché par la modale ou la roue)
    $('templateFile').addEventListener('change', (e) => {
        const file = e.target.files[0];
        e.target.value = ''; // permet de re-choisir le même fichier
        if (file && uiHandlers.onPickFile) {
            uiHandlers.onPickFile(file);
        }
    });
    $('dropzone').addEventListener('click', () => $('templateFile').click());
    $('btnBrowse').addEventListener('click', (e) => { e.stopPropagation(); $('templateFile').click(); });

    // Modale
    $('modalClose').addEventListener('click', uiCloseModal);
    $('modalFooterBtn').addEventListener('click', uiCloseModal);

    // Actions de téléchargement (barre + menu roue crantée si widget étroit)
    $('btnSingle').addEventListener('click', () => uiHandlers.onDownloadLine());
    $('btnBulk').addEventListener('click', () => uiHandlers.onDownloadZip());
    $('gearDownloadLine').addEventListener('click', () => { uiState.showGearMenu = false; uiRender(); uiHandlers.onDownloadLine(); });
    $('gearDownloadZip').addEventListener('click', () => { uiState.showGearMenu = false; uiRender(); uiHandlers.onDownloadZip(); });

    // Roue crantée : menu ancré sous le bouton, fermeture au clic extérieur
    $('btnGear').addEventListener('click', (e) => {
        e.stopPropagation();
        uiState.showGearMenu = !uiState.showGearMenu;
        uiState.showWarnMenu = false;
        uiRender();
    });
    $('gearChangeTemplate').addEventListener('click', () => {
        uiState.showGearMenu = false;
        uiRender();
        $('templateFile').click(); // directement le sélecteur, sans écran intermédiaire
    });
    $('gearInstructions').addEventListener('click', () => {
        uiState.showGearMenu = false;
        uiOpenModal();
    });
    $('btnLoadTemplate').addEventListener('click', uiOpenModal);

    // Pastille d'alerte : liste déroulante des balises sans correspondance
    $('warnChip').addEventListener('click', (e) => {
        e.stopPropagation();
        uiState.showWarnMenu = !uiState.showWarnMenu;
        uiState.showGearMenu = false;
        uiRender();
    });
    $('warnMenu').addEventListener('click', (e) => e.stopPropagation());

    document.addEventListener('click', () => {
        if (uiState.showGearMenu || uiState.showWarnMenu) {
            uiState.showGearMenu = false;
            uiState.showWarnMenu = false;
            uiRender();
        }
    });

    // Ajustement de l'aperçu à chaque redimensionnement, différé d'une frame
    // pour éviter la boucle "ResizeObserver loop"
    let fitRaf = null;
    const refit = () => {
        if (fitRaf) cancelAnimationFrame(fitRaf);
        fitRaf = requestAnimationFrame(uiFitPreview);
    };
    if (typeof ResizeObserver !== 'undefined') {
        new ResizeObserver(refit).observe($('previewScroll'));
    } else {
        window.addEventListener('resize', refit);
    }

    uiRender();
}

/* ── API pilotée par le contrôleur ──────────────────────────────── */
function uiSetTemplate(name) {
    uiState.hasTemplate = !!name;
    uiState.templateName = name || '';
    uiRender();
}

function uiOpenModal() {
    uiState.showModal = true;
    uiRender();
}

function uiCloseModal() {
    uiState.showModal = false;
    uiRender();
}

function uiEnableActions(line, zip) {
    uiState.canDownloadLine = line;
    uiState.canDownloadZip = zip;
    uiRender();
}

// tags : balises sans colonne correspondante ; notes : messages complémentaires
function uiSetWarnings(tags, notes) {
    uiState.unmatchedTags = tags || [];
    uiState.warningNotes = notes || [];
    if (uiState.unmatchedTags.length === 0 && uiState.warningNotes.length === 0) {
        uiState.showWarnMenu = false;
    }
    uiRender();
}

function uiShowPreview(visible) {
    $('previewScroll').hidden = !visible;
    $('previewEmpty').hidden = visible;
    if (visible) {
        uiFitPreview();
        // seconde passe après stabilisation de la mise en page (polices, images)
        requestAnimationFrame(uiFitPreview);
    }
}

// Ajuste la prévisualisation à la largeur disponible : la page garde ses
// proportions réelles (A4...) et est réduite par transform:scale pour tenir
// dans le widget, souvent affiché sur une demi-largeur d'écran. Le wrapper
// .preview-fit reçoit la taille APRÈS échelle (un transform ne modifie pas
// la boîte de mise en page, sinon la zone de scroll garderait la taille
// naturelle). Recalculé à chaque redimensionnement.
function uiFitPreview() {
    const scroll = $('previewScroll');
    const fit = $('previewFit');
    const page = $('preview-container');
    if (!scroll || !fit || !page || scroll.hidden) {
        return;
    }

    // PDF : l'iframe s'étire simplement sur toute la largeur disponible
    if (page.querySelector('iframe')) {
        fit.classList.add('pdf');
        fit.style.width = '';
        fit.style.height = '';
        page.style.transform = '';
        page.style.width = '';
        return;
    }
    fit.classList.remove('pdf');

    // mesure à l'échelle 1 (largeur naturelle de la plus large des pages)
    page.style.transform = '';
    page.style.width = '';
    fit.style.width = '';
    fit.style.height = '';
    let naturalW = 0;
    page.querySelectorAll('.docx_viewer').forEach((section) => {
        naturalW = Math.max(naturalW, section.offsetWidth);
    });
    if (!naturalW) {
        naturalW = page.offsetWidth;
    }
    if (!naturalW) {
        return;
    }
    const naturalH = page.getBoundingClientRect().height;

    const stylesZone = getComputedStyle(scroll);
    const avail = scroll.clientWidth
        - parseFloat(stylesZone.paddingLeft) - parseFloat(stylesZone.paddingRight);
    const scale = Math.min(1, avail / naturalW);

    page.style.width = naturalW + 'px';
    if (scale < 1) {
        page.style.transform = 'scale(' + scale + ')';
        fit.style.width = (naturalW * scale) + 'px';
        fit.style.height = (naturalH * scale) + 'px';
    }
}

function uiPreviewEmptyText(text) {
    $('previewEmptyText').textContent = text;
}

// Toast transitoire : type 'normal' | 'success' | 'error'
function uiToast(message, type) {
    const toast = $('toast');
    toast.textContent = message;
    toast.className = 'toast ' + (type === 'success' || type === 'error' ? type : '');
    toast.hidden = false;
    if (uiToastTimer) clearTimeout(uiToastTimer);
    if (type !== 'normal') { // 'normal' = état en cours, remplacé par l'issue
        uiToastTimer = setTimeout(() => { toast.hidden = true; }, type === 'error' ? 8000 : 3500);
    }
}

/* ── Rendu ──────────────────────────────────────────────────────── */
function uiRender() {
    const s = uiState;
    const hasWarnings = s.hasTemplate && (s.unmatchedTags.length > 0 || s.warningNotes.length > 0);

    // Modale
    $('modalOverlay').hidden = !s.showModal;
    $('modalClose').hidden = !s.hasTemplate;
    if (s.hasTemplate) {
        $('dropzoneMain').textContent = s.templateName;
        $('dropzoneSub').textContent = 'Cliquer pour changer de modèle';
        $('modalFooterBtn').textContent = 'Fermer';
        $('modalFooterBtn').className = 'btn btn-primary';
    } else {
        $('dropzoneMain').textContent = 'Déposer le modèle ici';
        $('dropzoneSub').textContent = '.docx ou .pdf';
        $('modalFooterBtn').textContent = 'Passer';
        $('modalFooterBtn').className = 'btn btn-muted';
    }

    // Pastille d'alerte
    $('warnSlot').hidden = !hasWarnings;
    $('warnSep').hidden = !hasWarnings;
    if (hasWarnings) {
        $('warnLabel').textContent = uiWarnLabel();
        $('warnMenu').hidden = !s.showWarnMenu;
        if (s.showWarnMenu) {
            uiFillWarnMenu();
        }
    }

    // Actions : en barre, ou dans le menu de la roue si le widget est étroit
    // (bascule en CSS, voir style.css)
    $('actionsInline').hidden = !s.hasTemplate;
    $('actionsSpacer').hidden = !s.hasTemplate;
    $('actionsSep').hidden = !s.hasTemplate;
    $('btnSingle').disabled = !s.canDownloadLine;
    $('btnBulk').disabled = !s.canDownloadZip;

    // Modèle + roue crantée / état sans modèle
    $('templateSlot').hidden = !s.hasTemplate;
    $('noTemplateSlot').hidden = s.hasTemplate;
    $('tplName').textContent = s.templateName;
    $('tplName').title = s.templateName;
    $('gearMenu').hidden = !s.showGearMenu;
}

// Libellé de la pastille : balises sans colonne, sinon autres alertes
function uiWarnLabel() {
    const tags = uiState.unmatchedTags.length;
    if (tags > 0) {
        return tags + ' balise' + (tags > 1 ? 's' : '') + ' sans correspondance';
    }
    const notes = uiState.warningNotes.length;
    return notes + ' alerte' + (notes > 1 ? 's' : '');
}

function uiFillWarnMenu() {
    const list = $('warnMenuList');
    list.textContent = '';
    for (const note of uiState.warningNotes) {
        const div = document.createElement('div');
        div.className = 'warn-menu-note';
        div.textContent = note;
        list.appendChild(div);
    }
    for (const tag of uiState.unmatchedTags) {
        const item = document.createElement('div');
        item.className = 'warn-menu-item';
        item.innerHTML = '<svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"></path><line x1="12" y1="9" x2="12" y2="13"></line><line x1="12" y1="17" x2="12.01" y2="17"></line></svg>';
        const code = document.createElement('code');
        code.textContent = '{' + tag + '}';
        item.appendChild(code);
        list.appendChild(item);
    }
}
