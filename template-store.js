/**
 * MODÈLE PERSISTANT
 * Objet : garder le modèle attaché au widget, pour tous les utilisateurs et
 * dans le temps, sans alourdir l'ouverture du document.
 *
 * Le modèle est une pièce jointe rangée dans la table Publipostage_Modeles
 * (table de données brutes, sans page) : une ligne par widget, repérée par le
 * numéro de sa section. Une pièce jointe rangée dans une cellule n'est pas
 * purgée par Grist, et seul le widget affiché la télécharge.
 *
 * Les options du widget ne gardent que le nom et l'empreinte du modèle : leur
 * changement prévient les autres utilisateurs (grist.onOptions) et Ctrl+Z
 * rétablit le modèle précédent. Le fichier lui-même n'y est plus : Grist
 * envoie les options de tous les widgets à chaque ouverture du document.
 *
 * Écrire les options demande le droit de modifier la structure ; ranger le
 * fichier dans la table demande seulement celui de modifier les données. Un
 * éditeur d'un document qui réserve la structure aux propriétaires enregistre
 * donc son modèle dans la table, dès qu'elle existe.
 *
 * Anciens formats repris (puis retirés des options par un utilisateur qui
 * peut les écrire) : modèle en base64 dans les options (templateData) et
 * pièce jointe non rangée (templateId).
 */

const TEMPLATE_MAX_BYTES = 8 * 1024 * 1024;
const TEMPLATE_TABLE = 'Publipostage_Modeles';
const TEMPLATE_COLUMNS = { Widget: 'Int', Modele: 'Text', Empreinte: 'Text', Fichier: 'Attachments' };

// Dernières options reçues de Grist (grist.onOptions) et attentes en cours
let storeOptions = null;
let storeWaiters = [];
let storeSection = null;

function noteOptions(options) {
    storeOptions = options || null;
    storeWaiters = storeWaiters.filter(w => !w.test(storeOptions) || (w.resolve(true), false));
}

function waitForOptions(test, timeoutMs) {
    if (test(storeOptions)) return Promise.resolve(true);
    return new Promise(resolve => {
        const waiter = { test, resolve };
        storeWaiters.push(waiter);
        setTimeout(() => { storeWaiters = storeWaiters.filter(w => w !== waiter); resolve(false); }, timeoutMs);
    });
}

/* ── Fichier ────────────────────────────────────────────────────── */

function base64ToBuffer(text) {
    const bin = atob(text);
    const bytes = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    return bytes.buffer;
}

async function templateHash(buffer) {
    const digest = await crypto.subtle.digest('SHA-256', buffer);
    return Array.from(new Uint8Array(digest).slice(0, 12), b => b.toString(16).padStart(2, '0')).join('');
}

// Options du widget pour un modèle : nom, empreinte et numéro du widget (qui
// départage deux widgets de la même table, même après une duplication)
function pointerOptions(previous, name, hash, widget) {
    const { templateId, templateData, sondePublipostage, ...rest } = previous || {};
    return { ...rest, templateName: name, templateHash: hash, templateWidget: widget };
}

// Modèle d'un ancien format encore présent dans les options
function legacyFromOptions(options) {
    if (!options || !options.templateName) return null;
    if (options.templateData) return { name: options.templateName, hash: options.templateHash, data: options.templateData };
    if (options.templateId) return { name: options.templateName, legacyId: options.templateId };
    return null;
}

const hasLegacy = options => !!options && !!(options.templateData || options.templateId);

async function uploadTemplateFile(buffer, name) {
    const token = await grist.docApi.getAccessToken({ readOnly: false });
    const form = new FormData();
    form.set('upload', new Blob([buffer]), name);
    const response = await fetch(`${token.baseUrl}/attachments?auth=${token.token}`, {
        method: 'POST', body: form, headers: { 'X-Requested-With': 'XMLHttpRequest' },
    });
    if (!response.ok) throw new Error(`Envoi du fichier refusé par Grist (${response.status})`);
    return (await response.json())[0];
}

async function downloadTemplateFile(attachmentId) {
    const token = await grist.docApi.getAccessToken({ readOnly: true });
    const response = await fetch(`${token.baseUrl}/attachments/${attachmentId}/download?auth=${token.token}`);
    if (!response.ok) throw new Error(`Pièce jointe ${attachmentId} : ${response.status}`);
    return response.arrayBuffer();
}

/* ── Section du widget ──────────────────────────────────────────── */

const isReadonly = () => new URL(location.href).searchParams.get('readonly') === 'true';
const isStructureDenied = e => /structure access rules/i.test(String(e && e.message));

function stableJson(value) {
    if (Array.isArray(value)) return `[${value.map(stableJson).join(',')}]`;
    if (value && typeof value === 'object') {
        return `{${Object.keys(value).sort().map(k => `${JSON.stringify(k)}:${stableJson(value[k])}`).join(',')}}`;
    }
    return JSON.stringify(value === undefined ? null : value);
}
const sameOptions = (a, b) => stableJson(a || null) === stableJson(b || null);

function parseSection(table, i) {
    let options = {}, view = {};
    try { options = JSON.parse(table.options[i] || '{}'); } catch (e) { /* options illisibles */ }
    try { view = JSON.parse(options.customView || '{}'); } catch (e) { /* widget sans définition */ }
    return { id: table.id[i], options, view };
}

function samePage(url) {
    try {
        const a = new URL(url, location.href);
        return a.origin + a.pathname === location.origin + location.pathname;
    } catch (e) {
        return false;
    }
}

// Sections personnalisées de la table du widget qui pourraient être celle-ci
async function candidateSections() {
    const tableId = await grist.getSelectedTableId();
    const tables = await grist.docApi.fetchTable('_grist_Tables');
    const tableRef = tables.id[tables.tableId.indexOf(tableId)];
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    const custom = sections.id.map((_, i) => i)
        .filter(i => sections.parentKey[i] === 'custom' && sections.tableRef[i] === tableRef)
        .map(i => parseSection(sections, i));
    const ours = custom.filter(s => samePage(s.view.url || (s.view.widgetDef && s.view.widgetDef.url) || ''));
    return ours.length ? ours : custom;
}

// Action qui remplace les options enregistrées d'une section
async function sectionAction(id, widgetOptions) {
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    const section = parseSection(sections, sections.id.indexOf(id));
    section.view.widgetOptions = widgetOptions;
    section.options.customView = JSON.stringify(section.view);
    return ['UpdateRecord', '_grist_Views_section', id, { options: JSON.stringify(section.options) }];
}

async function writeSection(id, widgetOptions) {
    await grist.docApi.applyUserActions([await sectionAction(id, widgetOptions)]);
}

// Retrouve la section du widget : ses options enregistrées sont celles que
// Grist nous a transmises. Si plusieurs sections ont les mêmes (deux widgets
// neufs sur la même table), on marque l'une puis l'autre et on regarde
// laquelle Grist nous renvoie : il faut pouvoir écrire les options, d'où
// { probe: false } quand on se contente de lire.
async function ownSection({ probe = true } = {}) {
    if (storeSection) return storeSection;
    const candidates = (await candidateSections()).filter(s => sameOptions(s.view.widgetOptions, storeOptions));
    if (candidates.length === 1) return (storeSection = candidates[0].id);
    if (!probe || !candidates.length) return null;
    for (const section of candidates) {
        const mark = Math.random().toString(36).slice(2);
        const original = section.view.widgetOptions || null;
        await writeSection(section.id, { ...(original || {}), sondePublipostage: mark });
        const found = await waitForOptions(o => !!o && o.sondePublipostage === mark, 2500);
        await writeSection(section.id, original);
        if (found) {
            await waitForOptions(o => sameOptions(o, original), 2500);
            return (storeSection = section.id);
        }
    }
    throw new Error('Section du widget introuvable');
}

/* ── Table des modèles ──────────────────────────────────────────── */

// Colonnes existantes de la table des modèles (null si elle n'existe pas)
async function templateTableColumns() {
    const tables = await grist.docApi.fetchTable('_grist_Tables');
    const i = tables.tableId.indexOf(TEMPLATE_TABLE);
    if (i < 0) return null;
    const cols = await grist.docApi.fetchTable('_grist_Tables_column');
    return new Set(cols.colId.filter((_, k) => cols.parentId[k] === tables.id[i]));
}

// Actions qui créent la table (sans page) ou complètent ses colonnes
function tableActions(columns) {
    const actions = [];
    if (!columns) {
        // AddRawTable crée aussi les colonnes A, B et C
        actions.push(['AddRawTable', TEMPLATE_TABLE]);
        for (const col of ['A', 'B', 'C']) actions.push(['RemoveColumn', TEMPLATE_TABLE, col]);
    }
    for (const [colId, type] of Object.entries(TEMPLATE_COLUMNS)) {
        if (!columns || !columns.has(colId)) actions.push(['AddColumn', TEMPLATE_TABLE, colId, { type, isFormula: false }]);
    }
    return actions;
}

// Lignes de la table des modèles : [{ rowId, widget, name, hash, attachmentId }]
async function templateRows() {
    let tbl;
    try {
        tbl = await grist.docApi.fetchTable(TEMPLATE_TABLE);
    } catch (e) {
        return []; // table absente (ou illisible)
    }
    const col = name => tbl[name] || [];
    return tbl.id.map((rowId, i) => {
        const files = col('Fichier')[i];
        const ids = Array.isArray(files) ? files.filter(v => typeof v === 'number') : [];
        return { rowId, widget: col('Widget')[i], name: col('Modele')[i], hash: col('Empreinte')[i], attachmentId: ids[ids.length - 1] };
    }).filter(r => r.attachmentId && r.name);
}

// Ligne du widget : celle de sa section ; à défaut, celle qui porte
// l'empreinte de ses options (widget dupliqué, section non départagée).
function rowFor(rows, sectionId, options) {
    const own = sectionId && rows.find(r => r.widget === sectionId);
    if (own) return own;
    const hash = options && options.templateHash;
    return (hash && rows.find(r => r.hash === hash)) || null;
}

// Le document réserve-t-il la structure aux propriétaires ? Seul un
// propriétaire voit les règles ; elles ne sont lues qu'à l'ouverture.
async function structureReserved() {
    try {
        const rules = await grist.docApi.fetchTable('_grist_ACLRules');
        return (rules.permissionsText || []).some(p => /-[A-Z]*S/.test(p || ''));
    } catch (e) {
        return false;
    }
}

// Ouverture par un propriétaire d'un document qui réserve la structure : la
// table est préparée, pour que les éditeurs puissent enregistrer leurs modèles.
async function prepareTableForEditors() {
    if (isReadonly() || await templateTableColumns() || !await structureReserved()) return;
    try {
        await grist.docApi.applyUserActions(tableActions(null));
    } catch (e) {
        console.warn('Table des modèles non créée', e);
    }
}

/**
 * Enregistre le modèle du widget dans le document.
 *  - 'saved'    : fichier rangé dans la table et options à jour ;
 *  - 'row'      : fichier rangé dans la table ; options non modifiables
 *                 (structure réservée aux propriétaires) ;
 *  - 'readonly' : utilisateur en lecture seule, rien n'est écrit ;
 *  - 'owners'   : la table n'existe pas et seul un propriétaire peut la créer ;
 *  - 'section'  : deux widgets de la même table, impossibles à départager sans
 *                 écrire les options.
 * Une autre erreur est levée.
 */
async function storeTemplate({ name, hash, buffer }, { reuseRow = null } = {}) {
    if (isReadonly()) return 'readonly';
    let sectionId;
    try {
        sectionId = await ownSection();
    } catch (e) {
        if (isStructureDenied(e)) return 'section';
        throw e;
    }
    if (!sectionId) throw new Error('Section du widget introuvable');
    const columns = await templateTableColumns();
    const rows = await templateRows();
    const row = reuseRow || rows.find(r => r.widget === sectionId);
    const data = [];
    if (!reuseRow) {
        const attachmentId = await uploadTemplateFile(buffer, name);
        const fields = { Widget: sectionId, Modele: name, Empreinte: hash, Fichier: ['L', attachmentId] };
        data.push(...tableActions(columns));
        data.push(row && row.widget === sectionId
            ? ['UpdateRecord', TEMPLATE_TABLE, row.rowId, fields]
            : ['AddRecord', TEMPLATE_TABLE, null, fields]);
    } else if (reuseRow.widget !== sectionId) {
        // Modèle hérité d'un autre widget (copie) : ligne propre à celui-ci
        data.push(['AddRecord', TEMPLATE_TABLE, null, { Widget: sectionId, Modele: name, Empreinte: hash, Fichier: ['L', reuseRow.attachmentId] }]);
    }
    const pointer = await sectionAction(sectionId, pointerOptions(storeOptions, name, hash, sectionId));
    try {
        // Une seule action : un Ctrl+Z rétablit le fichier et les options
        await grist.docApi.applyUserActions([...data, pointer]);
        return 'saved';
    } catch (e) {
        if (!isStructureDenied(e)) throw e;
    }
    if (!columns || data.some(a => a[0] !== 'AddRecord' && a[0] !== 'UpdateRecord')) return 'owners';
    if (data.length) await grist.docApi.applyUserActions(data);
    return 'row';
}

/**
 * Modèle à afficher, lu dans le document :
 *  { name, hash, buffer }   modèle de la table, ou d'un ancien format ;
 *  { name, missing: true }  les options citent un modèle introuvable (copie
 *                           sans données, table supprimée…) ;
 *  null                     aucun modèle.
 * `skip(hash)` évite de retélécharger le modèle déjà chargé.
 * Un utilisateur qui peut écrire les options retire l'ancien format.
 */
async function readTemplate(options, skip = () => false) {
    const [sectionId, rows] = await Promise.all([
        ownSection({ probe: false }).catch(() => null),
        templateRows(),
    ]);
    const row = rowFor(rows, sectionId, options);
    if (row) {
        const buffer = skip(row.hash) ? null : await downloadTemplateFile(row.attachmentId);
        const template = { name: row.name, hash: row.hash, buffer, row };
        if (hasLegacy(options) || row.widget !== sectionId) migrate(template);
        return template;
    }
    const legacy = legacyFromOptions(options);
    if (legacy && legacy.data) {
        const template = { name: legacy.name, hash: legacy.hash, buffer: base64ToBuffer(legacy.data) };
        if (!template.hash) template.hash = await templateHash(template.buffer);
        migrate(template);
        return template;
    }
    if (legacy && legacy.legacyId) {
        // Téléchargée une fois par session, malgré les notifications de Grist
        if (!legacyCache || legacyCache.id !== legacy.legacyId) {
            legacyCache = { id: legacy.legacyId, template: loadLegacyFile(legacy) };
        }
        const template = await legacyCache.template;
        if (!template.missing) migrate(template);
        return template;
    }
    if (options && options.templateName) return { name: options.templateName, missing: true };
    return null;
}

let legacyCache = null;
async function loadLegacyFile(legacy) {
    try {
        const buffer = await downloadTemplateFile(legacy.legacyId);
        return { name: legacy.name, hash: await templateHash(buffer), buffer };
    } catch (e) {
        console.warn('Ancienne pièce jointe illisible', e);
        return { name: legacy.name, missing: true };
    }
}

// Reprise d'un ancien format, ou d'un modèle hérité : en arrière-plan, sans
// message. Si les options ne peuvent pas être écrites, un propriétaire la
// terminera en ouvrant le widget.
let migrating = null;
function migrate(template) {
    if (isReadonly() || migrating === template.hash) return;
    migrating = template.hash;
    storeTemplate(template, { reuseRow: template.row || null })
        .catch(e => console.warn('Reprise du modèle impossible', e));
}
