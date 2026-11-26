/**
 * MODÈLE PERSISTANT
 * Objet : garder le modèle attaché au widget, pour tous les utilisateurs et
 * dans le temps, sans alourdir l'ouverture du document.
 *
 * Le modèle est une pièce jointe rangée dans la table
 * Publipostage_Modeles_NE_PAS_SUPPRIMER (données brutes, sans page) : une
 * ligne par widget, repérée par le numéro de sa section. Une pièce jointe
 * rangée dans une cellule n'est pas purgée par Grist, et seul le widget
 * affiché la télécharge.
 *
 * Les options du widget ne gardent que le nom et l'empreinte du modèle : leur
 * changement prévient les autres utilisateurs (grist.onOptions) et Ctrl+Z
 * rétablit le modèle précédent. Le fichier lui-même n'y est plus : Grist
 * envoie les options de tous les widgets à chaque ouverture du document.
 *
 * Chaque widget porte aussi son numéro dans ses options (templateWidget) :
 * c'est ce qui lui permet de se reconnaître parmi les widgets de la même
 * table. Tout utilisateur qui peut écrire les options remet d'aplomb, à
 * l'ouverture, celles de tous les widgets Publipostage du document (voir
 * healSections) : numéro manquant, options remises à zéro par Grist (widget
 * re-sélectionné, adresse changée), widget ou page dupliqués.
 *
 * Écrire les options demande le droit de modifier la structure ; ranger le
 * fichier dans la table demande seulement celui de modifier les données. Un
 * éditeur d'un document qui réserve la structure aux propriétaires enregistre
 * donc son modèle dans la table, dès qu'un propriétaire a ouvert un widget
 * (table créée par prepareTemplateTable, widgets numérotés par healSections).
 *
 * Anciens formats repris (puis retirés des options par un utilisateur qui
 * peut les écrire) : modèle en base64 dans les options (templateData) et
 * pièce jointe non rangée (templateId).
 */

// Le nom de la table prévient qu'elle sert au widget ; ses colonnes le
// rappellent dans leur description.
const TEMPLATE_TABLE = 'Publipostage_Modeles_NE_PAS_SUPPRIMER';
const TEMPLATE_WARNING = 'Table gérée par le widget Publipostage : ne pas la supprimer ni la modifier, sinon les modèles des widgets sont perdus.';
const TEMPLATE_COLUMNS = {
    Widget: { type: 'Int', description: `Numéro du widget Publipostage qui utilise ce modèle. ${TEMPLATE_WARNING}` },
    Modele: { type: 'Text', description: `Nom du fichier modèle. ${TEMPLATE_WARNING}` },
    Empreinte: { type: 'Text', description: `Empreinte du fichier, pour savoir si le modèle a changé. ${TEMPLATE_WARNING}` },
    Fichier: { type: 'Attachments', description: `Modèle Word ou PDF du widget. ${TEMPLATE_WARNING}` },
};

// Dernières options reçues de Grist (grist.onOptions) et section du widget
let storeOptions = null;
let storeSection = null;

function noteOptions(options) {
    storeOptions = options || null;
}

const pause = ms => new Promise(resolve => setTimeout(resolve, ms));

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
// départage deux widgets de la même table, même après une duplication).
// sondePublipostage : marque laissée par une ancienne version.
function pointerOptions(previous, name, hash, widget) {
    const { templateId, templateData, sondePublipostage, ...rest } = previous || {};
    return { ...rest, templateName: name, templateHash: hash, templateWidget: widget };
}

// Options d'un widget sans modèle enregistré : son numéro seulement
function numberedOptions(previous, widget) {
    const { sondePublipostage, ...rest } = previous || {};
    return { ...rest, templateWidget: widget };
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
    if (response.status === 413) {
        throw new Error(`fichier trop lourd pour ce serveur Grist (${Math.ceil(buffer.byteLength / 1048576)} Mo). Réduisez les images du document`);
    }
    if (!response.ok) throw new Error(`envoi du fichier refusé par Grist (${response.status})`);
    return (await response.json())[0];
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

// Adresse d'une page, sans protocole, paramètres, index.html ni / final
function pageKey(url) {
    if (!url) return null;
    try {
        const page = new URL(url, location.href);
        return page.host + page.pathname.replace(/index\.html$/, '').replace(/\/+$/, '');
    } catch (e) {
        return null;
    }
}

const OWN_OPTIONS = ['templateName', 'templateHash', 'templateWidget', 'templateId', 'templateData', 'sondePublipostage'];

// Widgets Publipostage du document : sections personnalisées qui affichent
// cette page (sameTable : sur la table de ce widget). Si aucune adresse
// enregistrée n'est la nôtre (redirection, catalogue de widgets…), ce sont
// les sections de la table du widget sans options ou avec les nôtres.
async function widgetSections() {
    const [tableId, tables, sections] = await Promise.all([
        grist.getSelectedTableId(),
        grist.docApi.fetchTable('_grist_Tables'),
        grist.docApi.fetchTable('_grist_Views_section'),
    ]);
    const tableRef = tableRefOf(tables, tableId);
    const custom = sections.id.map((_, i) => i)
        .filter(i => sections.parentKey[i] === 'custom')
        .map(i => ({ ...parseSection(sections, i), tableRef: sections.tableRef[i], sameTable: sections.tableRef[i] === tableRef }));
    const here = pageKey(location.href);
    const ours = custom.filter(s => pageKey(s.view.url || (s.view.widgetDef && s.view.widgetDef.url)) === here);
    if (ours.some(s => s.sameTable)) return ours;
    return custom.filter(s => s.sameTable && Object.keys(s.view.widgetOptions || {}).every(k => OWN_OPTIONS.includes(k)));
}

// Section telle qu'enregistrée dans le document
async function readSection(id) {
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    return parseSection(sections, sections.id.indexOf(id));
}

// Action qui remplace les options enregistrées d'une section
function sectionAction(section, widgetOptions) {
    const customView = JSON.stringify({ ...section.view, widgetOptions });
    return ['UpdateRecord', '_grist_Views_section', section.id, { options: JSON.stringify({ ...section.options, customView }) }];
}

// La section du widget est celle de sa table dont les options enregistrées
// sont celles que Grist nous a transmises, si elle est la seule.
function sectionWithOurOptions(sections) {
    const same = sections.filter(s => s.sameTable && sameOptions(s.view.widgetOptions, storeOptions));
    return same.length === 1 ? same[0].id : null;
}

// Retrouve la section du widget, avant d'y écrire. Les options des widgets
// du document sont d'abord remises d'aplomb (healSections) : chacun porte
// alors son numéro, et Grist nous transmet le nôtre. Grist transmet les
// options un instant après les avoir enregistrées (les nôtres ont pu être
// modifiées par un autre widget) : la réponse n'est retenue que si elles
// n'ont pas changé 300 ms plus tard.
// { heal: false } pour se contenter de lire : la réponse n'est pas gardée,
// et des widgets aux options identiques ne sont pas départagés.
async function ownSection({ heal = !isReadonly() } = {}) {
    if (storeSection) return storeSection;
    if (!heal) return sectionWithOurOptions(await widgetSections());
    const healed = await healSections() !== 'denied';
    for (let attempt = 0; attempt < 10; attempt++) {
        const options = storeOptions;
        const own = sectionWithOurOptions(await widgetSections());
        if (!own && !healed) break;
        await pause(300);
        if (own && storeOptions === options) return (storeSection = own);
    }
    return null;
}

// Remet d'aplomb les options des widgets Publipostage du document, d'après
// la table des modèles :
//  - un widget qui a une ligne dans la table la désigne : Grist remet les
//    options à zéro quand on re-sélectionne le widget ou change son adresse,
//    et un éditeur peut avoir rangé un modèle sans pouvoir écrire les options ;
//  - un widget dupliqué porte le numéro de son original : il reçoit le sien
//    (et sa propre ligne quand il est ouvert, voir readTemplate) ;
//  - des widgets d'une même table aux options identiques (widgets neufs)
//    reçoivent leur numéro, celui de leur section.
// Renvoie 'changed' si des options ont été écrites, 'same' s'il n'y avait
// rien à faire, 'denied' si elles n'ont pas pu l'être (structure réservée aux
// propriétaires).
let healDenied = false;
async function healSections() {
    const [sections, rows] = await Promise.all([widgetSections(), templateRows()]);
    const twinKey = s => `${s.tableRef}:${stableJson(s.view.widgetOptions || null)}`;
    const twins = new Map();
    for (const section of sections) twins.set(twinKey(section), (twins.get(twinKey(section)) || 0) + 1);
    const actions = [];
    for (const section of sections) {
        const options = section.view.widgetOptions || {};
        const row = rowFor(rows, section.id, options);
        const numbered = options.templateWidget !== undefined || twins.get(twinKey(section)) > 1;
        const wanted = row ? pointerOptions(options, row.name, row.hash, section.id)
            : numbered ? numberedOptions(options, section.id) : options;
        if (!sameOptions(wanted, options)) actions.push(sectionAction(section, wanted));
    }
    if (!actions.length) return 'same';
    try {
        await grist.docApi.applyUserActions(actions);
        return 'changed';
    } catch (e) {
        if (!isStructureDenied(e)) throw e;
        healDenied = true;
        return 'denied';
    }
}

/* ── Table des modèles ──────────────────────────────────────────── */

// Colonnes existantes de la table des modèles (null si elle n'existe pas)
async function templateTableColumns() {
    const tableRef = tableRefOf(await grist.docApi.fetchTable('_grist_Tables'), TEMPLATE_TABLE);
    if (tableRef === undefined) return null;
    const cols = await grist.docApi.fetchTable('_grist_Tables_column');
    return new Set(cols.colId.filter((_, i) => cols.parentId[i] === tableRef));
}

// Actions qui créent la table (sans page) ou complètent ses colonnes
function tableActions(columns) {
    const actions = [];
    if (!columns) {
        // AddRawTable crée aussi les colonnes A, B et C
        actions.push(['AddRawTable', TEMPLATE_TABLE]);
        for (const col of ['A', 'B', 'C']) actions.push(['RemoveColumn', TEMPLATE_TABLE, col]);
    }
    for (const [colId, { type, description }] of Object.entries(TEMPLATE_COLUMNS)) {
        if (columns && columns.has(colId)) continue;
        actions.push(['AddColumn', TEMPLATE_TABLE, colId, { type, isFormula: false }]);
        // AddColumn ignore la description
        actions.push(['ModifyColumn', TEMPLATE_TABLE, colId, { description }]);
    }
    return actions;
}

// Lignes de la table des modèles : [{ rowId, widget, name, hash, attachmentId }]
// Aucune si la table n'existe pas (ou est masquée par une règle d'accès) ;
// si elle existe, un échec de lecture est une erreur, pas un modèle perdu.
async function templateRows() {
    let tbl;
    try {
        tbl = await grist.docApi.fetchTable(TEMPLATE_TABLE);
    } catch (e) {
        if (await templateTableColumns()) throw e;
        return [];
    }
    const col = name => tbl[name] || [];
    return tbl.id.map((rowId, i) => {
        const files = col('Fichier')[i];
        const ids = Array.isArray(files) ? files.filter(v => typeof v === 'number') : [];
        return { rowId, widget: col('Widget')[i], name: col('Modele')[i], hash: col('Empreinte')[i], attachmentId: ids[ids.length - 1] };
    }).filter(r => r.attachmentId && r.name);
}

const rowFields = (widget, { name, hash, attachmentId }) =>
    ({ Widget: widget, Modele: name, Empreinte: hash, Fichier: ['L', attachmentId] });

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

// Crée la table ou complète ses colonnes. Renvoie false si seul un
// propriétaire le peut. Plusieurs widgets peuvent s'y essayer au même moment :
// un seul réussit, les autres échouent, relisent la table et la trouvent prête.
async function ensureTemplateTable() {
    for (let attempt = 0; ; attempt++) {
        const setup = tableActions(await templateTableColumns());
        if (!setup.length) return true;
        try {
            await grist.docApi.applyUserActions(setup);
            return true;
        } catch (e) {
            if (isStructureDenied(e)) return false;
            if (attempt) throw e;
        }
    }
}

// À l'ouverture du widget, par un utilisateur qui en a le droit :
//  - une table incomplète (colonne supprimée) est réparée ;
//  - dans un document qui réserve la structure aux propriétaires, la table
//    est créée, pour que les éditeurs puissent y enregistrer leurs modèles.
async function prepareTemplateTable() {
    if (isReadonly()) return;
    try {
        if (await structureReserved() || await templateTableColumns()) await ensureTemplateTable();
    } catch (e) {
        console.warn('Table des modèles non préparée', e);
    }
}

// Avant la première lecture du modèle, par un utilisateur qui peut écrire :
// les options des widgets du document sont remises d'aplomb (healSections).
// Si elles ont changé, Grist nous transmet les nôtres un instant plus tard.
function healBeforeReading() {
    if (isReadonly()) return Promise.resolve();
    return enqueueWrite(async () => { if (await healSections() === 'changed') await pause(300); })
        .catch(e => console.warn('Options des widgets non vérifiées', e));
}

/* ── Enregistrement et lecture ──────────────────────────────────── */

// Les écritures du modèle se suivent : une reprise en arrière-plan ne croise
// pas un envoi, et un envoi annule les reprises qui attendent encore.
let writeQueue = Promise.resolve();
let uploads = 0;

function enqueueWrite(task) {
    const run = writeQueue.then(task);
    writeQueue = run.catch(() => {});
    return run;
}

/**
 * Enregistre dans le document le modèle choisi par l'utilisateur.
 *  - 'saved'    : fichier rangé dans la table et options à jour ;
 *  - 'row'      : fichier rangé dans la table ; options non modifiables
 *                 (structure réservée aux propriétaires) ;
 *  - 'readonly' : utilisateur en lecture seule, rien n'est écrit ;
 *  - 'owners'   : table absente ou incomplète, et seul un propriétaire peut
 *                 la préparer ; rien n'est envoyé ;
 *  - 'section'  : plusieurs widgets de la même table, impossibles à
 *                 départager sans écrire les options.
 * Une autre erreur est levée.
 */
function storeTemplate(template) {
    uploads++;
    return enqueueWrite(() => writeTemplate(template));
}

// Reprise d'un ancien format ou d'un modèle hérité (widget dupliqué) : en
// arrière-plan, sans message, une fois par modèle. Si les options ne peuvent
// pas être écrites, un propriétaire la terminera en ouvrant le widget.
let migrating = null;
function migrate(template) {
    if (isReadonly() || migrating === template.hash) return;
    migrating = template.hash;
    const before = uploads;
    enqueueWrite(() => before === uploads && writeTemplate(template, { migration: true }))
        .catch(e => console.warn('Reprise du modèle impossible', e));
}

// template : { name, hash } et son fichier, déjà dans le document
// (attachmentId) ou à téléverser (buffer).
async function writeTemplate({ name, hash, buffer, attachmentId }, { migration = false } = {}) {
    if (isReadonly()) return 'readonly';
    const sectionId = await ownSection();
    if (!sectionId) {
        if (healDenied) return 'section';
        throw new Error('Section du widget introuvable');
    }

    const own = (await templateRows()).find(r => r.widget === sectionId);
    // Reprise : la ligne déjà enregistrée pour ce widget fait foi
    if (migration && own) ({ name, hash } = own);
    const data = [];
    if (!own || own.hash !== hash || own.name !== name) {
        // La table d'abord : si elle ne peut pas être préparée, rien n'est envoyé
        if (!await ensureTemplateTable()) return 'owners';
        if (!attachmentId) attachmentId = await uploadTemplateFile(buffer, name);
        const fields = rowFields(sectionId, { name, hash, attachmentId });
        data.push(own ? ['UpdateRecord', TEMPLATE_TABLE, own.rowId, fields] : ['AddRecord', TEMPLATE_TABLE, null, fields]);
    }
    const section = await readSection(sectionId);
    const pointer = sectionAction(section, pointerOptions(section.view.widgetOptions, name, hash, sectionId));
    try {
        // Une seule action : un Ctrl+Z rétablit le fichier et les options
        await grist.docApi.applyUserActions([...data, pointer]);
        return 'saved';
    } catch (e) {
        if (!isStructureDenied(e)) throw e;
    }
    if (data.length) await grist.docApi.applyUserActions(data);
    return 'row';
}

/**
 * Modèle à afficher, lu dans le document :
 *  { name, hash, buffer }   modèle de la table, ou d'un ancien format
 *                           (buffer null si `skip(hash)` : déjà chargé) ;
 *  { name, missing: true }  les options citent un modèle introuvable (copie
 *                           sans données, table supprimée ou illisible…) ;
 *  null                     aucun modèle.
 * Un ancien format ou un modèle hérité est repris au passage (migrate).
 */
async function readTemplate(options, skip = () => false) {
    const [sectionId, rows] = await Promise.all([
        ownSection({ heal: false }).catch(() => null),
        templateRows(),
    ]);
    const row = rowFor(rows, sectionId, options);
    if (row) {
        const buffer = skip(row.hash) ? null : await downloadAttachment(row.attachmentId);
        if (hasLegacy(options) || (sectionId && row.widget !== sectionId)) migrate(row);
        return { name: row.name, hash: row.hash, buffer };
    }
    const legacy = legacyFromOptions(options);
    if (legacy && legacy.data) {
        const buffer = base64ToBuffer(legacy.data);
        const template = { name: legacy.name, hash: legacy.hash || await templateHash(buffer), buffer };
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

// Ancienne pièce jointe, non rangée dans une cellule : elle est reprise telle
// quelle dans la table (attachmentId), sans nouvel envoi.
let legacyCache = null;
async function loadLegacyFile(legacy) {
    try {
        const buffer = await downloadAttachment(legacy.legacyId);
        return { name: legacy.name, hash: await templateHash(buffer), buffer, attachmentId: legacy.legacyId };
    } catch (e) {
        console.warn('Ancienne pièce jointe illisible', e);
        return { name: legacy.name, missing: true };
    }
}
