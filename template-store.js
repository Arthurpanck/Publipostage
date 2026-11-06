/**
 * MODÈLE PERSISTANT
 * Objet : garder le modèle attaché au widget, pour tous les utilisateurs et
 * dans le temps.
 *
 * Le modèle est rangé dans les options du widget (en base64), pas en pièce
 * jointe. Une pièce jointe qu'aucune cellule ne référence est marquée
 * supprimée par Grist dans l'heure, puis effacée 7 jours plus tard ; elle est
 * refusée aux utilisateurs soumis à des règles d'accès et retirée des copies
 * « comme modèle ».
 *
 * Le widget écrit lui-même ses options dans _grist_Views_section :
 * grist.setOptions ne fait qu'allumer le bouton « Enregistrer » de Grist, et
 * rien n'est conservé tant que personne ne clique dessus.
 */

const TEMPLATE_MAX_BYTES = 8 * 1024 * 1024;

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

/* ── Encodage ───────────────────────────────────────────────────── */

function bufferToBase64(buffer) {
    const bytes = new Uint8Array(buffer);
    let out = '';
    for (let i = 0; i < bytes.length; i += 0x8000) {
        out += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    }
    return btoa(out);
}

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

// Options du widget pour un nouveau modèle (l'ancien templateId disparaît)
async function templateOptions(previous, buffer, name) {
    const { templateId, sondePublipostage, ...rest } = previous || {};
    return { ...rest, templateName: name, templateHash: await templateHash(buffer), templateData: bufferToBase64(buffer) };
}

// Modèle décrit par les options : intégré, ancienne pièce jointe, ou rien
function templateFromOptions(options) {
    if (!options || !options.templateName) return null;
    if (options.templateData) {
        return { name: options.templateName, hash: options.templateHash, buffer: base64ToBuffer(options.templateData) };
    }
    if (options.templateId) return { name: options.templateName, legacyId: options.templateId };
    return null;
}

/* ── Section du widget ──────────────────────────────────────────── */

const isReadonly = () => new URL(location.href).searchParams.get('readonly') === 'true';

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

async function writeSection(id, widgetOptions) {
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    const section = parseSection(sections, sections.id.indexOf(id));
    section.view.widgetOptions = widgetOptions;
    section.options.customView = JSON.stringify(section.view);
    await grist.docApi.applyUserActions([['UpdateRecord', '_grist_Views_section', id, { options: JSON.stringify(section.options) }]]);
}

// Retrouve la section du widget : ses options enregistrées sont celles que
// Grist nous a transmises. Si plusieurs sections ont les mêmes (deux widgets
// neufs sur la même table), on marque l'une puis l'autre et on regarde
// laquelle Grist nous renvoie.
async function ownSection() {
    if (storeSection) return storeSection;
    const candidates = (await candidateSections()).filter(s => sameOptions(s.view.widgetOptions, storeOptions));
    if (candidates.length === 1) return (storeSection = candidates[0].id);
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

async function savedOptionsOf(id) {
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    const i = sections.id.indexOf(id);
    return i < 0 ? undefined : parseSection(sections, i).view.widgetOptions;
}

/**
 * Enregistre les options dans le document.
 *  - 'saved'    : écrites et relues dans le document ;
 *  - 'readonly' : utilisateur en lecture seule, rien n'est écrit ;
 *  - 'owners'   : le document réserve la modification de sa structure aux
 *                 propriétaires (le bouton « Enregistrer » de Grist échouerait
 *                 aussi) ; rien n'est écrit ;
 *  - 'pending'  : écriture directe impossible ; les options sont proposées à
 *                 Grist, qui affiche son bouton « Enregistrer » ;
 *  - 'failed'   : écriture directe impossible et { fallback: false }.
 */
async function persistOptions(options, { fallback = true } = {}) {
    if (isReadonly()) return 'readonly';
    try {
        const id = await ownSection();
        await writeSection(id, options);
        if (sameOptions(await savedOptionsOf(id), options)) return 'saved';
    } catch (e) {
        console.warn('Enregistrement direct impossible', e);
        if (/structure access rules/i.test(String(e && e.message))) return 'owners';
    }
    if (!fallback) return 'failed';
    await grist.setOptions(options);
    return 'pending';
}

// Un modèle en attente du bouton « Enregistrer » est-il désormais dans le document ?
async function isTemplateSaved(hash) {
    const sections = await grist.docApi.fetchTable('_grist_Views_section');
    return sections.id.some((_, i) => {
        const options = parseSection(sections, i).view.widgetOptions;
        return !!options && options.templateHash === hash;
    });
}
