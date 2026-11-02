/**
 * MODULE RENDUS
 * Version : 1.0
 * Mise en forme des valeurs selon le type de la colonne Grist (DOCX) :
 *   - Booléen : case à cocher ☑ / ☐, automatique.
 *   - Choix, Choix multiples, Référence, Référence multiple : tag teinté en
 *     gras, automatique. La couleur vient de l'option Grist (Choix) ou de la
 *     colonne affichée de la table liée quand c'est un Choix ; sinon gris.
 *   - {Etape|etapes} : progression dans l'ordre des options du Choix.
 *   - {Actuel|barre:Objectif} ou {Actuel|barre:Objectif:Pourcentage} : barre
 *     de fond. Toutes les valeurs viennent de colonnes de la ligne.
 * Texte, nombres sans option et dates restent insérés tels quels.
 *
 * Principe : le parser docxtemplater renvoie un marqueur à la place de la
 * valeur ; après le rendu, chaque marqueur est remplacé par des runs Word qui
 * héritent de la mise en forme de la balise (police, taille…), ou par un
 * tableau pour des étapes seules dans leur paragraphe.
 */

const RENDER = Symbol('render');
const MARK_OPEN = '', MARK_CLOSE = '';
const MARK_RE = /(\d+)/g;
const FORMATS = ['barre', 'etapes'];

// ---------------------------------------------------------------------------
// Métadonnées Grist
// ---------------------------------------------------------------------------

function parseWidgetOptions(text) {
    if (!text) return {};
    if (typeof text === 'object') return text;
    try { return JSON.parse(text) || {}; } catch (e) { return {}; }
}

const RENDER_META = Symbol('renderMeta');
const RENDER_INFO = Symbol('renderInfo');

// Spécifications de rendu des tables utiles au lot : table du widget et
// tables liées (batch.sources : nom normalisé -> identifiant réel).
// Un échec de lecture n'empêche pas la génération : les valeurs restent brutes.
async function loadRenderInfo(batch) {
    try {
        // Même modèle pour tout le lot : les spécifications sont calculées une fois
        if (batch.cache.has(RENDER_INFO)) return await batch.cache.get(RENDER_INFO);
        const meta = await cached(batch.cache, RENDER_META, async () => {
            const [cols, tables, docInfo] = await Promise.all([
                fetchTableCached(batch.cache, '_grist_Tables_column'),
                fetchTableCached(batch.cache, '_grist_Tables'),
                fetchTableCached(batch.cache, '_grist_DocInfo').catch(() => null),
            ]);
            const settings = parseWidgetOptions(docInfo && docInfo.documentSettings && docInfo.documentSettings[0]);
            return { cols, tables, locale: settings.locale || 'fr-FR', currency: settings.currency || 'EUR' };
        });
        const widgetTable = await cached(batch.cache, WIDGET_TABLE, () => grist.getSelectedTableId());
        const info = { locale: meta.locale, currency: meta.currency, root: null, byKey: new Map() };
        info.root = await tableSpec(batch, meta, widgetTable);
        for (const [key, tableId] of (batch.sources || new Map())) {
            info.byKey.set(key, await tableSpec(batch, meta, tableId));
        }
        batch.cache.set(RENDER_INFO, info);
        return info;
    } catch (e) {
        // Sans métadonnées, les valeurs sont insérées brutes (comme avant).
        console.warn('Mise en forme par type indisponible', e);
        return null;
    }
}

// { colonne normalisée: spec } pour une table
async function tableSpec(batch, meta, tableId) {
    const { cols, tables } = meta;
    const tableRef = tables.id[tables.tableId.indexOf(tableId)];
    const byRowId = new Map(cols.id.map((id, i) => [id, i]));
    const spec = {};
    for (let i = 0; i < cols.id.length; i++) {
        if (cols.parentId[i] !== tableRef) continue;
        const colId = cols.colId[i];
        if (/^(?:manualSort|gristHelper_)/.test(colId)) continue;
        const [, kind, target] = /^(Ref|RefList):(.+)$/.exec(cols.type[i]) || [];
        const opts = parseWidgetOptions(cols.widgetOptions[i]);
        const s = { type: kind || cols.type[i], label: cols.label ? cols.label[i] : colId, opts };
        if (kind) {
            s.refs = await refDisplay(batch, meta, target, byRowId.get(cols.visibleCol ? cols.visibleCol[i] : 0));
        }
        spec[sanitizeKey(colId)] = s;
    }
    return spec;
}

// id -> { text, fill } pour les fiches de la table liée
async function refDisplay(batch, meta, tableId, visibleIndex) {
    const { cols } = meta;
    const tbl = await fetchTableCached(batch.cache, tableId);
    const visibleCol = visibleIndex === undefined ? null : cols.colId[visibleIndex];
    const visibleOpts = visibleIndex === undefined ? {} : parseWidgetOptions(cols.widgetOptions[visibleIndex]);
    const isChoice = visibleIndex !== undefined && cols.type[visibleIndex] === 'Choice';
    const map = new Map();
    tbl.id.forEach((id, i) => {
        const raw = visibleCol ? tbl[visibleCol][i] : id;
        const text = Array.isArray(raw) ? raw.slice(raw[0] === 'L' ? 1 : 0).join(', ') : raw;
        map.set(id, { text: text == null ? '' : String(text), style: isChoice ? choiceStyle(visibleOpts, raw) : null });
    });
    return map;
}

function choiceStyle(opts, value) {
    const o = opts && opts.choiceOptions && opts.choiceOptions[value];
    return o ? { fill: o.fillColor, text: o.textColor } : null;
}

// Rattache les spécifications aux données normalisées (racine et lignes liées)
function attachRenderInfo(data, info) {
    if (!info || !data || typeof data !== 'object') return data;
    if (info.root) data[RENDER] = info.root;
    for (const [key, spec] of info.byKey) {
        const rows = data[key];
        if (Array.isArray(rows)) rows.forEach(row => { if (row && typeof row === 'object') row[RENDER] = spec; });
    }
    return data;
}

// ---------------------------------------------------------------------------
// Parser docxtemplater : {Colonne}, {Colonne|format:Col1:Col2}
// ---------------------------------------------------------------------------

function splitTag(tag) {
    const i = tag.indexOf('|');
    if (i < 0) return { base: tag, format: null, args: [] };
    const [format, ...args] = tag.slice(i + 1).split(':');
    return { base: tag.slice(0, i), format, args };
}

// ctx : { registry: [], unknownTags: Set, info } ; un parser par génération
function makeParser(ctx) {
    return function parser(tag, meta) {
        const part = meta && meta.tag;
        const { base, format, args } = splitTag(tag);
        if (format !== null) {
            if (!FORMATS.includes(format)) {
                throw new Error(`Mise en forme inconnue « ${format} » dans {${tag}}. Possibles : ${FORMATS.join(', ')}.`);
            }
            if (part && part.module) {
                throw new Error(`Une mise en forme ne s'utilise pas dans une boucle ou une condition : {${tag}}.`);
            }
            if (format === 'barre' && (args.length < 1 || args.length > 2 || args.some(a => !a))) {
                throw new Error(`Barre : indiquer la colonne objectif, et éventuellement la colonne pourcentage : {Actuel|barre:Objectif} ou {Actuel|barre:Objectif:Pourcentage}.`);
            }
            if (format === 'etapes' && args.length) {
                throw new Error(`Étapes : aucune option, l'ordre est celui des choix de la colonne : {${base}|etapes}.`);
            }
        }
        return {
            get(scope, context) {
                if (tag === '.') return scope;
                if (!scope) return scope;
                const value = scope[base];
                if (value === undefined) return undefined;
                if (part && part.module) return value; // boucles, conditions : valeur brute
                const spec = scope[RENDER] && scope[RENDER][base];
                const lookup = (name) => lookupScopes(context, name);
                const render = renderer(ctx, { value, spec, format, args, tag, lookup });
                if (!render) return value;
                return MARK_OPEN + (ctx.registry.push(render) - 1) + MARK_CLOSE;
            },
        };
    };
}

function lookupScopes(context, name) {
    const list = (context && context.scopeList) || [];
    for (let n = (context && context.num !== undefined ? context.num : list.length - 1); n >= 0; n--) {
        const scope = list[n];
        if (scope && typeof scope === 'object' && scope[name] !== undefined) {
            return { value: scope[name], spec: scope[RENDER] && scope[RENDER][name] };
        }
    }
    return undefined;
}

// Renvoie { inline(rPr) -> runs XML, block?() -> XML } ou null (valeur brute)
function renderer(ctx, { value, spec, format, args, tag, lookup }) {
    const info = ctx.info || {};
    if (format === 'barre') {
        const goal = lookup(args[0]);
        const pct = args[1] ? lookup(args[1]) : null;
        if (!goal) ctx.unknownTags.add(args[0]);
        if (args[1] && !pct) ctx.unknownTags.add(args[1]);
        if (!goal || (args[1] && !pct)) return null;
        return { inline: rPr => barRuns(rPr, { value, spec }, goal, pct, info) };
    }
    if (format === 'etapes') {
        const choices = spec && spec.opts && spec.opts.choices;
        if (!Array.isArray(choices) || !choices.length) {
            ctx.unknownTags.add(`${tag} (colonne sans liste de choix)`);
            return null;
        }
        const color = (choiceStyle(spec.opts, value) || {}).fill;
        return {
            inline: rPr => stepsRuns(rPr, choices, choices.indexOf(value), color),
            block: () => stepsTable(choices, choices.indexOf(value), color),
        };
    }
    if (!spec) return null;
    switch (spec.type) {
        case 'Bool':
            return { inline: rPr => checkboxRun(rPr, value === true || value === 1 || value === 'true') };
        case 'Choice':
            if (value === null || value === '') return { inline: () => '' };
            return { inline: rPr => tagRuns(rPr, [{ text: String(value), style: choiceStyle(spec.opts, value) }]) };
        case 'ChoiceList':
            return { inline: rPr => tagRuns(rPr, listValue(value).map(v => ({ text: String(v), style: choiceStyle(spec.opts, v) }))) };
        case 'Ref':
        case 'RefList':
            return { inline: rPr => tagRuns(rPr, refIds(value).map(id => (spec.refs && spec.refs.get(id)) || { text: String(id), style: null })
                .filter(r => r.text !== '')) };
        default:
            return null;
    }
}

function listValue(value) {
    if (value == null || value === '') return [];
    if (Array.isArray(value)) return value[0] === 'L' ? value.slice(1) : value;
    return [value];
}

// Ref : id, 0 ou fiche développée {id, …} ; RefList : [ids] ou ["L", ids]
function refIds(value) {
    return listValue(value).map(v => (v && typeof v === 'object' ? v.id : v)).filter(v => Number.isInteger(v) && v > 0);
}

// ---------------------------------------------------------------------------
// Couleurs
// ---------------------------------------------------------------------------

function hexToRgb(hex) {
    const m = /^#?([0-9a-f]{6})$/i.exec(String(hex || '').trim());
    if (!m) return null;
    const n = parseInt(m[1], 16);
    return [n >> 16, (n >> 8) & 255, n & 255];
}
const toHex = rgb => rgb.map(x => Math.round(x).toString(16).padStart(2, '0')).join('').toUpperCase();
const mix = (rgb, to, k) => rgb.map((x, i) => x + (to[i] - x) * k);
function luminance(rgb) {
    const [r, g, b] = rgb.map(x => { x /= 255; return x <= 0.03928 ? x / 12.92 : Math.pow((x + 0.055) / 1.055, 2.4); });
    return 0.2126 * r + 0.7152 * g + 0.0722 * b;
}
// Fond très clair et texte foncé dans la teinte de la couleur Grist
function tintColors(fill) {
    const rgb = hexToRgb(fill);
    if (!rgb) return { bg: 'EEF1F4', fg: '2B2F36' };
    let k = 0.84, bg = mix(rgb, [255, 255, 255], k);
    while (luminance(bg) < 0.78 && k < 0.95) { k += 0.02; bg = mix(rgb, [255, 255, 255], k); }
    let s = 0.4, fg = mix(rgb, [0, 0, 0], s);
    while (luminance(fg) > 0.09 && s < 0.85) { s += 0.05; fg = mix(rgb, [0, 0, 0], s); }
    return { bg: toHex(bg), fg: toHex(fg) };
}
// Couleur d'accent lisible sur fond blanc
function accentColor(fill) {
    const rgb = hexToRgb(fill);
    if (!rgb) return '0D5E66';
    let s = 0, c = rgb;
    while (luminance(c) > 0.18 && s < 0.8) { s += 0.05; c = mix(rgb, [0, 0, 0], s); }
    return toHex(c);
}

// ---------------------------------------------------------------------------
// XML Word
// ---------------------------------------------------------------------------

const xmlEscape = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');

const RPR_ORDER = ['w:rStyle', 'w:rFonts', 'w:b', 'w:bCs', 'w:i', 'w:iCs', 'w:caps', 'w:smallCaps', 'w:strike',
    'w:dstrike', 'w:outline', 'w:shadow', 'w:emboss', 'w:imprint', 'w:noProof', 'w:snapToGrid', 'w:vanish',
    'w:webHidden', 'w:color', 'w:spacing', 'w:w', 'w:kern', 'w:position', 'w:sz', 'w:szCs', 'w:highlight', 'w:u',
    'w:effect', 'w:bdr', 'w:shd', 'w:fitText', 'w:vertAlign', 'w:rtl', 'w:cs', 'w:em', 'w:lang',
    'w:eastAsianLayout', 'w:specVanish', 'w:oMath'];

// Propriétés héritées de la balise + surcharges (null = supprimer), dans
// l'ordre exigé par le schéma Word.
function mergeRPr(inherited, overrides) {
    const props = new Map(), others = [];
    const re = /<([\w:]+)(?:\s[^>]*?)?\/>|<([\w:]+)(?:\s[^>]*)?>[\s\S]*?<\/\2>/g;
    let m;
    while ((m = re.exec(inherited || '')) !== null) {
        const name = m[1] || m[2];
        if (RPR_ORDER.includes(name)) props.set(name, m[0]); else others.push(m[0]);
    }
    for (const [name, xml] of Object.entries(overrides || {})) {
        if (xml === null) props.delete(name); else props.set(name, xml);
    }
    const body = RPR_ORDER.filter(n => props.has(n)).map(n => props.get(n)).join('') + others.join('');
    return body ? `<w:rPr>${body}</w:rPr>` : '';
}

const runXml = (rPr, overrides, text) =>
    `<w:r>${mergeRPr(rPr, overrides)}<w:t xml:space="preserve">${xmlEscape(text)}</w:t></w:r>`;

const fonts = name => `<w:rFonts w:ascii="${name}" w:hAnsi="${name}" w:cs="${name}" w:eastAsia="${name}"/>`;
const SYMBOL_FONT = 'Segoe UI Symbol';
const MONO_FONT = 'Consolas';
const NBSP = ' ';

function checkboxRun(rPr, checked) {
    // U+FE0E : présentation texte, pas d'emoji coloré
    return runXml(rPr, { 'w:rFonts': fonts(SYMBOL_FONT) }, (checked ? '☑' : '☐') + '\uFE0E');
}

// Tags teintés en gras, séparés par une espace normale (retour à la ligne possible)
function tagRuns(rPr, items) {
    return items.map(({ text, style }) => {
        const { bg, fg } = tintColors(style && style.fill);
        return runXml(rPr, {
            'w:b': '<w:b/>', 'w:bCs': '<w:bCs/>', 'w:color': `<w:color w:val="${fg}"/>`,
            'w:highlight': null, 'w:u': null, 'w:strike': null, 'w:bdr': null,
            'w:shd': `<w:shd w:val="clear" w:color="auto" w:fill="${bg}"/>`,
        }, NBSP + text + NBSP);
    }).join(runXml(rPr, {}, ' '));
}

// Formatage d'un nombre comme dans Grist (mode, devise, décimales)
function formatNumber(value, spec, info) {
    if (typeof value !== 'number' || !isFinite(value)) return value == null ? '' : String(value);
    const o = (spec && spec.opts) || {};
    const intl = {};
    if (o.numMode === 'currency') Object.assign(intl, { style: 'currency', currency: o.currency || info.currency || 'EUR' });
    else if (o.numMode === 'percent') intl.style = 'percent';
    else if (o.numMode === 'scientific') intl.notation = 'scientific';
    else if (o.numMode !== 'decimal') Object.assign(intl, { useGrouping: false, maximumFractionDigits: 10 });
    if (o.decimals != null) intl.minimumFractionDigits = o.decimals;
    if (o.maxDecimals != null) intl.maximumFractionDigits = Math.max(o.maxDecimals, intl.minimumFractionDigits || 0);
    try {
        return new Intl.NumberFormat(info.locale || 'fr-FR', intl).format(value).replace(/ /g, NBSP);
    } catch (e) {
        return String(value);
    }
}

// Pourcentage en fraction : colonne au format pourcentage (0,61) ou valeur en % (61)
function percentFraction(pct) {
    if (!pct || typeof pct.value !== 'number') return null;
    return pct.spec && pct.spec.opts && pct.spec.opts.numMode === 'percent' ? pct.value : pct.value / 100;
}

const BAR_FILL = 'DCEFEE', BAR_BORDER = 'C9D1D8';

// [pourcentage] [valeur sur la barre de fond] sur [objectif]
function barRuns(rPr, actual, goal, pct, info) {
    const a = typeof actual.value === 'number' ? actual.value : Number(actual.value);
    const b = typeof goal.value === 'number' ? goal.value : Number(goal.value);
    let fraction = percentFraction(pct);
    if (fraction === null) fraction = b ? a / b : 0;
    if (!isFinite(fraction)) fraction = 0;

    const text = formatNumber(actual.value, actual.spec, info);
    const width = Math.max(20, Array.from(text).length + 4);
    // Espaces normales pour le remplissage : la bordure reste continue
    const chars = Array.from(' ' + text + ' '.repeat(width));
    chars.length = width;
    const filled = Math.round(Math.min(Math.max(fraction, 0), 1) * width);
    const bar = {
        'w:rFonts': fonts(MONO_FONT), 'w:b': '<w:b w:val="0"/>', 'w:bCs': '<w:bCs w:val="0"/>',
        'w:i': null, 'w:iCs': null, 'w:color': '<w:color w:val="1F1F1F"/>', 'w:highlight': null, 'w:u': null,
        'w:bdr': `<w:bdr w:val="single" w:sz="4" w:space="0" w:color="${BAR_BORDER}"/>`,
    };
    let xml = '';
    if (pct) {
        const shown = formatNumber(fraction, { opts: { numMode: 'percent', maxDecimals: (pct.spec && pct.spec.opts && pct.spec.opts.maxDecimals) || 0 } }, info);
        xml += runXml(rPr, { 'w:color': '<w:color w:val="0A4A50"/>' }, shown + NBSP + NBSP);
    }
    if (filled > 0) xml += runXml(rPr, { ...bar, 'w:shd': `<w:shd w:val="clear" w:color="auto" w:fill="${BAR_FILL}"/>` }, chars.slice(0, filled).join(''));
    if (filled < width) xml += runXml(rPr, { ...bar, 'w:shd': '<w:shd w:val="clear" w:color="auto" w:fill="FFFFFF"/>' }, chars.slice(filled).join(''));
    xml += runXml(rPr, { 'w:color': '<w:color w:val="6B7580"/>' }, NBSP + NBSP + 'sur' + NBSP + formatNumber(goal.value, goal.spec, info));
    return xml;
}

// Étapes en ligne : ● Offre › ● Signature › ◉ Accueil › ○ Formation
function stepsRuns(rPr, choices, current, fill) {
    const accent = accentColor(fill);
    return choices.map((label, i) => {
        const state = current < 0 ? 'next' : i < current ? 'done' : i === current ? 'cur' : 'next';
        const color = state === 'next' ? '9AA3AB' : accent;
        return runXml(rPr, { 'w:rFonts': fonts(SYMBOL_FONT), 'w:color': `<w:color w:val="${color}"/>` },
            state === 'done' ? '●' : state === 'cur' ? '◉' : '○')
            + runXml(rPr, state === 'cur' ? { 'w:b': '<w:b/>', 'w:bCs': '<w:bCs/>' } : { 'w:color': `<w:color w:val="${state === 'next' ? '8A939C' : '3C4650'}"/>` }, NBSP + label);
    }).join(runXml(rPr, { 'w:color': '<w:color w:val="9AA3AB"/>' }, '  ›  '));
}

// Étapes en tableau : cercles reliés par un trait, libellés dessous
function stepsTable(choices, current, fill) {
    const accent = accentColor(fill), grey = 'C9D1D8';
    const n = choices.length, TOTAL = 9600, SYM = 420;
    const half = Math.max(120, Math.floor((TOTAL - n * SYM) / (2 * n)));
    const noBorders = '<w:tblBorders><w:top w:val="nil"/><w:left w:val="nil"/><w:bottom w:val="nil"/><w:right w:val="nil"/><w:insideH w:val="nil"/><w:insideV w:val="nil"/></w:tblBorders>';
    const tblPr = `<w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>${noBorders}<w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>`;
    const grid = '<w:tblGrid>' + choices.map(() => `<w:gridCol w:w="${half}"/><w:gridCol w:w="${SYM}"/><w:gridCol w:w="${half}"/>`).join('') + '</w:tblGrid>';
    const tight = '<w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>';
    const cell = (w, content, extra = '') => `<w:tc><w:tcPr><w:tcW w:w="${w}" w:type="dxa"/>${extra}<w:vAlign w:val="center"/></w:tcPr>${content}</w:tc>`;
    const below = '<w:tcMar><w:bottom w:w="160" w:type="dxa"/></w:tcMar>';
    // Demi-trait : paragraphe vide de 2 pt dont la bordure basse passe au milieu du cercle
    const line = color => color
        ? `<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="12" w:space="0" w:color="${color}"/></w:pBdr><w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>`
        : `<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="40" w:lineRule="exact"/><w:rPr><w:sz w:val="2"/></w:rPr></w:pPr></w:p>`;
    const state = i => (current < 0 ? 'next' : i < current ? 'done' : i === current ? 'cur' : 'next');
    const symbols = choices.map((_, i) => {
        const s = state(i);
        const left = i === 0 ? null : (current >= 0 && i <= current ? accent : grey);
        const right = i === n - 1 ? null : (current >= 0 && i + 1 <= current ? accent : grey);
        const glyph = s === 'done' ? '●' : s === 'cur' ? '◉' : '○';
        const size = s === 'cur' ? 30 : 24;
        const sym = `<w:p><w:pPr>${tight}<w:jc w:val="center"/></w:pPr><w:r><w:rPr>${fonts(SYMBOL_FONT)}<w:color w:val="${s === 'next' ? 'AEB6BE' : accent}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr><w:t>${glyph}</w:t></w:r></w:p>`;
        return cell(half, line(left)) + cell(SYM, sym) + cell(half, line(right));
    }).join('');
    const labels = choices.map((label, i) => {
        const s = state(i);
        const rPr = `<w:rPr>${s === 'cur' ? '<w:b/><w:bCs/>' : ''}<w:color w:val="${s === 'cur' ? '1F1F1F' : s === 'done' ? '3C4650' : '8A939C'}"/><w:sz w:val="17"/><w:szCs w:val="17"/></w:rPr>`;
        return cell(2 * half + SYM, `<w:p><w:pPr>${tight}<w:jc w:val="center"/></w:pPr><w:r>${rPr}<w:t xml:space="preserve">${xmlEscape(label)}</w:t></w:r></w:p>`, '<w:gridSpan w:val="3"/>' + below);
    }).join('');
    return `<w:tbl>${tblPr}${grid}<w:tr>${symbols}</w:tr><w:tr>${labels}</w:tr></w:tbl>`;
}

// ---------------------------------------------------------------------------
// Remplacement des marqueurs après le rendu docxtemplater
// ---------------------------------------------------------------------------

function applyRenders(zip, registry) {
    if (!registry.length) return;
    for (const file of docxTextFiles(zip)) {
        const xml = file.asText();
        if (xml.indexOf(MARK_OPEN) < 0) continue;
        zip.file(file.name, withTextBoxes(xml, part => replaceRuns(replaceBlocks(part, registry), registry)));
    }
}

// Paragraphe ne contenant qu'un marqueur à rendu bloc : remplacé par le bloc
function replaceBlocks(xml, registry) {
    return xml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[ />])[\s\S])*?<\/w:p>(<\/w:tc>)?/g, (paragraph, cellEnd) => {
        const texts = [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
        const m = /^\s*(\d+)\s*$/.exec(texts);
        const render = m && registry[+m[1]];
        if (!render || !render.block || /<w:(?:drawing|pict|object|sectPr)\b/.test(paragraph)) return paragraph;
        // Une cellule doit se terminer par un paragraphe
        return render.block() + (cellEnd ? '<w:p/>' + cellEnd : '');
    });
}

// Run contenant des marqueurs : découpé, chaque marqueur remplacé par ses runs
function replaceRuns(xml, registry) {
    return xml.replace(/<w:r(?:\s[^>]*)?>(?:(?!<w:r[ >])[\s\S])*?<\/w:r>/g, (run) => {
        if (run.indexOf(MARK_OPEN) < 0) return run;
        const open = /^<w:r(?:\s[^>]*)?>/.exec(run)[0];
        let inner = run.slice(open.length, -'</w:r>'.length);
        const rPrMatch = /^\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(inner);
        const rPr = rPrMatch ? rPrMatch[1] : '';
        if (rPrMatch) inner = inner.slice(rPrMatch[0].length);
        const rPrXml = rPr ? `<w:rPr>${rPr}</w:rPr>` : '';
        let out = '', pending = '';
        const flush = () => { if (pending) out += `${open}${rPrXml}${pending}</w:r>`; pending = ''; };
        const tokens = /<w:t(\s[^>]*)?>([^<]*)<\/w:t>|<[^>]+\/>|<([\w:]+)(?:\s[^>]*)?>[\s\S]*?<\/\3>/g;
        let t;
        while ((t = tokens.exec(inner)) !== null) {
            if (t[2] === undefined || t[2].indexOf(MARK_OPEN) < 0) { pending += t[0]; continue; }
            let last = 0, m;
            MARK_RE.lastIndex = 0;
            while ((m = MARK_RE.exec(t[2])) !== null) {
                const before = t[2].slice(last, m.index);
                if (before) pending += `<w:t xml:space="preserve">${before}</w:t>`;
                flush();
                const render = registry[+m[1]];
                out += render ? render.inline(rPr) : '';
                last = m.index + m[0].length;
            }
            const after = t[2].slice(last);
            if (after) pending += `<w:t xml:space="preserve">${after}</w:t>`;
        }
        flush();
        return out;
    });
}
