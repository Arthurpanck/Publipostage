/**
 * MODULE RENDUS
 * Version : 1.0
 * Mise en forme des valeurs selon le type de la colonne Grist (DOCX) :
 *   - Booléen : case à cocher ☑ / ☐ suivie du nom de la colonne, automatique.
 *   - Choix, Choix multiples : tag teinté en semi-gras, couleur de l'option
 *     Grist, police et taille de la balise ; retour à la ligne si trop long.
 *   - Référence, Référence multiple : tag gris, valeur affichée de la fiche.
 *   - {Etape|etapes} : progression dans l'ordre des options du Choix.
 *   - {Actuel|barre:Objectif} ou {Actuel|barre:Objectif:Pourcentage} : barre
 *     de fond. Toutes les valeurs viennent de colonnes de la ligne. Au-delà
 *     de 100 %, la barre et le pourcentage passent en rouge.
 *   - Pièce jointe : images insérées, {Photo|petit} (2 cm de haut au plus),
 *     {Photo|moyen} (4 cm), {Photo} ou {Photo|grand} (9 cm).
 * Texte, nombres sans option et dates restent insérés tels quels.
 *
 * Principe : le parser docxtemplater renvoie un marqueur à la place de la
 * valeur ; après le rendu, chaque marqueur est remplacé par une image PNG
 * dessinée au canvas du navigateur (4x), fidèle à la maquette : pilules,
 * case, barre arrondie, cercles reliés. Les tailles suivent la police de la
 * balise, l'image est centrée sur la ligne (w:position) et porte le texte en
 * description (accessibilité). Sans canvas, la valeur est insérée en texte.
 */

const RENDER = Symbol('render');
const MARK_OPEN = '', MARK_CLOSE = '';
const MARK_RE = /(\d+)/g;
// Hauteur maximale des images d'une pièce jointe : 2, 4 et 9 cm (9 cm par défaut)
const IMAGE_SIZES = { petit: 76, moyen: 151, grand: 340 };
const FORMATS = ['barre', 'etapes', ...Object.keys(IMAGE_SIZES)];

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
            const visible = byRowId.get(cols.visibleCol ? cols.visibleCol[i] : 0);
            s.refs = await refDisplay(batch, meta, target, visible);
            s.visibleKey = visible === undefined ? null : sanitizeKey(cols.colId[visible]);
        }
        spec[sanitizeKey(colId)] = s;
    }
    return spec;
}

// id -> { text, style } pour les fiches de la table liée (étiquettes grises)
async function refDisplay(batch, meta, tableId, visibleIndex) {
    const { cols } = meta;
    const tbl = await fetchTableCached(batch.cache, tableId);
    const visibleCol = visibleIndex === undefined ? null : cols.colId[visibleIndex];
    const map = new Map();
    tbl.id.forEach((id, i) => {
        const raw = visibleCol ? tbl[visibleCol][i] : id;
        const text = Array.isArray(raw) ? raw.slice(raw[0] === 'L' ? 1 : 0).join(', ') : raw;
        map.set(id, { text: text == null ? '' : String(text), style: null });
    });
    return map;
}

function choiceStyle(opts, value) {
    const o = opts && opts.choiceOptions && opts.choiceOptions[value];
    return o ? { fill: o.fillColor } : null;
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
            if (format in IMAGE_SIZES && args.length) {
                throw new Error(`Taille d'image : petit, moyen ou grand, sans autre option : {${base}|${format}}.`);
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

// Renvoie { inline(rPr, env) -> runs XML, block?(env, pPr) -> XML, raw } ou null
// (valeur brute). Les rendus sont dessinés au canvas du navigateur (env.draw) ;
// sans canvas, la valeur est insérée en texte (raw).
function renderer(ctx, { value, spec, format, args, tag, lookup }) {
    const info = ctx.info || {};
    if (format === 'barre') {
        const goal = lookup(args[0]);
        const pct = args[1] ? lookup(args[1]) : null;
        if (!goal) ctx.unknownTags.add(args[0]);
        if (args[1] && !pct) ctx.unknownTags.add(args[1]);
        if (!goal || (args[1] && !pct)) return null;
        const actual = { value, spec };
        return { raw: String(value), inline: (rPr, env) => barImage(rPr, env, actual, goal, pct, info) };
    }
    if (format === 'etapes') {
        const choices = spec && spec.opts && spec.opts.choices;
        if (!Array.isArray(choices) || !choices.length) {
            ctx.unknownTags.add(`${tag} (colonne sans liste de choix)`);
            return null;
        }
        const color = (choiceStyle(spec.opts, value) || {}).fill;
        const current = choices.indexOf(value);
        return {
            raw: value == null ? '' : String(value),
            // Seule dans son paragraphe : frise sur toute la largeur ; dans une phrase : frise compacte
            inline: (rPr, env) => stepsInline(rPr, env, choices, current, color),
            block: (env, pPr) => stepsImage(env, pPr, choices, current, color),
        };
    }
    if (format in IMAGE_SIZES && (!spec || spec.type !== 'Attachments')) {
        ctx.unknownTags.add(`${tag} (taille réservée aux colonnes Pièce jointe)`);
        return null;
    }
    if (!spec) return null;
    const tags = items => ({ raw: items.map(i => i.text).join(', '), inline: (rPr, env) => tagImages(rPr, env, items) });
    switch (spec.type) {
        case 'Bool': {
            const checked = value === true || value === 1 || value === 'true';
            // La case est suivie du libellé de la colonne Grist (texte modifiable)
            return {
                raw: String(value),
                inline: (rPr, env) => checkboxImage(rPr, env, checked) + runXml(rPr, {}, NBSP + (spec.label || tag)),
            };
        }
        case 'Choice':
            return tags(value === null || value === '' ? [] : [{ text: String(value), style: choiceStyle(spec.opts, value) }]);
        case 'ChoiceList':
            return tags(listValue(value).map(v => ({ text: String(v), style: choiceStyle(spec.opts, v) })));
        case 'Attachments': {
            const ids = refIds(value);
            const files = ids.map(id => (info.files && info.files.get(id)) || null).filter(Boolean);
            const maxHeight = IMAGE_SIZES[format] || IMAGE_SIZES.grand;
            return { raw: files.map(f => f.name).join(', '), inline: (rPr, env) => attachmentRuns(rPr, env, files, maxHeight) };
        }
        case 'Ref':
        case 'RefList': {
            // Selon la provenance, une référence arrive sous la forme de sa valeur
            // affichée (ligne du widget, par défaut dans Grist), de son id (tables
            // relues), d'une fiche { id, … }, ou d'une ligne liée sans id quand la
            // colonne porte le nom de la table liée ({Lieux} avec {Lieux.Nom}).
            const text = v => {
                if (v && typeof v === 'object') {
                    if ('id' in v) return ((spec.refs && spec.refs.get(v.id)) || {}).text || '';
                    return String((spec.visibleKey && v[spec.visibleKey]) ?? '');
                }
                if (Number.isInteger(v)) return v > 0 ? ((spec.refs && spec.refs.get(v)) || { text: String(v) }).text : '';
                return v == null ? '' : String(v);
            };
            return tags(listValue(value).map(v => ({ text: text(v), style: null })).filter(r => r.text !== ''));
        }
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
// Même teinte, foncée et saturée (texte lisible sur un fond clair)
function darkSameHue(rgb) {
    const [r, g, b] = rgb.map(x => x / 255), max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
    let h = 0;
    if (d) h = max === r ? ((g - b) / d) % 6 : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h = (h * 60 + 360) % 360;
    const s = d ? 0.6 : 0, l = 0.27, c = (1 - Math.abs(2 * l - 1)) * s, x = c * (1 - Math.abs(((h / 60) % 2) - 1)), m = l - c / 2;
    const [r1, g1, b1] = h < 60 ? [c, x, 0] : h < 120 ? [x, c, 0] : h < 180 ? [0, c, x] : h < 240 ? [0, x, c] : h < 300 ? [x, 0, c] : [c, 0, x];
    return [r1 + m, g1 + m, b1 + m].map(v => v * 255);
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

const NBSP = ' ';

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

const BAR_FILL = 'DCEFEE';
// Au-delà de l'objectif : barre pleine en rouge clair, pourcentage en rouge
const OVER = { fill: 'FCE1DD', border: 'E5948A', text: '8A1C12', pct: 'B42318' };
const isOver = fraction => fraction > 1 + 1e-9;

// ---------------------------------------------------------------------------
// Rendus dessinés (canvas) : mêmes proportions que la maquette, à 4x pour
// l'impression. Les tailles suivent la taille de police de la balise.
// ---------------------------------------------------------------------------

const DRAW_SCALE = 4;
const SANS = '"Segoe UI", "Helvetica Neue", Arial, "Liberation Sans", sans-serif';
const MONO = 'Consolas, "SFMono-Regular", Menlo, "DejaVu Sans Mono", "Liberation Mono", monospace';
const BODY_PX = 14.667; // 11 pt, taille de référence de la maquette

function canvasDrawer() {
    if (typeof document === 'undefined' || !document || typeof document.createElement !== 'function') return null;
    let probe;
    try { probe = document.createElement('canvas').getContext('2d'); } catch (e) { return null; }
    if (!probe) return null;
    return {
        measure(font, text) { probe.font = font; return probe.measureText(text).width; },
        paint(w, h, draw) {
            const c = document.createElement('canvas');
            c.width = Math.ceil(w * DRAW_SCALE); c.height = Math.ceil(h * DRAW_SCALE);
            const g = c.getContext('2d');
            g.scale(DRAW_SCALE, DRAW_SCALE);
            draw(g);
            return { data: c.toDataURL('image/png'), w, h };
        },
    };
}

// Taille de police de la balise, en px (11 pt par défaut)
function fontPx(rPr) {
    const m = /<w:sz w:val="(\d+)"/.exec(rPr || '');
    return (m ? +m[1] / 2 : 11) * 96 / 72;
}

// Police réellement appliquée à une balise, résolue comme Word : valeurs par
// défaut du document, style de paragraphe, style de caractère, puis mise en
// forme directe. Polices du thème (« Corps », « Titres ») lues dans le thème.
function fontResolver(zip) {
    const read = name => (zip.file(name) ? zip.file(name).asText() : '');
    const styles = read('word/styles.xml'), theme = read('word/theme/theme1.xml');
    const themeFont = kind => (new RegExp(`<a:${kind}Font>\\s*<a:latin typeface="([^"]*)"`).exec(theme) || [])[1];
    const fonts = { minor: themeFont('minor'), major: themeFont('major') };
    const props = (rPr) => {
        const out = {};
        const sz = /<w:sz w:val="(\d+)"/.exec(rPr || ''), f = /<w:rFonts\b[^>]*\/>/.exec(rPr || '');
        if (sz) out.px = +sz[1] / 2 * 96 / 72;
        if (f) {
            const ascii = /w:ascii="([^"]+)"/.exec(f[0]), themed = /w:asciiTheme="([^"]+)"/.exec(f[0]);
            if (ascii) out.family = ascii[1];
            else if (themed) out.family = fonts[/major/i.test(themed[1]) ? 'major' : 'minor'];
        }
        return out;
    };
    const byId = new Map();
    let defaultParagraph = null;
    for (const m of styles.matchAll(/<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g)) {
        const id = (/w:styleId="([^"]+)"/.exec(m[1]) || [])[1];
        if (!id) continue;
        const rPr = (/<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(m[2]) || [])[1];
        byId.set(id, { basedOn: (/<w:basedOn w:val="([^"]+)"/.exec(m[2]) || [])[1], props: props(rPr) });
        if (/w:type="paragraph"/.test(m[1]) && /w:default="(?:1|true)"/.test(m[1])) defaultParagraph = id;
    }
    const chain = (id, seen = new Set()) => {
        const st = id && byId.get(id);
        if (!st || seen.has(id)) return {};
        seen.add(id);
        return { ...chain(st.basedOn, seen), ...st.props };
    };
    const defaults = props((/<w:rPrDefault>\s*<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(styles) || [])[1]);
    return (pStyle, rPr) => {
        const rStyle = (/<w:rStyle w:val="([^"]+)"/.exec(rPr || '') || [])[1];
        const f = { px: 11 * 96 / 72, family: fonts.minor, ...defaults, ...chain(pStyle || defaultParagraph), ...chain(rStyle), ...props(rPr) };
        return { px: f.px, family: f.family ? `"${f.family.replace(/"/g, '')}", ${SANS}` : SANS };
    };
}

function roundRect(g, x, y, w, h, r) {
    r = Math.min(r, h / 2, w / 2);
    g.beginPath();
    g.moveTo(x + r, y);
    g.arcTo(x + w, y, x + w, y + h, r); g.arcTo(x + w, y + h, x, y + h, r);
    g.arcTo(x, y + h, x, y, r); g.arcTo(x, y, x + w, y, r);
    g.closePath();
}
const css = hex => '#' + hex;

// Coupe un texte en lignes d'au plus maxW px (mots, puis caractères si un mot dépasse)
function wrapText(draw, font, text, maxW) {
    const lines = [];
    let line = '';
    for (const word of String(text).split(/\s+/).filter(Boolean)) {
        const next = line ? line + ' ' + word : word;
        if (draw.measure(font, next) <= maxW || !line) line = next; else { lines.push(line); line = word; }
        while (draw.measure(font, line) > maxW && line.length > 1) {
            let cut = line.length - 1;
            while (cut > 1 && draw.measure(font, line.slice(0, cut)) > maxW) cut--;
            lines.push(line.slice(0, cut)); line = line.slice(cut);
        }
    }
    if (line) lines.push(line);
    return lines.length ? lines : [''];
}

// Étiquette en pilule : police et taille de la balise en semi-gras, marges
// 0,3 × 0,75 em, fond et texte teintés. Trop large pour la place disponible
// (page ou cellule), elle passe sur plusieurs lignes.
function drawTag(draw, text, style, font, maxW = Infinity) {
    const fs = font.px, css600 = `600 ${fs}px ${font.family}`;
    const padX = 0.75 * fs, padY = 0.3 * fs, lineH = 1.25 * fs;
    const lines = wrapText(draw, css600, text, Math.max(fs, maxW - 2 * padX));
    const w = Math.max(...lines.map(l => draw.measure(css600, l))) + 2 * padX;
    const h = lines.length * lineH + 2 * padY;
    const { bg, fg } = tagColors(style && style.fill);
    const img = draw.paint(w, h, g => {
        roundRect(g, 0, 0, w, h, lines.length > 1 ? lineH / 2 + padY : h / 2); g.fillStyle = css(bg); g.fill();
        g.font = css600; g.fillStyle = css(fg); g.textBaseline = 'middle';
        lines.forEach((l, i) => g.fillText(l, padX, padY + lineH * (i + 0.5) + fs * 0.04));
    });
    // Ligne de base de la première ligne alignée sur celle du texte
    img.drop = h - (padY + lineH / 2 + 0.35 * fs);
    return img;
}

// Case carrée fine, coche tracée (comme ☑ / ☐)
function drawCheckbox(draw, checked, px) {
    const s = Math.round(px * 0.86 * 2) / 2, lw = Math.max(1, s * 0.08);
    return draw.paint(s, s, g => {
        roundRect(g, lw / 2, lw / 2, s - lw, s - lw, s * 0.12);
        g.lineWidth = lw; g.strokeStyle = '#1F1F1F'; g.stroke();
        if (!checked) return;
        g.beginPath(); g.moveTo(s * 0.24, s * 0.52); g.lineTo(s * 0.43, s * 0.71); g.lineTo(s * 0.77, s * 0.29);
        g.lineWidth = lw * 1.5; g.lineCap = 'round'; g.lineJoin = 'round'; g.stroke();
    });
}

// [61 %] [ 2 450,00 €  ░░░░] sur 4 000,00 € : mono 10,5 pt, barre 260 px arrondie
function drawBar(draw, text, fraction, pctText, goalText, px, maxW = Infinity) {
    const k = px / BODY_PX, fs = 14 * k, font = `400 ${fs}px ${MONO}`;
    const padX = 10 * k, padY = 5 * k, gap = 10 * k, h = Math.round(fs * 1.15 + 2 * padY);
    const pctW = pctText ? Math.max(3.2 * fs, draw.measure(font, pctText)) : 0;
    const goal = 'sur ' + goalText;
    const around = (pctText ? pctW + gap : 0) + gap + draw.measure(font, goal);
    // 260 px comme la maquette, réduite si la place manque (cellule de tableau)
    const minBar = draw.measure(font, text) + 2 * padX + 10 * k;
    const barW = Math.max(minBar, Math.min(260 * k, maxW - around));
    const w = around + barW;
    const x0 = pctText ? pctW + gap : 0;
    const f = Math.min(Math.max(fraction, 0), 1), over = isOver(fraction);
    return draw.paint(w, h, g => {
        g.textBaseline = 'middle';
        if (pctText) {
            g.font = over ? `600 ${fs}px ${MONO}` : font;
            g.fillStyle = over ? css(OVER.pct) : '#0A4A50'; g.textAlign = 'right'; g.fillText(pctText, pctW, h / 2); g.textAlign = 'left';
        }
        g.font = font;
        roundRect(g, x0 + 0.5, 0.5, barW - 1, h - 1, 4 * k);
        g.save(); g.clip(); g.fillStyle = '#FFFFFF'; g.fillRect(x0, 0, barW, h);
        g.fillStyle = css(over ? OVER.fill : BAR_FILL); g.fillRect(x0, 0, barW * f, h); g.restore();
        g.lineWidth = 1; g.strokeStyle = over ? css(OVER.border) : '#D3DAE0'; g.stroke();
        g.fillStyle = over ? css(OVER.text) : '#1F1F1F'; g.fillText(text, x0 + padX, h / 2);
        g.fillStyle = '#6B7580'; g.fillText(goal, x0 + barW + gap, h / 2);
    });
}

// Cercles reliés : 20 px, bordure 2 px, ✓ dans les étapes franchies, halo sur
// l'étape en cours, libellés 8,5 pt dessous. Largeur : celle du paragraphe.
function drawSteps(draw, choices, current, fill, width) {
    const n = choices.length, D = 20, R = D / 2, halo = 3;
    const lfs = 11.33, top = halo + 1, cy = top + R, ly = top + D + 5;
    const h = Math.ceil(ly + lfs * 1.35 + 2);
    const accent = css(accentColor(fill)), grey = '#D5DADF', haloColor = css(toHex(mix(hexToRgb(accentColor(fill)), [255, 255, 255], 0.8)));
    const step = width / n, x = i => (i + 0.5) * step;
    const state = i => (current < 0 ? 'next' : i < current ? 'done' : i === current ? 'cur' : 'next');
    return draw.paint(width, h, g => {
        for (let i = 1; i < n; i++) {
            g.fillStyle = current >= 0 && i <= current ? accent : grey;
            g.fillRect(x(i - 1), cy - 1, x(i) - x(i - 1), 2);
        }
        choices.forEach((label, i) => {
            const s = state(i);
            if (s === 'cur') { g.beginPath(); g.arc(x(i), cy, R + halo, 0, 2 * Math.PI); g.fillStyle = haloColor; g.fill(); }
            g.beginPath(); g.arc(x(i), cy, R - 1, 0, 2 * Math.PI);
            g.fillStyle = s === 'done' ? accent : '#FFFFFF'; g.fill();
            g.lineWidth = 2; g.strokeStyle = s === 'next' ? grey : accent; g.stroke();
            if (s === 'done') {
                g.beginPath(); g.moveTo(x(i) - 4.2, cy + 0.2); g.lineTo(x(i) - 1.2, cy + 3.2); g.lineTo(x(i) + 4.6, cy - 3.4);
                g.lineWidth = 1.8; g.strokeStyle = '#FFFFFF'; g.lineCap = 'round'; g.lineJoin = 'round'; g.stroke();
            }
            g.font = `${s === 'cur' ? 700 : 400} ${lfs}px ${SANS}`;
            g.fillStyle = s === 'cur' ? '#1F1F1F' : '#777777'; g.textAlign = 'center'; g.textBaseline = 'top';
            let t = String(label);
            while (t.length > 1 && g.measureText(t).width > step - 4) t = t.slice(0, -2) + '…';
            g.fillText(t, x(i), ly);
        });
    });
}

// Couleurs des étiquettes de la maquette : fond à 84 % de blanc, texte à 40 %
// de noir. Une couleur Grist déjà claire sert de fond telle quelle, avec un
// texte foncé de la même teinte.
function tagColors(fill) {
    const rgb = hexToRgb(fill);
    if (!rgb) return { bg: 'ECEEF1', fg: '2B2F36' };
    if (luminance(rgb) > 0.55) return { bg: toHex(rgb), fg: toHex(darkSameHue(rgb)) };
    return { bg: toHex(mix(rgb, [255, 255, 255], 0.84)), fg: toHex(mix(rgb, [0, 0, 0], 0.4)) };
}

// Propriétés du texte de la balise qui n'ont pas de sens sur une image
const NOT_ON_IMAGES = { 'w:shd': null, 'w:highlight': null, 'w:u': null, 'w:bdr': null };

// Image en ligne, centrée sur le texte (abaissée avec w:position)
function imageRun(rPr, env, img, alt) {
    const rel = env.media.add({ key: img.data, data: img.data, ext: 'png' });
    const px = env.font ? env.font.px : fontPx(rPr);
    const drop = img.drop !== undefined ? img.drop : img.h / 2 - 0.33 * px;
    const lower = Math.max(0, Math.round(drop * 0.75 * 2));
    const props = mergeRPr(rPr, { ...NOT_ON_IMAGES, 'w:position': lower ? `<w:position w:val="-${lower}"/>` : null });
    return `<w:r>${props}${drawingXml(rel, img, alt)}</w:r>`;
}

function drawingXml(rel, img, alt) {
    const cx = Math.round(img.w * 9525), cy = Math.round(img.h * 9525);
    return '<w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" distT="0" distB="0" distL="0" distR="0">'
        + `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/>`
        + `<wp:docPr id="${rel.id}" name="Rendu ${rel.id}" descr="${xmlEscape(alt).replace(/"/g, '&quot;')}"/>`
        + '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/></wp:cNvGraphicFramePr>'
        + '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + '<pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        + `<pic:nvPicPr><pic:cNvPr id="${rel.id}" name="rendu_${rel.id}.png"/><pic:cNvPicPr/></pic:nvPicPr>`
        + `<pic:blipFill><a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rel.rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>`
        + `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>`
        + '</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>';
}

function tagImages(rPr, env, items) {
    const font = env.font || { px: fontPx(rPr), family: SANS };
    return items.map(item => imageRun(rPr, env, drawTag(env.draw, item.text, item.style, font, env.width()), item.text))
        .join(runXml(rPr, {}, ' '));
}

function checkboxImage(rPr, env, checked) {
    return imageRun(rPr, env, drawCheckbox(env.draw, checked, env.font ? env.font.px : fontPx(rPr)), checked ? 'Case cochée' : 'Case non cochée');
}

function barImage(rPr, env, actual, goal, pct, info) {
    const a = Number(actual.value), b = Number(goal.value);
    let fraction = percentFraction(pct);
    if (fraction === null) fraction = b ? a / b : 0;
    if (!isFinite(fraction)) fraction = 0;
    const pctText = pct ? formatNumber(fraction, { opts: { numMode: 'percent', maxDecimals: (pct.spec && pct.spec.opts && pct.spec.opts.maxDecimals) || 0 } }, info) : null;
    const text = formatNumber(actual.value, actual.spec, info), goalText = formatNumber(goal.value, goal.spec, info);
    const img = drawBar(env.draw, text, fraction, pctText, goalText, env.font ? env.font.px : fontPx(rPr), env.width());
    return imageRun(rPr, env, img, `${pctText ? pctText + ' ' : ''}${text} sur ${goalText}`);
}

const stepsAlt = (choices, current) => choices.map((c, i) => (i === current ? `[${c}]` : c)).join(' > ');

// Dans une phrase : chaque étape prend la place de son libellé
function stepsInline(rPr, env, choices, current, fill) {
    const font = `700 11.33px ${SANS}`;
    const step = Math.min(100, Math.max(64, ...choices.map(c => env.draw.measure(font, String(c)) + 14)));
    const img = drawSteps(env.draw, choices, current, fill, Math.min(step * choices.length, env.width()));
    return imageRun(rPr, env, img, stepsAlt(choices, current));
}

// Seule dans son paragraphe : frise sur toute la largeur, dans le paragraphe d'origine
function stepsImage(env, pPr, choices, current, fill) {
    const img = drawSteps(env.draw, choices, current, fill, env.width());
    const rel = env.media.add({ key: img.data, data: img.data, ext: 'png' });
    return `<w:p>${pPr}<w:r>${drawingXml(rel, img, stepsAlt(choices, current))}</w:r></w:p>`;
}

// Pièces jointes : images à leur taille, réduites à la largeur disponible (page
// ou cellule de tableau) et à la hauteur choisie (petit 2 cm, moyen 4 cm,
// grand 9 cm par défaut) ; autres fichiers : étiquette du nom.
function attachmentRuns(rPr, env, files, maxHeight = IMAGE_SIZES.grand) {
    return files.map(file => {
        if (!file.image) return tagImages(rPr, env, [{ text: file.name, style: null }]);
        const scale = Math.min(1, env.width() / file.w, maxHeight / file.h);
        const img = { w: Math.round(file.w * scale), h: Math.round(file.h * scale) };
        const rel = env.media.add({ key: 'pj:' + file.id, bytes: file.bytes, ext: file.ext });
        return `<w:r>${mergeRPr(rPr, NOT_ON_IMAGES)}${drawingXml(rel, img, file.name)}</w:r>`;
    }).join(runXml(rPr, {}, ' '));
}

// Télécharge les pièces jointes citées par les colonnes de type Pièce jointe
// (ligne et tables liées), une fois par lot : info.files = id -> fichier.
async function loadAttachments(batch, data, info) {
    if (!info) return;
    const wanted = new Set();
    const scan = (row, spec) => {
        if (!row || !spec) return;
        for (const key in row) {
            const s = spec[sanitizeKey(key)];
            if (s && s.type === 'Attachments') refIds(row[key]).forEach(id => wanted.add(id));
        }
    };
    scan(data, info.root);
    for (const [key, spec] of info.byKey) {
        // data[clé normalisée] : lignes liées (une colonne homonyme garde son nom d'origine)
        if (Array.isArray(data[key])) data[key].forEach(row => scan(row, spec));
    }
    if (!wanted.size) return;
    const meta = await fetchTableCached(batch.cache, '_grist_Attachments');
    const index = new Map(meta.id.map((id, i) => [id, i]));
    info.files = info.files || new Map();
    await Promise.all([...wanted].map(async (id) => {
        const i = index.get(id);
        if (i === undefined) return;
        const name = meta.fileName[i] || `piece-jointe-${id}`;
        const ext = (/\.([a-z0-9]+)$/i.exec(name) || [])[1];
        const w = meta.imageWidth ? meta.imageWidth[i] : 0, h = meta.imageHeight ? meta.imageHeight[i] : 0;
        const file = { id, name, ext: ext && ext.toLowerCase(), w, h, image: false };
        if (IMAGE_TYPES[file.ext] && w > 0 && h > 0) {
            try {
                file.bytes = new Uint8Array(await cached(batch.cache, 'pj:' + id, () => downloadAttachment(id)));
                file.image = true;
                if (file.ext === 'jpeg') file.ext = 'jpg';
            } catch (e) {
                batch.notes && batch.notes.add(`Pièce jointe « ${name} » illisible : son nom est inséré à la place.`);
            }
        }
        info.files.set(id, file);
    }));
}

async function downloadAttachment(id) {
    const token = await grist.docApi.getAccessToken({ readOnly: true });
    const response = await fetch(`${token.baseUrl}/attachments/${id}/download?auth=${token.token}`);
    if (!response.ok) throw new Error(`Pièce jointe ${id} : ${response.status}`);
    return response.arrayBuffer();
}

const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp' };

// Images ajoutées au DOCX : word/media, relations de la partie, types de contenu.
// media : { key (doublons), data (dataURL) ou bytes, ext }
function mediaStore(zip) {
    let next = 0;
    const files = new Map(); // clé -> chemin du média
    const rels = new Map();  // partie -> { xml, ids: Map(chemin -> rId) }
    const declareType = (ext) => {
        const types = zip.file('[Content_Types].xml');
        if (types && !new RegExp(`Extension="${ext}"`, 'i').test(types.asText())) {
            zip.file('[Content_Types].xml', types.asText().replace('</Types>', `<Default Extension="${ext}" ContentType="${IMAGE_TYPES[ext]}"/></Types>`));
        }
    };
    return {
        forPart(part) {
            const relsPath = part.replace(/^word\/(.*)$/, 'word/_rels/$1.rels');
            return {
                add(media) {
                    if (!files.has(media.key)) {
                        const name = `media/rendu_${files.size + 1}.${media.ext}`;
                        let bytes = media.bytes;
                        if (!bytes) {
                            const bin = atob(media.data.slice(media.data.indexOf(',') + 1));
                            bytes = new Uint8Array(bin.length);
                            for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
                        }
                        declareType(media.ext);
                        zip.file('word/' + name, bytes);
                        files.set(media.key, name);
                    }
                    const target = files.get(media.key);
                    if (!rels.has(relsPath)) {
                        const f = zip.file(relsPath);
                        const xml = f ? f.asText() : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
                        // Word écrit parfois <Relationships …/> quand la partie n'a encore aucune relation
                        rels.set(relsPath, { xml: xml.replace(/<Relationships([^>]*?)\s*\/>/, '<Relationships$1></Relationships>'), ids: new Map() });
                    }
                    const r = rels.get(relsPath);
                    if (!r.ids.has(target)) {
                        const rId = `rIdRendu${r.ids.size + 1}`;
                        r.ids.set(target, rId);
                        r.xml = r.xml.replace('</Relationships>', `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="${target}"/></Relationships>`);
                    }
                    return { rId: r.ids.get(target), id: 4000 + (++next) };
                },
            };
        },
        flush() { for (const [path, r] of rels) if (r.ids.size) zip.file(path, r.xml); },
    };
}

// Largeur utile en px : cellule qui contient le paragraphe, sinon la page
function textWidthPx(zip) {
    const doc = zip.file('word/document.xml');
    const xml = doc ? doc.asText() : '';
    const sect = xml.slice(xml.lastIndexOf('<w:sectPr'));
    const pg = /<w:pgSz[^>]*\bw:w="(\d+)"/.exec(sect);
    const left = /<w:pgMar[^>]*\bw:left="(\d+)"/.exec(sect), right = /<w:pgMar[^>]*\bw:right="(\d+)"/.exec(sect);
    const twips = pg ? +pg[1] - (left ? +left[1] : 1134) - (right ? +right[1] : 1134) : 9638;
    return Math.max(240, twips / 15);
}
// Style du paragraphe qui contient la position
function paragraphStyle(xml, offset) {
    const start = Math.max(xml.lastIndexOf('<w:p>', offset), xml.lastIndexOf('<w:p ', offset));
    if (start < 0) return null;
    const m = /<w:pStyle w:val="([^"]+)"/.exec(xml.slice(start, offset));
    return m ? m[1] : null;
}

function cellWidthPx(xml, offset, fallback) {
    const open = Math.max(xml.lastIndexOf('<w:tc>', offset), xml.lastIndexOf('<w:tc ', offset));
    if (open < 0 || xml.lastIndexOf('</w:tc>', offset) > open) return fallback;
    const tcW = /<w:tcW\b[^>]*\/>/.exec(xml.slice(open, offset));
    const m = tcW && /w:type="dxa"/.test(tcW[0]) && /w:w="(\d+)"/.exec(tcW[0]);
    return m ? Math.max(120, (+m[1] - 216) / 15) : fallback * 0.5;
}

// ---------------------------------------------------------------------------
// Remplacement des marqueurs après le rendu docxtemplater
// ---------------------------------------------------------------------------

function applyRenders(zip, registry) {
    if (!registry.length) return;
    const draw = canvasDrawer();
    const media = draw ? mediaStore(zip) : null;
    const pageWidth = draw ? textWidthPx(zip) : 0;
    for (const file of docxTextFiles(zip)) {
        const xml = file.asText();
        if (xml.indexOf(MARK_OPEN) < 0) continue;
        const env = { draw, media: media && media.forPart(file.name), width: () => pageWidth, fontOf: draw ? fontResolver(zip) : null };
        zip.file(file.name, withTextBoxes(xml, part => replaceRuns(replaceBlocks(part, registry, env), registry, env)));
    }
    if (media) media.flush();
}

// Paragraphe ne contenant qu'un marqueur à rendu bloc : remplacé par le bloc
function replaceBlocks(xml, registry, env) {
    return xml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[ />])[\s\S])*?<\/w:p>(<\/w:tc>)?/g, (paragraph, cellEnd, offset) => {
        const texts = [...paragraph.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)].map(m => m[1]).join('');
        const m = /^\s*(\d+)\s*$/.exec(texts);
        const render = m && registry[+m[1]];
        if (!render || !render.block || !env.draw || /<w:(?:drawing|pict|object|sectPr)\b/.test(paragraph)) return paragraph;
        // Image dans le paragraphe d'origine (alignement, espacement conservés)
        const pPr = (/^<w:p(?:\s[^>]*)?>\s*(<w:pPr>[\s\S]*?<\/w:pPr>)/.exec(paragraph) || [])[1] || '';
        const width = cellWidthPx(xml, offset, env.width());
        return render.block({ ...env, width: () => width }, pPr) + (cellEnd || '');
    });
}

// Run contenant des marqueurs : découpé, chaque marqueur remplacé par ses runs
function replaceRuns(xml, registry, env) {
    return xml.replace(/<w:r(?:\s[^>]*)?>(?:(?!<w:r[ >])[\s\S])*?<\/w:r>/g, (run, offset) => {
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
                const local = env.draw ? { ...env, width: () => cellWidthPx(xml, offset, env.width()), font: env.fontOf(paragraphStyle(xml, offset), rPr) } : env;
                out += !render ? '' : env.draw ? render.inline(rPr, local) : (render.raw ? runXml(rPr, {}, render.raw) : '');
                last = m.index + m[0].length;
            }
            const after = t[2].slice(last);
            if (after) pending += `<w:t xml:space="preserve">${after}</w:t>`;
        }
        flush();
        return out;
    });
}
