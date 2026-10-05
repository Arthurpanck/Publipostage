// Normalisation commune aux balises Word/PDF et aux clés reçues de Grist.
function sanitizeKey(value) {
    return String(value ?? '').normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[^a-zA-Z0-9_]+/g, '_').replace(/^_+/, '').toLowerCase();
}

function normalizeMergeData(value) {
    if (Array.isArray(value)) return value.map(normalizeMergeData);
    if (!value || Object.prototype.toString.call(value) !== '[object Object]') return value;
    const proto = Object.getPrototypeOf(value);
    if (proto && proto.constructor?.name !== 'Object') return value;
    return Object.fromEntries(Object.entries(value).map(([key, item]) =>
        [sanitizeKey(key), normalizeMergeData(item)]));
}

function normalizeDocxTag(key) {
    // Le point signifie l'élément courant ; la directive = change les délimiteurs.
    if (key === '.' || key.startsWith('=')) return key;
    const prefix = /^[#\/^@]/.test(key) ? key[0] : '';
    return prefix + key.slice(prefix.length).split('.').map(sanitizeKey).join('.');
}

// Conserver accents et casse, mais produire un nom portable, jamais un chemin.
function buildExportFilename(value, fallback, extension) {
    let name = typeof value === 'string' || typeof value === 'number' ? String(value) : '';
    name = name.normalize('NFC').replace(/[<>:"/\\|?*\x00-\x1f\x7f]/g, '_').trim();
    name = name.replace(/\.(docx|pdf)$/i, '').replace(/^[. ]+|[. ]+$/g, '');
    // Garder de la place pour l’extension et les suffixes sur les systèmes à limite en octets.
    const chars = Array.from(name);
    while (new TextEncoder().encode(chars.join('')).length > 180) chars.pop();
    name = chars.join('').replace(/[. ]+$/g, '');
    if (!name) name = fallback;
    if (/^(con|prn|aux|nul|com[1-9]|lpt[1-9])(?:\.|$)/i.test(name)) name = '_' + name;
    return `${name}.${extension}`;
}

function uniqueExportFilename(filename, usedNames) {
    const dot = filename.lastIndexOf('.');
    const stem = filename.slice(0, dot), extension = filename.slice(dot);
    let candidate = filename, suffix = 2;
    while (usedNames.has(candidate.toLowerCase())) candidate = `${stem} (${suffix++})${extension}`;
    usedNames.add(candidate.toLowerCase());
    return candidate;
}
