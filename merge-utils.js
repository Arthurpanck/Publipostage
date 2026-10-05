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
