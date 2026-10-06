/**
 * MODULE RELATIONS
 * Version : 2.0
 * Lignes des tables citées dans un modèle Word par {Table.Colonne} ou
 * {#Table} (syntaxe : voir docx-tools.js), pour une ligne du widget :
 *   - Table du widget : lignes transmises au widget (onRecords), avec les
 *     liaisons, les filtres et le tri appliqués par Grist.
 *   - Table contenant une colonne Ref:/RefList: vers la table du widget :
 *     ses lignes liées à la ligne du widget.
 *   - À défaut, colonne Ref:/RefList: de la table du widget vers Table :
 *     les lignes qu'elle désigne, dans son ordre. C'est le cas de la colonne
 *     group d'une table regroupée.
 * Les noms du modèle sont normalisés (sanitizeKey) ; les identifiants réels
 * des tables restent utilisés pour lire Grist.
 *
 * Limites connues :
 *   - DOCX uniquement (un PDF à formulaire ne peut pas agrandir un tableau).
 *   - Si plusieurs colonnes relient les deux tables dans le même sens, la
 *     première trouvée est utilisée et un avertissement indique le choix.
 *   - Les valeurs des tables liées sont lues brutes (dates = nombre, refs = id).
 */

// Lectures partagées le temps d'un lot de génération (voir newBatch dans
// main.js). Chaque lot a son propre cache : il relit des données fraîches et
// une lecture échouée sera retentée par le lot suivant.
function cached(cache, key, compute) {
    if (!cache.has(key)) {
        cache.set(key, compute());
    }
    return cache.get(key);
}

function fetchTableCached(cache, tableId) {
    return cached(cache, tableId, () => grist.docApi.fetchTable(tableId));
}

const WIDGET_TABLE = Symbol('widgetTable');
const VIEW_ROWS = Symbol('viewRows');
const ROW_INDEX = Symbol('rowIndex');

// Liste encodée par fetchTable (["L", …]) -> valeur telle que reçue par le widget
function decodeList(value) {
    return Array.isArray(value) && value[0] === 'L' ? value.slice(1) : value;
}

// Grist peut ne transmettre au widget qu'une partie des colonnes, malgré
// includeColumns: 'normal' (version de Grist, colonnes associées du widget) :
// les colonnes absentes de la ligne reçue sont relues dans la table du widget,
// une fois par lot. Les valeurs reçues restent prioritaires. En cas d'échec de
// lecture, la ligne reçue est utilisée telle quelle.
async function completeRow(row, batch) {
    try {
        const widgetTable = await cached(batch.cache, WIDGET_TABLE, () => grist.getSelectedTableId());
        const tbl = await fetchTableCached(batch.cache, widgetTable);
        const index = cached(batch.cache, ROW_INDEX, () => new Map(tbl.id.map((id, i) => [id, i])));
        const i = index.get(row.id);
        if (i === undefined) {
            return row;
        }
        const full = { ...row };
        for (const col in tbl) {
            if (!(col in full)) full[col] = decodeList(tbl[col][i]);
        }
        return full;
    } catch (e) {
        console.warn("Impossible de compléter les colonnes de la ligne", e);
        return row;
    }
}

// Ligne i d'une table lue par fetchTable (format colonnes)
function tableRow(tbl, i) {
    const row = {};
    for (const col in tbl) {
        row[col] = tbl[col][i];
    }
    return rowData(row);
}

// Colonnes Ref:/RefList: reliant la table citée (nom normalisé) et la table du
// widget, dans les deux sens : child = colonne de la table citée vers le
// widget, sinon colonne du widget. tableId = identifiant réel à lire.
async function findLinks(cache, widgetTable, table) {
    const [cols, tables] = await Promise.all([
        fetchTableCached(cache, '_grist_Tables_column'),
        fetchTableCached(cache, '_grist_Tables'),
    ]);
    const tableIds = new Map(tables.id.map((id, i) => [id, tables.tableId[i]]));

    const links = [];
    cols.id.forEach((_, i) => {
        const [, kind, target] = /^(Ref|RefList):(.+)$/.exec(cols.type[i]) || [];
        const owner = tableIds.get(cols.parentId[i]);
        const link = { column: cols.colId[i], isList: kind === 'RefList' };
        if (target === widgetTable && sanitizeKey(owner) === table) {
            links.push({ ...link, tableId: owner, child: true });
        } else if (owner === widgetTable && target && sanitizeKey(target) === table) {
            links.push({ ...link, tableId: target, child: false });
        }
    });
    return links;
}

// Lignes de la table liée à `record`, la ligne du widget.
async function linkedRows(cache, link, record) {
    const tbl = await fetchTableCached(cache, link.tableId);

    if (link.child) {
        // Valeurs brutes de fetchTable : Ref = id, RefList = ["L", id, ...]
        return tbl.id.flatMap((_, i) => {
            const ref = tbl[link.column][i];
            const linked = link.isList
                ? Array.isArray(ref) && ref.slice(1).includes(record.id)
                : ref === record.id;
            return linked ? [tableRow(tbl, i)] : [];
        });
    }

    // Valeur décodée par Grist : Ref = id ou 0, RefList = [id, ...] ou null
    const value = record[link.column];
    if (value === undefined) {
        throw new Error(`La colonne "${link.column}" n'a pas été transmise au widget.`);
    }
    const ids = link.isList ? value || [] : [value].filter(Boolean);
    if (!Array.isArray(ids) || !ids.every(id => Number.isInteger(id) && id > 0)) {
        throw new Error(`La colonne "${link.column}" contient des références invalides.`);
    }
    const indexes = new Map(tbl.id.map((id, i) => [id, i]));
    return ids.map(id => {
        if (!indexes.has(id)) {
            throw new Error(`La ligne ${id} de "${link.tableId}" est introuvable ou inaccessible.`);
        }
        return tableRow(tbl, indexes.get(id));
    });
}

// Complète les données de `record` avec les tables citées dans le modèle
// (noms normalisés) : data devient { champs de la ligne..., membres: [..] }.
// Une lecture échouée interrompt la génération : un document incomplet ne
// doit pas être annoncé comme réussi.
async function addLinkedTables(data, record, tables, batch) {
    if (tables.length === 0) {
        return data; // aucune balise {Table.Colonne} -> aucune lecture
    }

    const widgetTable = await cached(batch.cache, WIDGET_TABLE, () => grist.getSelectedTableId());
    for (const table of tables) {
        if (table === sanitizeKey(widgetTable)) {
            data[table] = await cached(batch.cache, VIEW_ROWS, async () => {
                const rows = batch.viewRows.filter(row => row.id !== 'new');
                return (await Promise.all(rows.map(row => completeRow(row, batch)))).map(rowData);
            });
            continue;
        }

        // Le lien porté par la table citée est prioritaire (référence à double
        // sens comprise) ; à défaut, la référence portée par la ligne du widget.
        const links = await findLinks(batch.cache, widgetTable, table);
        const candidates = links.some(link => link.child) ? links.filter(link => link.child) : links;
        if (candidates.length === 0) {
            // Un {#Champ} peut être une section conditionnelle sur un champ de la ligne.
            if (!Object.keys(data).some(key => sanitizeKey(key) === table)) {
                batch.notes.add(`"${table}" n'est reliée à "${widgetTable}" par aucune colonne Référence ou Liste de références (ou nom mal orthographié).`);
            }
            continue;
        }
        if (candidates.length > 1) {
            batch.notes.add(`"${table}" et "${widgetTable}" sont reliées par plusieurs colonnes : ${candidates.map(link => link.column).join(', ')}. La colonne "${candidates[0].column}" est utilisée.`);
        }
        try {
            data[table] = await linkedRows(batch.cache, candidates[0], record);
        } catch (error) {
            throw new Error(`Impossible de lire la table liée "${table}" : ${error.message}`);
        }
    }
    return data;
}
