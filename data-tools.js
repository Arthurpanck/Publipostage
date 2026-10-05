/**
 * MODULE DONNÉES
 * Version : 2.0
 * Données de publipostage d'une ligne Grist : clés nettoyées des colonnes et
 * lignes des tables citées dans un modèle Word par {Table.Colonne}.
 *
 * Résolution de {Table.Colonne} pour une ligne du widget :
 *   - Table du widget : lignes transmises au widget, avec les liaisons, les
 *     filtres et le tri appliqués par Grist.
 *   - Table contenant une colonne Ref:/RefList: vers la table du widget :
 *     ses lignes liées à la ligne du widget.
 *   - À défaut, colonne Ref:/RefList: de la table du widget vers Table :
 *     les lignes qu'elle désigne, dans son ordre. C'est le cas de la colonne
 *     group d'une table regroupée.
 *
 * Limites connues :
 *   - Si plusieurs colonnes relient les deux tables dans le même sens, la
 *     première trouvée est utilisée et un avertissement indique le choix.
 *   - Les valeurs des tables liées sont lues brutes (dates = nombre, refs = id).
 */

// Identifiant de colonne -> nom de balise (sans accents ni caractères spéciaux)
function sanitizeKey(keytoSanitize) {
    if (!keytoSanitize) {
        return "";
    }
    let sanitize = keytoSanitize.toString();
    sanitize = sanitize.normalize('NFKD');
    sanitize = sanitize.replace(/[̀-ͯ]/g, ""); // Supprime les accents
    sanitize = sanitize.replace(/[^a-zA-Z0-9_]+/g, "_"); // Remplace caractères spéciaux
    sanitize = sanitize.replace(/^_+/, ""); // Supprime _ au début
    return sanitize;
}

// Identifiant de ligne et colonnes techniques de Grist, jamais publipostés
const TECHNICAL_COLUMN = /^(?:id$|__|manualSort|gristHelper_)/;

// Ligne Grist { colonne: valeur } -> données de publipostage
function rowData(row) {
    const data = {};
    for (const key in row) {
        if (!TECHNICAL_COLUMN.test(key)) {
            data[sanitizeKey(key)] = row[key];
        }
    }
    return data;
}

// Ligne i d'une table lue par fetchTable (format colonnes)
function tableRow(tbl, i) {
    const row = {};
    for (const col in tbl) {
        row[col] = tbl[col][i];
    }
    return rowData(row);
}

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

// Colonnes Ref:/RefList: reliant `table` et la table du widget, dans les deux
// sens : child = colonne de `table` vers le widget, sinon colonne du widget.
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
        if ((owner === table && target === widgetTable) || (owner === widgetTable && target === table)) {
            links.push({ column: cols.colId[i], isList: kind === 'RefList', child: owner === table });
        }
    });
    return links;
}

// Lignes de `table` liées à `record`, la ligne du widget.
async function linkedRows(cache, table, link, record) {
    const tbl = await fetchTableCached(cache, table);

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
            throw new Error(`La ligne ${id} de "${table}" est introuvable ou inaccessible.`);
        }
        return tableRow(tbl, indexes.get(id));
    });
}

// Complète les données de `record` avec les tables citées dans le modèle :
// data devient { champs de la ligne..., Membres: [ {..}, {..} ], ... }.
// Une lecture échouée interrompt la génération : un document incomplet ne
// doit pas être annoncé comme réussi.
async function addLinkedTables(data, record, tables, batch) {
    if (tables.length === 0) {
        return data; // aucune balise {Table.Colonne} -> aucune lecture
    }

    const widgetTable = await cached(batch.cache, WIDGET_TABLE, () => grist.getSelectedTableId());
    for (const table of tables) {
        if (table === widgetTable) {
            data[table] = cached(batch.cache, VIEW_ROWS,
                () => batch.viewRows.filter(row => row.id !== 'new').map(rowData));
            continue;
        }

        // Le lien porté par la table citée est prioritaire (référence à double
        // sens comprise) ; à défaut, la référence portée par la ligne du widget.
        const links = await findLinks(batch.cache, widgetTable, table);
        const candidates = links.some(link => link.child) ? links.filter(link => link.child) : links;
        if (candidates.length === 0) {
            // Un {#Champ} peut être une section conditionnelle sur un champ de la ligne.
            if (!(table in data)) {
                batch.notes.add(`"${table}" n'est reliée à "${widgetTable}" par aucune colonne Référence ou Liste de références (ou nom mal orthographié).`);
            }
            continue;
        }
        if (candidates.length > 1) {
            batch.notes.add(`"${table}" et "${widgetTable}" sont reliées par plusieurs colonnes : ${candidates.map(link => link.column).join(', ')}. La colonne "${candidates[0].column}" est utilisée.`);
        }
        try {
            data[table] = await linkedRows(batch.cache, table, candidates[0], record);
        } catch (error) {
            throw new Error(`Impossible de lire la table liée "${table}" : ${error.message}`);
        }
    }
    return data;
}
