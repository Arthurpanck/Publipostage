/**
 * MODULE RELATIONS
 * Version : 1.6
 * Publipostage relationnel parent -> enfants pour les modèles Word (.docx)
 *
 * Syntaxe dans le modèle :
 *   - Champ du parent (comme avant, n'importe où)        : {Nom}
 *   - Champ d'une table enfant, DANS une ligne de tableau : {Membres.Nom_Membre}
 *     "Membres" = nom exact de la table enfant, qui doit contenir une colonne
 *     de type Ref: ou RefList: vers la table du widget. La ligne de tableau
 *     contenant des balises {Table.Colonne} est répétée pour chaque
 *     enregistrement enfant lié.
 *
 * Limites connues :
 *   - DOCX uniquement (un PDF à formulaire ne peut pas agrandir un tableau).
 *   - Une seule table enfant par ligne de tableau répétée.
 *   - Si la table enfant a plusieurs colonnes Ref vers le parent, la première
 *     trouvée est utilisée.
 *   - Les valeurs encodées (dates = nombre, refs = id) sont affichées brutes.
 */

// Balise pointée {Table.Colonne} (accents tolérés côté colonne, nettoyés ensuite)
const DOTTED_TAG_SRC = '\\{([A-Za-z0-9_]+)\\.([A-Za-z0-9_À-ÿ]+)\\}';

// Cache des fetchs (métadonnées + tables enfants), le temps d'un lot de
// génération. Vidé à chaque événement Grist (voir main.js) pour ne pas servir
// de données périmées. Évite de re-télécharger la table enfant à chaque ligne
// lors d'un export ZIP.
const relationsCache = new Map();

function clearRelationsCache() {
    relationsCache.clear();
}

function fetchTableCached(tableId) {
    if (!relationsCache.has(tableId)) {
        relationsCache.set(tableId, grist.docApi.fetchTable(tableId));
    }
    return relationsCache.get(tableId);
}

// Quelles tables enfants sont réellement citées dans le modèle ?
// Fetch paresseux : aucune balise pointée => aucun fetch, comportement de base.
let templateScanCache = { buffer: null, tables: [] };
function getReferencedTables(buffer) {
    if (templateScanCache.buffer === buffer) {
        return templateScanCache.tables;
    }
    const zip = new PizZip(buffer);
    const file = zip.file("word/document.xml");
    let tables = [];
    if (file) {
        let xml = file.asText();
        // reconstruction des balises éclatées entre deux runs Word avant de scanner
        xml = xml.replace(/<\/w:t><\/w:r><w:r[^>]*><w:t[^>]*>/g, "");

        const found = new Set();
        const re = new RegExp(DOTTED_TAG_SRC, 'g');
        let m;
        while ((m = re.exec(xml)) !== null) {
            found.add(m[1]);
        }
        tables = [...found];
    }
    templateScanCache = { buffer: buffer, tables: tables };
    return tables;
}

// Quelles colonnes référencent la table `cible` ? (via les métadonnées Grist)
// Renvoie [{ table: "Membres", column: "Association", isList: false }, ...]
async function findTablesReferencing(cible) {
    const [cols, tables] = await Promise.all([
        fetchTableCached('_grist_Tables_column'),
        fetchTableCached('_grist_Tables'),
    ]);

    const nomTable = {};
    tables.id.forEach((id, i) => { nomTable[id] = tables.tableId[i]; });

    const cibles = [`Ref:${cible}`, `RefList:${cible}`];
    const res = [];
    cols.id.forEach((_, i) => {
        const type = cols.type[i];
        if (cibles.includes(type)) {
            res.push({
                table: nomTable[cols.parentId[i]],
                column: cols.colId[i],
                isList: type.startsWith('RefList:'),
            });
        }
    });
    return res;
}

// Lignes de la table enfant liées à l'enregistrement parent courant.
// Gère le lien Ref simple (nombre = id parent) ou RefList (["L", id, ...]).
async function fetchChildRows(childTable, childCol, isList, parentId) {
    const tbl = await fetchTableCached(childTable);

    const lignes = [];
    tbl.id.forEach((rowId, i) => {
        const ref = tbl[childCol][i];
        const lie = isList
            ? Array.isArray(ref) && ref.slice(1).includes(parentId) // saute le "L"
            : ref === parentId;
        if (!lie) {
            return;
        }

        const ligne = {};
        for (const col in tbl) {
            if (col === 'id' || col.startsWith('manualSort') || col.startsWith('gristHelper_')) {
                continue;
            }
            ligne[sanitizeKey(col)] = tbl[col][i];
        }
        lignes.push(ligne);
    });
    return lignes;
}

// Complète les données du parent avec les tables enfants citées dans le modèle :
// data devient { champs_du_parent..., Membres: [ {..}, {..} ], ... }
async function addChildTablesData(data, parentId, buffer) {
    const used = getReferencedTables(buffer);
    if (used.length === 0) {
        return data; // aucune balise {Table.Colonne} -> aucun fetch
    }

    const cible = await grist.getSelectedTableId(); // table du widget (le parent)
    const refs = await findTablesReferencing(cible);

    for (const table of used) {
        const link = refs.find(r => r.table === table);
        if (!link) {
            console.warn(`La table "${table}" ne référence pas "${cible}" — boucle vide.`);
            data[table] = [];
            continue;
        }
        data[table] = await fetchChildRows(link.table, link.column, link.isList, parentId);
    }
    return data;
}

// Traduction de syntaxe au moment de générer :
// {Membres.Col} dans une ligne de tableau -> {#Membres}{Col}...{/Membres}
// enroulant toute la ligne, pour que docxtemplater la répète.
// ⚠ À appeler AVANT la sanitisation des clés (sinon le "." devient "_").
function transformDottedLoops(zip) {
    const file = zip.file("word/document.xml");
    if (!file) {
        return;
    }
    let xml = file.asText();

    // reconstruction des balises éclatées sur plusieurs runs Word
    xml = xml.replace(/<\/w:t><\/w:r><w:r[^>]*><w:t[^>]*>/g, "");

    xml = xml.replace(/<w:tr\b[\s\S]*?<\/w:tr>/g, (row) => {
        const tablesInRow = new Set();
        const tagRe = new RegExp(DOTTED_TAG_SRC, 'g');
        let m;
        while ((m = tagRe.exec(row)) !== null) {
            tablesInRow.add(m[1]);
        }

        if (tablesInRow.size === 0) {
            return row; // ligne normale, on ne touche pas
        }
        if (tablesInRow.size > 1) {
            console.warn("Plusieurs tables enfants dans la même ligne, seule la première est répétée :", [...tablesInRow]);
        }
        const table = [...tablesInRow][0];

        // {Table.Col} -> {Col} (on entre dans le contexte de la boucle)
        let newRow = row.replace(
            new RegExp(`\\{${table}\\.([A-Za-z0-9_À-ÿ]+)\\}`, 'g'),
            '{$1}'
        );
        // Enrouler la ligne : {#table} dans le 1er nœud texte, {/table} dans le dernier.
        // (?:\s[^>]*)? cible uniquement <w:t>, pas <w:tr>/<w:tc>/<w:tbl>...
        newRow = newRow
            .replace(/(<w:t(?:\s[^>]*)?>)/, `$1{#${table}}`)
            .replace(/(<\/w:t>)(?![\s\S]*<\/w:t>)/, `{/${table}}$1`);
        return newRow;
    });

    zip.file("word/document.xml", xml);
}
