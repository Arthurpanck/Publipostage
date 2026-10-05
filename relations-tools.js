/**
 * MODULE RELATIONS
 * Version : 1.6
 * Publipostage relationnel parent -> enfants pour les modèles Word (.docx)
 *
 * Syntaxe dans le modèle :
 *   - Champ du parent (comme avant, n'importe où)        : {Nom}
 *   - Champ d'une table enfant : {Membres.Nom_Membre}
 *     "Membres" = nom exact de la table enfant, qui doit contenir une colonne
 *     de type Ref: ou RefList: vers la table du widget.
 *     Dans une ligne de tableau : la ligne est répétée pour chaque enfant lié.
 *     Hors tableau : le paragraphe est répété pour chaque enfant lié.
 *   - Boucle manuelle docxtemplater : {#Membres}{Nom_Membre}{/Membres}
 *     (le fetch des enfants est déclenché aussi par ces marqueurs).
 *   - Depuis un regroupement Grist : {TableSource.Colonne} liste les lignes
 *     désignées par sa colonne group, sans recalculer le regroupement.
 *   - Si TableSource est la table du widget : {TableSource.Colonne} liste
 *     les lignes transmises au widget, avec les filtres et le tri Grist.
 *
 * Limites connues :
 *   - DOCX uniquement (un PDF à formulaire ne peut pas agrandir un tableau).
 *   - Une seule table enfant par ligne de tableau répétée.
 *   - Si la table enfant a plusieurs colonnes Ref vers le parent, la première
 *     trouvée est utilisée et un avertissement indique le choix.
 *   - Les valeurs encodées (dates = nombre, refs = id) sont affichées brutes.
 */

// Balise pointée {Table.Colonne} (accents tolérés côté colonne, nettoyés ensuite)
const DOTTED_TAG_SRC = '\\{([A-Za-z0-9_]+)\\.([A-Za-z0-9_À-ÿ]+)\\}';

// Cache des fetchs (métadonnées + tables enfants), le temps d'un lot de
// génération. Vidé à chaque événement Grist (voir main.js) pour ne pas servir
// de données périmées. Évite de re-télécharger la table enfant à chaque ligne
// lors d'un export ZIP.
const relationsCache = new Map();
const selectedViewCacheKey = Symbol('selectedView');

// Avertissements de la dernière construction de données (tables citées dans le
// modèle mais sans lien vers la table du widget). Affichés dans le statut.
let relationsWarnings = [];

function getRelationsWarnings() {
    return relationsWarnings;
}

function clearRelationsCache() {
    relationsCache.clear();
}

function fetchTableCached(tableId) {
    return fetchRelationCached(tableId, () => grist.docApi.fetchTable(tableId));
}

function fetchRelationCached(key, fetcher) {
    if (!relationsCache.has(key)) {
        const pending = fetcher().catch(error => {
            // Ne pas conserver un échec réseau, ni effacer une requête plus
            // récente démarrée après clearRelationsCache().
            if (relationsCache.get(key) === pending) relationsCache.delete(key);
            throw error;
        });
        relationsCache.set(key, pending);
    }
    return relationsCache.get(key);
}

// Grist applique la liaison (référence ou regroupement), les filtres et le tri.
// Ne pas remplacer cette lecture par fetchTable : cela inclurait les lignes exclues.
async function fetchSelectedViewRows() {
    const tbl = await fetchRelationCached(selectedViewCacheKey, () => grist.fetchSelectedTable({
        format: 'columns', includeColumns: 'normal', keepEncoded: false,
    }));
    return tbl.id.flatMap((id, i) => id === 'new' ? [] : [relationRowData(tbl, i)]);
}

// Quelles tables enfants sont réellement citées dans le modèle ?
// Fetch paresseux : aucune balise pointée => aucun fetch, comportement de base.
let templateScanCache = { buffer: null, tables: [] };
function getReferencedTables(buffer) {
    if (templateScanCache.buffer === buffer) {
        return templateScanCache.tables;
    }
    const zip = new PizZip(buffer);
    const found = new Set();
    for (const file of docxTextFiles(zip)) {
        // même réparation que la génération (balises éclatées par Word, correcteur...)
        const xml = repairDocxXml(file.asText());

        let m;
        const re = new RegExp(DOTTED_TAG_SRC, 'g');
        while ((m = re.exec(xml)) !== null) {
            found.add(m[1]);
        }
        // boucles explicites {#Table}...{/Table} écrites à la main dans le modèle
        const loopRe = /\{[#^]([A-Za-z0-9_]+)\}/g;
        while ((m = loopRe.exec(xml)) !== null) {
            found.add(m[1]);
        }
    }
    const tables = [...found];
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

        lignes.push(relationRowData(tbl, i));
    });
    return lignes;
}

function relationRowData(tbl, index) {
    const row = {};
    for (const col in tbl) {
        if (col === 'id' || col.startsWith('manualSort') || col.startsWith('gristHelper_')) continue;
        row[sanitizeKey(col)] = tbl[col][index];
    }
    return row;
}

// Lire group dans la table, même si cette colonne est masquée dans le widget.
// Les identifiants fournis par Grist sont la seule source d'appartenance au groupe.
async function fetchSummaryRows(summaryTable, sourceTable, parentId) {
    const summary = await fetchTableCached(summaryTable);
    const index = summary.id.indexOf(parentId);
    const group = index < 0 ? undefined : summary.group?.[index];
    if (!Array.isArray(group) || group[0] !== 'L') {
        throw new Error(`Le groupe sélectionné dans "${summaryTable}" est indisponible. Sélectionnez à nouveau une ligne.`);
    }
    const ids = group.slice(1);
    if (!ids.every(id => Number.isInteger(id) && id > 0)) {
        throw new Error(`La colonne group de "${summaryTable}" contient des références invalides.`);
    }
    if (!ids.length) return [];
    const source = await fetchTableCached(sourceTable);
    const indexes = new Map(source.id.map((id, i) => [id, i]));
    return ids.map(id => {
        if (!indexes.has(id)) throw new Error(`La ligne ${id} de "${sourceTable}" est introuvable ou inaccessible.`);
        return relationRowData(source, indexes.get(id));
    });
}

// Complète les données du parent avec les colonnes ABSENTES du record reçu
// par grist.onRecord : Grist n'envoie au widget que ses "colonnes visibles",
// donc une colonne ajoutée à la table APRÈS la création du widget est masquée
// et manquerait au publipostage. On relit la ligne complète via fetchTable.
// Les valeurs déjà présentes (décodées par Grist) restent prioritaires.
async function completeParentData(data, rowId) {
    try {
        const tableId = await grist.getSelectedTableId();
        const tbl = await fetchTableCached(tableId);
        const idx = tbl.id.indexOf(rowId);
        if (idx === -1) {
            return;
        }
        for (const col in tbl) {
            if (col === 'id' || col.startsWith('manualSort') || col.startsWith('gristHelper_')) {
                continue;
            }
            const cle = sanitizeKey(col);
            if (!(cle in data)) {
                data[cle] = tbl[col][idx];
            }
        }
    } catch (e) {
        console.warn("Impossible de compléter les colonnes masquées du widget", e);
    }
}

// Complète les données du parent avec les tables enfants citées dans le modèle :
// data devient { champs_du_parent..., Membres: [ {..}, {..} ], ... }
async function addChildTablesData(data, parentId, buffer) {
    relationsWarnings = [];
    const used = getReferencedTables(buffer);
    if (used.length === 0) {
        return data; // aucune balise {Table.Colonne} -> aucun fetch
    }

    const cible = await grist.getSelectedTableId(); // table du widget (le parent)
    // Une balise portant le nom de la table du widget désigne sa vue filtrée.
    if (used.includes(sanitizeKey(cible))) data[sanitizeKey(cible)] = await fetchSelectedViewRows();
    const related = used.filter(table => table !== sanitizeKey(cible));
    if (!related.length) return data;
    const refs = await findTablesReferencing(cible);
    const tables = await fetchTableCached('_grist_Tables');
    const sourceRef = tables.summarySourceTable?.[tables.tableId.indexOf(cible)];
    const summarySource = sourceRef ? tables.tableId[tables.id.indexOf(sourceRef)] : null;

    for (const table of related) {
        if (table === sanitizeKey(summarySource)) {
            data[table] = await fetchSummaryRows(cible, summarySource, parentId);
            continue;
        }
        const links = refs.filter(r => sanitizeKey(r.table) === table);
        const link = links[0];
        if (!link) {
            // Pas une table enfant liée : on ne touche pas aux données. Un
            // {#Champ} peut être une section conditionnelle sur un champ parent.
            if (!Object.keys(data).some(key => sanitizeKey(key) === table)) {
                const msg = `"${table}" ne référence pas la table "${cible}" (aucune colonne Ref:/RefList:, ou nom mal orthographié).`;
                console.warn(msg);
                relationsWarnings.push(msg);
            }
            continue;
        }
        if (links.length > 1) {
            relationsWarnings.push(`"${table}" possède plusieurs liens vers "${cible}" : ${links.map(r => r.column).join(', ')}. Le lien "${link.column}" est utilisé.`);
        }
        try {
            data[table] = await fetchChildRows(link.table, link.column, link.isList, parentId);
        } catch (error) {
            throw new Error(`Impossible de lire la table liée "${table}" : ${error.message}`);
        }
    }
    return data;
}

// Tables enfants citées via {Table.Col} dans un fragment XML
function dottedTablesIn(fragment) {
    const found = new Set();
    const tagRe = new RegExp(DOTTED_TAG_SRC, 'g');
    let m;
    while ((m = tagRe.exec(fragment)) !== null) {
        found.add(m[1]);
    }
    return [...found];
}

// Traduction de syntaxe au moment de générer :
// - {Membres.Col} dans une ligne de tableau -> {#Membres}{Col}...{/Membres}
//   enroulant toute la ligne, pour que docxtemplater répète la ligne.
// - {Membres.Col} ailleurs (hors tableau, cellule unique, tableau imbriqué) ->
//   le paragraphe entier est répété (marqueurs + paragraphLoop).
// SEULES les tables réellement résolues dans `data` sont transformées : une
// balise pointée non résolue reste visible dans le document au lieu de faire
// disparaître silencieusement la ligne/le paragraphe (et les balises simples
// comme {EEE} qui s'y trouvent).
// ⚠ À appeler AVANT la sanitisation des clés (sinon le "." devient "_").
function transformDottedLoops(zip, data) {
    for (const file of docxTextFiles(zip)) {
        zip.file(file.name, transformDottedXml(repairDocxXml(file.asText()), data));
    }
}

function transformDottedXml(xml, data) {

    const estResolue = (t) => data && Array.isArray(data[t]);
    const enBoucle = (fragment, table) => fragment.replace(
        new RegExp(`\\{${table}\\.([A-Za-z0-9_À-ÿ]+)\\}`, 'g'),
        '{$1}'
    );

    // 1) Lignes de tableau LES PLUS INTERNES (le garde-fou (?!<\/?w:tr[ >])
    //    empêche d'enjamber un tableau imbriqué) d'au moins 2 cellules : la
    //    ligne est répétée. Avec 1 seule cellule, docxtemplater répéterait le
    //    contenu bout à bout au lieu de la ligne -> passe 2.
    xml = xml.replace(/<w:tr[ >](?:(?!<\/?w:tr[ >])[\s\S])*?<\/w:tr>/g, (row) => {
        const tablesInRow = dottedTablesIn(row).filter(estResolue);
        if (tablesInRow.length === 0) {
            return row; // ligne normale, on ne touche pas
        }
        const cellules = (row.match(/<w:tc[ >]/g) || []).length;
        if (cellules < 2) {
            return row; // cellule unique : répétition de paragraphe (passe 2)
        }
        if (tablesInRow.length > 1) {
            console.warn("Plusieurs tables enfants dans la même ligne, seule la première est répétée :", tablesInRow);
        }
        const table = tablesInRow[0];

        // {Table.Col} -> {Col} puis enrouler la ligne : {#table} dans le 1er
        // nœud texte, {/table} dans le dernier.
        // (?:\s[^>]*)? cible uniquement <w:t>, pas <w:tr>/<w:tc>/<w:tbl>...
        return enBoucle(row, table)
            .replace(/(<w:t(?:\s[^>]*)?>)/, `$1{#${table}}`)
            .replace(/(<\/w:t>)(?![\s\S]*<\/w:t>)/, `{/${table}}$1`);
    });

    // 2) Balises pointées restantes : le paragraphe est répété pour chaque
    //    enfant (valable hors tableau comme dans une cellule). Le garde-fou
    //    (?!<w:p[ />]) évite d'enjamber un autre paragraphe, et l'ouverture
    //    <w:p(?:\s...)?> exclut les paragraphes vides auto-fermés <w:p/>.
    xml = xml.replace(/<w:p(?:\s[^>]*)?>(?:(?!<w:p[ />])[\s\S])*?<\/w:p>/g, (para) => {
        const tablesInPara = dottedTablesIn(para).filter(estResolue);
        if (tablesInPara.length === 0) {
            return para;
        }
        if (tablesInPara.length > 1) {
            console.warn("Plusieurs tables enfants dans le même paragraphe, seule la première est répétée :", tablesInPara);
        }
        const table = tablesInPara[0];

        // paragraphes marqueurs : avec paragraphLoop, docxtemplater les retire
        // et répète le paragraphe central pour chaque enregistrement enfant
        return `<w:p><w:r><w:t>{#${table}}</w:t></w:r></w:p>`
            + enBoucle(para, table)
            + `<w:p><w:r><w:t>{/${table}}</w:t></w:r></w:p>`;
    });

    return xml;
}
