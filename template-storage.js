/** Les options du widget ne comptent pas comme références de pièces jointes.
 * Conserver les modèles dans une vraie colonne Attachments évite leur nettoyage
 * et permet à Grist d'appliquer les droits de lecture habituels des cellules.
 */
const TEMPLATE_TABLE = 'Publipostage_Modeles';

async function retainTemplateAttachment(attachmentId, name) {
    const tables = await grist.docApi.fetchTable('_grist_Tables');
    let tableIndex = tables.tableId.indexOf(TEMPLATE_TABLE);
    if (tableIndex === -1) {
        try {
            const result = await grist.docApi.applyUserActions([['AddTable', TEMPLATE_TABLE, [
                { id: 'Nom', type: 'Text', isFormula: false },
                { id: 'Fichier', type: 'Attachments', isFormula: false },
            ]]]);
            if (result.retValues[0].table_id !== TEMPLATE_TABLE) {
                throw new Error('La table de modèles a été créée sous un autre nom. Réessayez.');
            }
        } catch (error) {
            // Une autre session peut avoir créé la table entre-temps.
            const refreshed = await grist.docApi.fetchTable('_grist_Tables');
            if (!refreshed.tableId.includes(TEMPLATE_TABLE)) throw error;
        }
    }
    const schema = await grist.docApi.fetchTable('_grist_Tables');
    const tableRef = schema.id[schema.tableId.indexOf(TEMPLATE_TABLE)];
    const columns = await grist.docApi.fetchTable('_grist_Tables_column');
    for (const [colId, type] of [['Nom', 'Text'], ['Fichier', 'Attachments']]) {
        const i = columns.id.findIndex((_, i) => columns.parentId[i] === tableRef && columns.colId[i] === colId);
        if (i === -1 || columns.type[i] !== type || columns.isFormula[i]) {
            throw new Error(`La table ${TEMPLATE_TABLE} doit contenir une colonne ${colId} de type ${type}, non calculée.`);
        }
    }
    const existing = await grist.docApi.fetchTable(TEMPLATE_TABLE);
    if (!existing.Fichier.some(value => Array.isArray(value) && value.slice(1).includes(attachmentId))) {
        await grist.docApi.applyUserActions([['AddRecord', TEMPLATE_TABLE, null,
            { Nom: name, Fichier: ['L', attachmentId] }]]);
    }
    // Ne pas annoncer la conservation tant que la référence n'est pas relue.
    const saved = await grist.docApi.fetchTable(TEMPLATE_TABLE);
    if (!saved.Fichier.some(value => Array.isArray(value) && value.slice(1).includes(attachmentId))) {
        throw new Error('La pièce jointe n’a pas été conservée dans le document.');
    }
}
