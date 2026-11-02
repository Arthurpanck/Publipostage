// Extrait réel d'un document Grist (métadonnées et données), relu par
// l'API SQL : colonnes typées, couleurs des choix, références.
// fetchTable renvoie des colonnes ; listes encodées ["L", …].
const W = o => JSON.stringify(o);
const COLS = [
    [5, 2, 'manualSort', 'ManualSortPos', 'manualSort', '', 0],
    [6, 2, 'Nom', 'Text', 'Nom', '', 0],
    [7, 2, 'Pole', 'Choice', 'Pôle', W({ choices: ['RH', 'Technique', 'Finance'], choiceOptions: { Finance: { fillColor: '#3B6FD8', textColor: '#FFFFFF' }, RH: { fillColor: '#8E5BD9', textColor: '#FFFFFF' }, Technique: { fillColor: '#E8A33D', textColor: '#000000' } } }), 0],
    [8, 3, 'manualSort', 'ManualSortPos', 'manualSort', '', 0],
    [9, 3, 'Nom', 'Text', 'Nom', '', 0],
    [10, 3, 'Fonction', 'Text', 'Fonction', '', 0],
    [11, 4, 'manualSort', 'ManualSortPos', 'manualSort', '', 0],
    [12, 4, 'Nom', 'Text', 'Nom', '', 0],
    [13, 4, 'Statut', 'Choice', 'Statut', W({ choices: ['CDI', 'CDD', "Période d'essai"], choiceOptions: { CDD: { fillColor: '#F2B13C', textColor: '#3D2A00' }, CDI: { fillColor: '#2E9E5B', textColor: '#FFFFFF' }, "Période d'essai": { fillColor: '#3B6FD8', textColor: '#FFFFFF' } } }), 0],
    [14, 4, 'Competences', 'ChoiceList', 'Compétences', W({ choices: ['Sourcing', 'Entretiens', 'LinkedIn Recruiter', 'Anglais'], choiceOptions: { Anglais: { fillColor: '#F2B13C', textColor: '#3D2A00' }, Entretiens: { fillColor: '#2E9E5B', textColor: '#FFFFFF' }, 'LinkedIn Recruiter': { fillColor: '#3B6FD8', textColor: '#FFFFFF' }, Sourcing: { fillColor: '#8E5BD9', textColor: '#FFFFFF' } } }), 0],
    [15, 4, 'Etape', 'Choice', 'Étape', W({ choices: ['Offre', 'Signature', 'Accueil', 'Formation', 'Autonome'] }), 0],
    [16, 4, 'Mutuelle', 'Bool', 'Mutuelle', '', 0],
    [17, 4, 'Teletravail', 'Bool', 'Télétravail', '', 0],
    [18, 4, 'Salaire', 'Numeric', 'Salaire', W({ numMode: 'currency', currency: 'EUR', decimals: 2, maxDecimals: 2 }), 0],
    [19, 4, 'Objectif', 'Numeric', 'Objectif', W({ numMode: 'currency', currency: 'EUR', decimals: 2, maxDecimals: 2 }), 0],
    [20, 4, 'Taux', 'Numeric', 'Taux', W({ numMode: 'percent' }), 0],
    [21, 4, 'Conges_pris', 'Int', 'Congés pris', '', 0],
    [22, 4, 'Conges_acquis', 'Int', 'Congés acquis', '', 0],
    [23, 4, 'Equipe', 'Ref:Equipes', 'Équipe', '', 6],
    [24, 4, 'Pole_equipe', 'Ref:Equipes', 'Pôle', '', 7],
    [25, 4, 'Collegues', 'RefList:Personnes', 'Collègues', '', 9],
];
const columns = rows => Object.fromEntries(Object.keys(rows[0]).map(k => [k, rows.map(r => r[k])]));
const TABLES = {
    _grist_Tables: { id: [2, 3, 4], tableId: ['Equipes', 'Personnes', 'Salaries'] },
    _grist_Tables_column: {
        id: COLS.map(c => c[0]), parentId: COLS.map(c => c[1]), colId: COLS.map(c => c[2]), type: COLS.map(c => c[3]),
        label: COLS.map(c => c[4]), widgetOptions: COLS.map(c => c[5]), visibleCol: COLS.map(c => c[6]),
    },
    _grist_DocInfo: { id: [1], documentSettings: [W({ locale: 'fr-FR', engine: 'python3' })] },
    Equipes: columns([
        { id: 1, Nom: 'Talents', Pole: 'RH' }, { id: 2, Nom: 'Maintenance', Pole: 'Technique' }, { id: 3, Nom: 'Comptabilité', Pole: 'Finance' }]),
    Personnes: columns([
        { id: 1, Nom: 'Julien Petit', Fonction: 'Gestionnaire de paie' }, { id: 2, Nom: 'Inès Garnier', Fonction: 'Juriste social' }, { id: 3, Nom: 'Hugo Lemaire', Fonction: 'Technicien' }]),
};
// Lignes telles que reçues par le widget (onRecords : valeurs décodées)
const RECORDS = [
    { id: 1, Nom: 'Camille Durand', Statut: 'CDI', Competences: ['Sourcing', 'Entretiens', 'Anglais'], Etape: 'Accueil', Mutuelle: true, Teletravail: false, Salaire: 2450, Objectif: 4000, Taux: 0.6125, Conges_pris: 18, Conges_acquis: 25, Equipe: 1, Pole_equipe: 1, Collegues: [1, 2] },
    { id: 2, Nom: 'Karim Benali', Statut: 'CDD', Competences: ['LinkedIn Recruiter'], Etape: 'Offre', Mutuelle: false, Teletravail: true, Salaire: 2100, Objectif: 2000, Taux: 1.05, Conges_pris: 5, Conges_acquis: 25, Equipe: 2, Pole_equipe: 2, Collegues: [3] },
    { id: 3, Nom: 'Sophie de la Tour', Statut: "Période d'essai", Competences: null, Etape: 'Autonome', Mutuelle: true, Teletravail: true, Salaire: 3100, Objectif: 4000, Taux: 0.775, Conges_pris: 0, Conges_acquis: 4, Equipe: 3, Pole_equipe: 3, Collegues: null },
];
const encode = v => (Array.isArray(v) ? ['L', ...v] : v);
TABLES.Salaries = Object.fromEntries(Object.keys(RECORDS[0]).map(k => [k, RECORDS.map(r => encode(r[k]))]));

function gristMock(fetched = []) {
    return {
        getSelectedTableId: async () => 'Salaries',
        docApi: { fetchTable: async (t) => { fetched.push(t); if (!TABLES[t]) throw new Error('table inconnue ' + t); return TABLES[t]; } },
    };
}
module.exports = { TABLES, RECORDS, gristMock };
