# Tests de publipostage

Avec Node.js 18 ou plus :

```sh
npm ci
npm test
```

Les tests exécutent les bibliothèques PizZip et Docxtemplater déjà présentes
sous `inc/`, sur de petits DOCX créés en mémoire. `@xmldom/xmldom` fournit les
API XML du navigateur uniquement aux tests Node. L'application reste statique,
sans étape de build ni nouvelle dépendance en production.

Les tests du contrôleur simulent Grist et vérifient qu'un fichier invalide
n'est ni uploadé ni sauvegardé à la place du modèle précédent, et qu'un aperçu
périmé n'écrase pas le plus récent. Ils ne remplacent pas un test de
persistance avec le vrai service Grist.

Chaque génération (aperçu, document, ZIP) est un lot (`newBatch` dans
`main.js`) : les tables lues, les lignes du widget et les alertes lui
appartiennent. Les tests passent un lot explicite (`batch` dans `helpers.cjs`).

Selon la version de Grist ou les colonnes associées du widget, la ligne reçue
peut ne pas contenir toutes ses colonnes, malgré `includeColumns: 'normal'`.
`completeRow` (`relations-tools.js`) relit alors la table du widget une fois par
lot ; les valeurs reçues restent prioritaires. `tests/columns.test.cjs` rejoue ce
cas avec le vrai `main.js` et un extrait du modèle de test Donnees/Lieux.

## Vérification manuelle dans Grist

- Charger un DOCX contenant `{Titre}`, `{{Titre}}` et `{{Participants}}`.
  Vérifier l'aperçu, le changement de ligne, l'export individuel et le ZIP.
- Couper une balise sur plusieurs fragments Word en mettant une partie en gras.
  Vérifier que le texte après la balise et un paragraphe sans balise conservent
  leur mise en forme.
- Mettre une balise accentuée dans le corps, l'en-tête et le pied de page.
  Les trois emplacements doivent utiliser la même colonne.
- Charger un DOCX contenant `{{Titre}` ou un faux fichier `.docx` : le modèle
  précédent doit rester utilisable, y compris après rechargement de Grist.
- Charger un modèle avec une extension `.DOCX` en majuscules.
- Masquer dans le widget une colonne utilisée par le modèle : elle doit rester
  publipostée dans l'aperçu, le document et le ZIP.
- Placer une balise dans un paragraphe qui ancre une zone de texte : elle doit
  être remplie comme les autres.
- Réduire le widget de 1000 à 300 px avec un nom de fichier long et une balise
  inconnue : le titre disparaît sous 900 px, les libellés courts apparaissent
  sous 720 px, puis les actions passent dans la roue sous 460 px. La roue doit
  rester accessible sans défilement horizontal ; son menu et le diagnostic des
  balises doivent s'ouvrir.

Les filtres Grist peuvent exclure une ligne sélectionnée dans la table source.
Vérifier leur configuration avant de conclure à une erreur de sélection.

## Relations

Les tests simulent les réponses Grist (Ref, RefList, métadonnées), puis passent
les données dans le vrai moteur DOCX. Ils couvrent le filtrage des enfants,
les parents sans enfant, les boucles manuelles, le cache par lot, les erreurs
réseau, les liens ambigus, les références portées par la table du widget
(Ref et RefList, dans leur ordre) et les références à double sens, où le lien
de la table citée est prioritaire. Ils ne valident pas les permissions ni les
événements du service Grist réel.

Pour un test manuel, utiliser une table Enfants avec une colonne Parent
de type Référence vers la table du widget et un modèle `{Enfants.Nom}`.
Vérifier deux parents avec des enfants distincts, puis un parent sans enfant.
Recommencer avec Parent de type Liste de références et un enfant partagé.
Recommencer avec la référence dans l’autre sens : une colonne Lieux de type
Liste de références dans la table du widget et un modèle `{Lieux.Nom}`.
Comparer l’aperçu, le document individuel et les documents du ZIP.

## Regroupements Grist

Une table regroupée n’est pas un cas particulier : sa colonne `group` est une
Liste de références vers la table source, portée par la ligne du widget.
`tests/summary.test.cjs` simule ses métadonnées et la valeur décodée de
`group` reçue par le widget, puis génère de vrais DOCX avec les bibliothèques
embarquées. Les tests couvrent plusieurs groupes, une ligne partagée, le
groupe vide, les tableaux, les en-têtes, les boucles manuelles, l’actualisation
par un nouveau lot et les références indisponibles.

Dans Grist, sélectionner comme source du widget une table regroupée, puis
utiliser `{Donnees.Titre}` dans le modèle si la table d’origine est Donnees.
Le nom à utiliser est l’identifiant de la table source, pas celui de la table
récapitulative. Les colonnes du groupe restent disponibles sous leur nom
simple, par exemple `{count}`. La colonne `group` peut rester masquée.

Pour la validation manuelle : sélectionner deux groupes successifs et
comparer les lignes produites à celles du groupe dans Grist. Modifier le
titre d’une ligne source sans changer les totaux puis relancer un export ;
le nouveau titre doit apparaître. Vérifier aussi le ZIP, avec un document
par ligne du regroupement transmise au widget. Les tests simulés ne
remplacent pas ces vérifications sur le service Grist réel.

## Vue du widget filtrée par une liaison

Si le widget utilise BDD_SERVICES, la balise
`{BDD_SERVICES.Service_commune_Prenom_Nom}` répète les lignes transmises
à cette vue. Grist applique lui-même la liaison (référence ou regroupement),
les filtres et le tri. Aucun filtre d’une autre vue indépendante n’est copié.
La balise simple `{Service_commune_Prenom_Nom}` reste celle de la ligne
sélectionnée. Le ZIP garde son fonctionnement par ligne : un modèle qui
contient toute la vue y répétera cette liste dans chaque document.

Ces lignes sont celles reçues par `onRecords`. Si Grist ne transmet pas
toutes leurs colonnes, elles sont complétées par une lecture de la table, une
fois par lot, sans réintroduire les lignes exclues par le widget.
`tests/filtered-view.test.cjs` vérifie l’ordre, la restriction aux lignes
transmises, les boucles, la vue vide et l’absence de lecture de la table.
Pour vérifier dans Grist : brancher le widget sur la table détaillée,
configurer sa liaison au sélecteur de commune ou au regroupement, puis
changer le sélecteur sans cliquer sur chaque service. Comparer la liste
publipostée aux lignes transmises au widget, puis tester un filtre et un tri.

## Normalisation

Les noms de balises et les clés Grist sont normalisés par `merge-utils.js`,
une seule fois : les balises Word pendant la réparation du modèle
(`normalizeDocxTag`), les données à l'entrée des générateurs
(`normalizeMergeData`) :
minuscules, accents supprimés, caractères spéciaux remplacés par des underscores.
Les suffixes comme `_2` restent distincts. Les identifiants réels des tables
sont conservés pour les appels Grist. Les boucles et conditions Word suivent
la même règle, ainsi que les champs PDF. Les valeurs ne changent pas de casse.

## Filtres

`{filtre: Table.Colonne == "valeur"}` (ou `!=`) ne garde que les lignes
correspondantes de Table dans sa portée : le tableau Word qui contient la
balise, sinon jusqu’à `{fin filtre}` (ou `{/filtre}`) ou la fin du document.
Un second filtre sur la même table remplace le premier. Les lignes
`{Table.Colonne}`, les boucles `{#Table}` et les conditions `{^Table}` de la
portée suivent le filtre. La balise affiche la valeur (rien pour `!=`) ; une
ligne réduite à `{fin filtre}` ou à un filtre `!=` disparaît, sauf dans une
cellule de tableau. La comparaison ignore casse, accents et ponctuation ; une
liste (choix multiples, références) correspond si l’un de ses éléments
correspond. Les guillemets droits, typographiques ou chevrons sont facultatifs.

`tests/filter.test.cjs` couvre le bandeau de tableau, la portée limitée au
tableau, les puces jusqu’à `{fin filtre}`, les filtres successifs, `!=`, les
boucles et conditions inversées, les valeurs encodées de Grist, la colonne ou
la table inconnue (signalée dans la pastille) et le filtre mal écrit (refusé
dès l’import).

Pour vérifier dans Grist : un tableau dont le bandeau contient
`{filtre: BDD_SERVICES.Fonction == "Direction territoriale"}` et une ligne
`{BDD_SERVICES.Nom}` ; puis une liste à puces encadrée par un filtre et
`{fin filtre}`, suivie d’une liste non filtrée. Comparer l’aperçu, le document
individuel et le ZIP.
