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
n'est ni uploadé ni sauvegardé à la place du modèle précédent, que les lignes
sont demandées avec leurs colonnes masquées (`includeColumns: 'normal'`) et
qu'un aperçu périmé n'écrase pas le plus récent. Ils ne remplacent pas un test
de persistance avec le vrai service Grist.

Chaque génération (aperçu, document, ZIP) est un lot (`newBatch` dans
`main.js`) : les tables lues, les lignes du widget et les alertes lui
appartiennent. Les tests passent un lot explicite (`batch` dans `helpers.cjs`).

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

Ces lignes sont celles reçues par `onRecords`, sans nouvelle lecture.
`tests/filtered-view.test.cjs` vérifie l’ordre, la restriction aux lignes
transmises, les boucles, la vue vide et l’absence de lecture de la table.
Pour vérifier dans Grist : brancher le widget sur la table détaillée,
configurer sa liaison au sélecteur de commune ou au regroupement, puis
changer le sélecteur sans cliquer sur chaque service. Comparer la liste
publipostée aux lignes transmises au widget, puis tester un filtre et un tri.
