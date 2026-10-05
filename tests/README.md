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
n'est ni uploadé ni sauvegardé à la place du modèle précédent. Ils ne remplacent
pas un test de persistance avec le vrai service Grist.

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
- Réduire le widget à environ 350–460 px avec un nom de fichier long et une
  balise inconnue. La roue des options doit rester accessible sans défilement
  horizontal ; son menu et le diagnostic des balises doivent s'ouvrir.

Les filtres Grist peuvent exclure une ligne sélectionnée dans la table source.
Vérifier leur configuration avant de conclure à une erreur de sélection.

## Relations

Les tests simulent les réponses Grist (Ref, RefList, métadonnées), puis passent
les données dans le vrai moteur DOCX. Ils couvrent le filtrage des enfants,
les parents sans enfant, les boucles manuelles, le cache par lot, les erreurs
réseau et les liens ambigus. Ils ne valident pas les permissions ni les
événements du service Grist réel.

Pour un test manuel, utiliser une table Enfants avec une colonne Parent
de type Référence vers la table du widget et un modèle `{Enfants.Nom}`.
Vérifier deux parents avec des enfants distincts, puis un parent sans enfant.
Recommencer avec Parent de type Liste de références et un enfant partagé.
Comparer l’aperçu, le document individuel et les documents du ZIP.

## Regroupements Grist

`tests/summary.test.cjs` simule la table récapitulative, son
`summarySourceTable` et les références encodées de `group`, puis génère de
vrais DOCX avec les bibliothèques embarquées. Les tests couvrent plusieurs
groupes, une ligne partagée, le groupe vide, les tableaux, les en-têtes, les
boucles manuelles, l’actualisation du cache et les références indisponibles.

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

`tests/filtered-view.test.cjs` vérifie l’ordre, la restriction aux lignes
transmises, les boucles, le cache, la vue vide et les erreurs de lecture.
Pour vérifier dans Grist : brancher le widget sur la table détaillée,
configurer sa liaison au sélecteur de commune ou au regroupement, puis
changer le sélecteur sans cliquer sur chaque service. Comparer la liste
publipostée aux lignes transmises au widget, puis tester un filtre et un tri.

## Normalisation

Les noms de balises et les clés Grist sont normalisés par `merge-utils.js` :
minuscules, accents supprimés, caractères spéciaux remplacés par des underscores.
Les suffixes comme `_2` restent distincts. Les identifiants réels des tables
sont conservés pour les appels Grist. Les boucles et conditions Word suivent
la même règle, ainsi que les champs PDF. Les valeurs ne changent pas de casse.
