# Le modèle reste-t-il attaché au widget ?

Scénarios de bout en bout contre un vrai Grist (Docker) : plusieurs comptes,
rechargements, règles d'accès, purge des pièces jointes, copies, export,
migration de l'ancien format. Ils ne tournent pas avec `npm test`.

## Lancer

```sh
docker run -d --name grist-test -p 8484:8484 \
  -e GRIST_IN_SERVICE=true -e GRIST_TEST_LOGIN=1 -e GRIST_SINGLE_ORG=docs \
  -e GRIST_DEFAULT_EMAIL=owner@test.fr -e GRIST_FORCE_LOGIN=true \
  -e GRIST_SANDBOX_FLAVOR=unsandboxed -e APP_HOME_URL=http://localhost:8484 \
  gristlabs/grist:latest

npm install --no-save playwright && npx playwright install chromium
node tests/grist-docker/run.cjs            # tous les scénarios
node tests/grist-docker/run.cjs Migration  # filtre sur le nom
```

Variables : `GRIST_URL` (défaut `http://localhost:8484`), `CHROMIUM_PATH`, et
`WIDGET_ROOT` pour tester une autre version du widget (dossier, par exemple un
`git worktree`).
Le script sert lui-même le widget du dépôt sur un port libre ; chaque
scénario crée son document, avec `owner@test.fr`, `editor@test.fr` (éditeur)
et `viewer@test.fr` (lecteur).

## Ce que Grist fait des pièces jointes (grist-core)

- Une pièce jointe qu'aucune cellule de type Pièce jointe ne référence est
  marquée supprimée au premier passage du nettoyage (toutes les heures et à la
  fermeture du document), puis effacée 7 jours plus tard.
- Un utilisateur soumis à des règles d'accès ne peut télécharger une pièce
  jointe que s'il voit une cellule qui la contient.
- `grist.setOptions` n'enregistre rien : il allume le bouton « Enregistrer »
  de Grist, et les options sont perdues au rechargement si personne ne clique.

- Si la case « Autoriser les éditeurs à modifier la structure » est décochée
  (règle `user.Access != OWNER` → `-S` sur toutes les tables), un éditeur ne
  peut modifier aucune option de widget, pas même avec le bouton
  « Enregistrer » de Grist : seul un propriétaire peut attacher un modèle.
  Le widget le dit à l'éditeur et reprend l'ancien format quand un
  propriétaire l'ouvre.

- Quand on re-sélectionne un widget (« Custom URL » → « Add widget », par
  exemple pour changer son adresse), Grist remet ses options à zéro et son
  accès à « aucun ». Un widget ajouté par l'interface démarre aussi sans
  accès, tant qu'on n'a pas cliqué sur « Accepter ».

- Grist transmet au widget ses nouvelles options (`onOptions`) un instant
  après les avoir enregistrées : une lecture des sections peut être en
  avance sur les options reçues.

- Grist envoie à chaque ouverture du document les options de tous les
  widgets, quelle que soit la page affichée. Un modèle rangé dans les options
  (version précédente, en base64) était donc téléchargé par tout le monde, à
  chaque ouverture, multiplié par le nombre de widgets.

D'où les choix du widget (`template-store.js`) : le modèle est une pièce jointe
rangée dans la table `Publipostage_Modeles_NE_PAS_SUPPRIMER` (données brutes, sans page), une
ligne par widget repérée par le numéro de sa section ; rangée dans une cellule,
elle n'est pas purgée. Les options du widget ne gardent que le nom, l'empreinte
et le numéro du widget : leur changement prévient les autres utilisateurs, et
l'envoi (fichier + options) est une seule action, annulable par Ctrl+Z.
Écrire dans la table ne demande que le droit de modifier les données : quand
la structure est réservée aux propriétaires, la table est préparée à
l'ouverture du widget par un propriétaire, qui inscrit aussi dans les options
de chaque widget son numéro (sans lui, un éditeur ne peut pas départager des
widgets neufs d'une même table) ; les éditeurs y enregistrent ensuite leurs
modèles. La purge `removeUnused?expiredonly=0` simule les 7 jours.

Un widget se reconnaît parmi ceux de sa table à ses options, qui portent son
numéro. À l'ouverture, tout utilisateur qui peut écrire les options remet
d'aplomb celles de tous les widgets Publipostage du document, d'après la
table (`healSections`) : options remises à zéro par Grist, widgets neufs aux
options identiques, widget ou page dupliqués. Avant ce correctif, deux
widgets d'une même table dont les options avaient été remises à zéro
perdaient leur modèle à chaque rechargement, pour tout le monde, alors que
les fichiers étaient toujours dans la table.

## Résultats

Première version (modèle en pièce jointe non rangée + `grist.setOptions`) :

| Scénario | Résultat |
|---|---|
| Envoi sans cliquer « Enregistrer », puis rechargement ou autre utilisateur | perdu (le widget affichait pourtant « Modèle sauvegardé ») |
| Envoi + « Enregistrer », puis rechargement, éditeur, lecteur | ok |
| Règle d'accès sur une colonne, ouverture par un lecteur | perdu (téléchargement refusé) |
| Purge des pièces jointes inutilisées (7 jours) | perdu |
| Copie du document comme modèle | perdu |
| Deux widgets sur la même page, « Enregistrer » cliqué une fois | le second perdu |

Version précédente (modèle en base64 dans les options) : scénarios ok, sauf
pour un éditeur quand la structure est réservée aux propriétaires, mais
ouverture du document très lente (ci-dessous).

Version actuelle : tous les scénarios de `run.cjs` passent (29), dont un
éditeur qui charge un modèle dans un document à structure réservée, une fois
le widget ouvert par le propriétaire. Changement de comportement : une copie
« comme modèle » (sans données) perd les modèles, le widget dit lequel
recharger. Cinq widgets ont chacun leur modèle (trois sur la même page), et
le widget ne limite plus la taille du modèle (20 Mo testés) : seule compte
la limite des pièces jointes du serveur Grist. Trois widgets neufs d'une même
page, ouverts par le propriétaire d'un document à structure réservée, ne
créent qu'une table, et un éditeur enregistre ensuite un modèle dans chacun.

Vérifié aussi à la main avec Playwright, widgets ajoutés par l'interface de
Grist : redémarrage du serveur (pages fermées, puis page restée ouverte),
document laissé ouvert 35 minutes par trois utilisateurs, propriétaire et
éditeur qui ouvrent et envoient en même temps, page dupliquée par le menu de
Grist (chaque copie garde son modèle quand celui de l'original change),
widgets re-sélectionnés par la galerie.

Vitesse, débit limité à 20 Mbit/s (Chromium), 10 widgets avec un modèle de
2 Mo chacun :

| | Version précédente (base64 dans les options) | Version actuelle |
|---|---|---|
| Données reçues à l'ouverture du document | 26,7 Mo | ≈ 0 |
| Page sans widget affichée | 32 s | 0,7 s |
| Page du widget, aperçu affiché | 33,5 s | 3,1 s |
| Envoi d'un modèle de 2 Mo | plus de 20 s | 1,5 s |
