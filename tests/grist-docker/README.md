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

D'où les choix du widget (`template-store.js`) : le modèle est rangé dans les
options du widget, et le widget les écrit lui-même dans le document. La
purge `removeUnused?expiredonly=0` simule les 7 jours.

## Résultats

Ancienne version (modèle en pièce jointe + `grist.setOptions`) :

| Scénario | Résultat |
|---|---|
| Envoi sans cliquer « Enregistrer », puis rechargement ou autre utilisateur | perdu (le widget affichait pourtant « Modèle sauvegardé ») |
| Envoi + « Enregistrer », puis rechargement, éditeur, lecteur | ok |
| Règle d'accès sur une colonne, ouverture par un lecteur | perdu (téléchargement refusé) |
| Purge des pièces jointes inutilisées (7 jours) | perdu |
| Copie du document comme modèle | perdu |
| Deux widgets sur la même page, « Enregistrer » cliqué une fois | le second perdu |

Version actuelle : tous les scénarios de `run.cjs` passent.
