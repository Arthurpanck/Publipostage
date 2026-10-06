# Modèles d'exemple

Modèles du document Grist de démonstration « Publipostage — guide des
fonctionnalités » : une page par fonctionnalité, un modèle par page. Ils
utilisent les tables `Evenements` (Titre, Type, Lieux, Debut, Fin,
Participants, Responsable, Confirme, et les colonnes formules Debut_texte,
Fin_texte, Lieux_texte) et `Lieux` (Nom, Jauge).

| Modèle | Page | Widget posé sur | Ce qu'il montre |
|---|---|---|---|
| `1-publipostage-simple.docx` | 1 · Publipostage simple | Evenements | Balises de colonnes, majuscules, accents, doubles accolades, texte conditionnel, en-tête et pied de page |
| `2-tables-liees.docx` | 2 · Tables liées | Lieux | Lignes d'une table qui référence la table du widget : ligne de tableau et puce répétées, message si aucune ligne |
| `3-reference-portee.docx` | 3 · Référence portée par la ligne | Evenements | Lignes désignées par une colonne Référence de la ligne du widget |
| `4-vue-filtree.docx` | 4 · Vue filtrée | Evenements, lié au tableau des lieux | Lignes transmises au widget (liaison, filtres, tri) et ligne sélectionnée |
| `5-regroupement.docx` | 5 · Regroupement | Evenements regroupée par Type | Colonnes du groupe et lignes du groupe |
| `6-filtres.docx` | 6 · Filtres | Evenements | Filtre dans un bandeau de tableau, dans un titre avec liste à puces, `!=`, filtre sans résultat |
| `7-formulaire.pdf` | 7 · PDF à formulaire | Evenements | Champs de formulaire nommés d'après les colonnes, case à cocher |

Pour essayer : télécharger le modèle, puis le déposer dans le widget
Publipostage de la page correspondante (« Charger un modèle » ou roue crantée
puis « Changer de modèle »).
