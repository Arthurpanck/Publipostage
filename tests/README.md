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
