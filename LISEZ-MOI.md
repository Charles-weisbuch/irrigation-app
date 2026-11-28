# Tableaux récapitulatifs mensuels (à faire une seule fois)

Les tableaux « Mois | Pluie | Arrosage | Cumul | Besoins » (mai → avril + TOTAL période estivale) sont en haut de chaque onglet (La Fontaine, Bousquet et Riou, La baume, Pilles) et se calculent seuls à partir de TOUTES les saisies situées en dessous (formules ouvertes, sans limite de ligne).

1. **Fichier > Créer une copie** (sauvegarde de sécurité).
2. Colle le nouveau `Code.gs` en entier dans Apps Script, puis **Déployer > Gérer les déploiements > Nouvelle version**.
3. Lance `verifierTableauxRecap` (affiche ce qui sera fait, ne modifie rien).
4. Lance `mettreAJourTableauxRecap` **une seule fois** : il remonte le tableau en haut des onglets qui ne l'ont pas encore, ajoute les mois manquants, réécrit les formules et les couleurs (vert / jaune / rouge).
5. Remplace `index.html` et `sw.js` sur GitHub (Commit / Push), puis Ctrl+Maj+R.

Notes : les besoins d'octobre à avril sont laissés vides (à remplir si tu veux un suivi) ; janvier à avril sont comptés sur 2027 ; les mm d'arrosage utilisent la cellule « Surface couverte (ha) » de chaque ligne.

# Suivi Irrigation — mise en place (15 minutes)

## Module Récolte

Accueil → **Récolte** :

- **Pointage récolte** → parcelle (La Fontaine, La Baume, Bousquet, Riou, Les Pilles) → **appui de 2 s** sur un poids (<10 g, 20 g, 50 g, 100 g, >100 g, Brumale) = une truffe dans l'onglet **Récolte pointage** (n°, date, parcelle, poids, position GPS cliquable, commentaire). Le bouton se remplit pendant l'appui ; « ✓ n°X » n'apparaît que si la ligne est bien arrivée dans le Sheet.
- **Poids après récolte** → une seule page : date, **parcelle**, type de poids (**Brut non lavé** / **Lavé non trié** / **Net vendable**), poids en **grammes**, commentaire → onglet **Récolte volume**. Une ligne par date (ordre chronologique) ; le poids va dans la colonne de la parcelle et du type choisis ; plusieurs poids identiques le même jour se **cumulent**. Totaux par ligne (V brut, W lavé, X net, Y % de perte = 1 − net / lavé) et totaux / moyenne **en haut (ligne 1) sur toute la colonne** (`=SUM(V4:V)`…), nouvelles saisies comprises. Un total tapé à la main (ex. net du 05/11) n'est jamais écrasé tant qu'aucune parcelle n'a de net ce jour-là. Bousquet n'a pas de colonne « Commentaire » (la colonne F « poids brut » est ignorée) : le commentaire est alors mis en note sur la cellule.
- **Vente** (nouveau, sur la page d'accueil) → date + pour chaque qualité vendue (Brisure, Morceaux, 2e choix, 1er choix, Extra) prix €/kg et poids g → onglet **Vente**. Une ligne par date ; même qualité le même jour : poids additionnés, prix = moyenne pondérée. Colonnes L (total €), M (prix moyen), N (total g) par date ; tableau rouge du haut (ligne 3) = totaux / moyennes de toute la colonne ; c'est lui qu'affiche « Consulter les ventes ». **Client** : liste déroulante (noms dans la colonne « Liste clients » du Sheet, modifiable) + « Autre » pour saisir un nouveau nom, ajouté automatiquement à la liste ; une ligne par date ET par client ; colonnes « Client » et « Liste clients » créées automatiquement à droite de N.
- **Consulter les poids après récolte** : totaux brut / net vendable et perte, totaux par parcelle, détail par date (avec les commentaires).
- **Carte des récoltes** : fonds IGN (photos aériennes, plan, cadastre) ou OpenStreetMap ; les truffes pointées apparaissent d'elles-mêmes (rafraîchissement toutes les 15 s).

Mise en place : remplacer `Code.gs` → redéployer une nouvelle version → exécuter **verifierOngletVolume** et **verifierOngletVente** (affiche comment les en-têtes de « Récolte volume » sont lus, sans rien écrire) → remplacer `index.html`, `manifest.json`, `sw.js` et ajouter les icônes (`icon-192.png`, `icon-512.png`, `icon-maskable-512.png`, `apple-touch-icon.png`) → Commit/Push → Ctrl+Shift+R. Pour changer l'icône et le nom sur le téléphone : désinstaller l'appli puis la réinstaller depuis Chrome.

---

## Nouveau : module Heures — mise en place (à faire une seule fois)

L'appli s'appelle maintenant **SCEA Les Pilles** et propose deux pôles sur l'écran d'accueil : **Irrigation** (inchangé) et **Heures** (nouveau).

1. Remplace `Code.gs` par la dernière version (contient tout : irrigation + heures)
2. Menu déroulant de fonctions → sélectionne **configurerTableauxHeures**
3. Clique **Exécuter** — ça pose les formules dans les 3 tableaux récapitulatifs déjà présents dans l'onglet "Heures" (Total par tâche, Total par lieu, Total par personne/mois)
4. Vérifie dans le Journal d'exécution, puis dans le Sheet que les colonnes I à O (tâche), R à X (lieu) et Z à AM (personne) affichent bien des totaux (0 pour l'instant, tant qu'aucune heure n'est saisie)
5. Redéploie une nouvelle version, remplace `index.html` (URL déjà intégrée) → Commit → Push → Ctrl+Shift+R

**Fonctionnement de la saisie** : Nom (6 choix fixes : Paul, René, Charles, Externe 1/2/3), Date, Parcelle/Lieu (liste + "Autre…"), Tâche (liste + "Autre…"), Durée en heures décimales (1,5 = 1h30). Un "Autre" saisi sur Lieu ou Tâche s'ajoute automatiquement à la liste de référence (colonnes H et Q) et au tableau récapitulatif correspondant, pour être proposé directement au prochain essai. L'insertion respecte l'ordre chronologique du journal, comme pour l'irrigation.

**Consultation** : le total par personne apparaît en premier (priorité donnée à l'info utile pour les payes), puis le total par tâche, puis par lieu.

---

L'appli est composée de 4 fichiers : `index.html`, `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png`, et un script séparé `Code.gs` à coller dans Google Sheets. Aucun de ces fichiers ne doit être renommé.

Cette version colle exactement à la structure de `irrigation_truffiere_2026.xlsx` : chaque onglet (**La Fontaine**, **La baume**, **Pilles**, **Bousquet et Riou**) contient 2 ou 3 "lignes" d'irrigation côte à côte (ex. Bousquet et Riou = Ligne 1 "Bousquet" + Ligne 2 "Riou"). L'appli te fait donc choisir d'abord la parcelle, puis la ligne, avant le formulaire.

Le script retrouve automatiquement la bonne colonne et la bonne ligne libre dans le tableau — il ne touche jamais aux colonnes **Débit (m3/h)** et **mm**, qui restent calculées par formule, ni aux lignes "TOTAL Mois". Il ne remplit que : Date, Durée, Compteur début, Compteur fin et M3.

## Étape 0 ter — Ajouter la colonne "Pluie (mm)" (à faire une seule fois)

Cette fois, pas besoin de toucher aux colonnes à la main : le nouveau `Code.gs` contient une fonction qui le fait pour toi.

1. Colle le nouveau `Code.gs` dans Apps Script (remplace tout l'ancien contenu).
2. En haut de l'éditeur Apps Script, à côté du bouton **Exécuter**, il y a un menu déroulant de fonctions — sélectionne **ajouterColonnePluie**.
3. Clique sur **Exécuter**. Autorise les permissions si demandé.
4. Une fois terminé, va dans **Affichage > Journal d'exécution** (ou l'icône horloge "Exécutions") pour vérifier : chaque ligne de chaque onglet doit indiquer "colonne ajoutée en colonne X, ligne d'en-tête Y".
5. Vérifie rapidement dans le Sheet que la colonne "Pluie (mm)" apparaît bien juste après "mm" dans chacun des 11 blocs.

⚠️ Ne relance pas cette fonction une deuxième fois par erreur (elle se protège en ignorant les blocs déjà traités, mais mieux vaut ne pas y toucher une fois que c'est fait).

## Nouveau fonctionnement : Pluie / Arrosage

L'appli démarre maintenant sur un choix **Arrosage** ou **Pluie**.

- **Arrosage** : exactement comme avant (parcelle → ligne → formulaire complet). Les données s'écrivent en **noir**.
- **Pluie** : un seul écran avec juste **Date** et **mm**. Pas besoin de choisir de parcelle — la valeur est automatiquement inscrite dans les 11 blocs (toutes les parcelles, toutes les lignes) en même temps, dans la colonne "Pluie (mm)", en **bleu foncé**.

Ça te permet d'analyser séparément les apports par pluie (colonne "Pluie (mm)") et par arrosage (colonne "mm") dans tes propres calculs de synthèse.

## Étape 0 — Ajouter les colonnes Heure début / Heure fin (à faire une seule fois)

Ces deux colonnes ne font pas partie du tableau original — il faut les insérer toi-même, **juste après la colonne "Date" et avant "Durée (h)"**, dans chacun des 11 blocs (les onglets ont 2 ou 3 "lignes" côte à côte). Fais bien attention à traiter chaque onglet **de droite à gauche** (Ligne 3 avant Ligne 2, avant Ligne 1) — sinon les colonnes se décalent et tu perds le repère.

Pour chaque bloc : sélectionne les 2 colonnes situées juste après "Date" (clique sur l'en-tête de la 1ère, puis Maj+clique sur la 2e), clic droit → **Insérer 2 colonnes à gauche**. Puis tape "Heure début" et "Heure fin" dans les 2 nouvelles cellules, sur la même ligne que les autres en-têtes ("Date", "Durée (h)"...).

**Onglet La Fontaine** (en-têtes en ligne 5) :
1. Ligne 3 (Date = colonne P) → sélectionne Q:R → Insérer 2 colonnes à gauche → Q5 = "Heure début", R5 = "Heure fin"
2. Ligne 2 (Date = colonne I) → sélectionne J:K → Insérer 2 colonnes à gauche → J5 = "Heure début", K5 = "Heure fin"
3. Ligne 1 (Date = colonne B) → sélectionne C:D → Insérer 2 colonnes à gauche → C5 = "Heure début", D5 = "Heure fin"

**Onglets La baume et Pilles** (en-têtes en ligne 3, même disposition B/I/P) :
1. Ligne 3 (Date = P) → Q:R → Insérer 2 colonnes à gauche → Q3 = "Heure début", R3 = "Heure fin"
2. Ligne 2 (Date = I) → J:K → Insérer 2 colonnes à gauche → J3 = "Heure début", K3 = "Heure fin"
3. Ligne 1 (Date = B) → C:D → Insérer 2 colonnes à gauche → C3 = "Heure début", D3 = "Heure fin"

**Onglet Bousquet et Riou** (en-têtes en ligne 3, seulement 2 lignes) :
1. Ligne 2 / Riou (Date = I) → J:K → Insérer 2 colonnes à gauche → J3 = "Heure début", K3 = "Heure fin"
2. Ligne 1 / Bousquet (Date = B) → C:D → Insérer 2 colonnes à gauche → C3 = "Heure début", D3 = "Heure fin"

Une fois les 11 blocs traités, vérifie rapidement que les colonnes Débit/mm affichent toujours des résultats cohérents (les formules se décalent automatiquement avec l'insertion, elles ne devraient pas casser) — et que les libellés "Ligne 1/2/3" en haut de chaque bloc couvrent toujours bien toute la largeur du bloc.

💡 Conseil : fais une copie de sauvegarde du fichier avant de commencer (**Fichier > Créer une copie**), au cas où.

## Étape 1 — Créer le script Google Apps Script

1. Ouvre ton Google Sheet.
2. Menu **Extensions > Apps Script**.
3. Supprime le code par défaut, colle le contenu de `Code.gs`.
4. Clique **Déployer > Nouveau déploiement**.
5. Type de déploiement : **Application Web**.
6. "Exécuter en tant que" : **Moi**. "Qui a accès" : **Tout le monde**.
7. Clique **Déployer**, accepte les autorisations Google demandées.
8. Copie l'URL fournie (elle ressemble à `https://script.google.com/macros/s/XXXXX/exec`).

## Étape 2 — Configurer l'appli

Ouvre `index.html`, trouve la ligne suivante en haut du `<script>` :

```js
const SCRIPT_URL = "COLLE_ICI_TON_URL_APPS_SCRIPT";
```

Remplace par ton URL copiée à l'étape 2.

## Étape 3 — Mettre l'appli en ligne (nécessaire pour l'installer sur Android)

Le plus simple, vu que tu utilises déjà GitHub : crée un nouveau dépôt, ajoute les 5 fichiers (`index.html`, `manifest.json`, `sw.js`, `icon-192.png`, `icon-512.png` — pas besoin de `Code.gs` ni de ce fichier), puis active **GitHub Pages** :

1. Sur GitHub, crée un repo (ex : `irrigation-app`), ajoute les fichiers.
2. Va dans **Settings > Pages**.
3. Source : **Deploy from a branch**, branche `main`, dossier `/ (root)`.
4. Après 1-2 minutes, ton appli est en ligne à une adresse du type :
   `https://ton-pseudo.github.io/irrigation-app/`

(Toute autre solution d'hébergement HTTPS statique fonctionne aussi : Netlify, Vercel, etc. GitHub Pages est gratuit et simple vu que tu as déjà l'habitude de Git.)

## Étape 4 — Installer sur ton téléphone Android

1. Ouvre l'URL ci-dessus dans **Chrome** sur ton téléphone.
2. Menu ⋮ (trois points) > **Installer l'application** (ou "Ajouter à l'écran d'accueil").
3. Une icône apparaît sur ton écran d'accueil, elle s'ouvre en plein écran comme une vraie appli.

## Utilisation au quotidien

1. Ouvre l'appli → choisis la parcelle (onglet).
2. Choisis la ligne (ex : Bousquet ou Riou).
3. Renseigne date, durée, compteur début, compteur fin — le volume M3 se calcule automatiquement (modifiable si besoin).
4. Appuie sur **Enregistrer**.
5. La donnée part dans la première ligne encore libre du bon bloc de colonnes, dans le bon onglet — sans jamais toucher aux formules Débit/mm ni aux lignes TOTAL. Le compteur début suivant est pré-rempli avec le compteur fin que tu viens de saisir, pour aller plus vite sur la saisie suivante.

## Si quelque chose ne fonctionne pas

- **"Onglet introuvable"** : le nom de l'onglet dans le Sheet ne correspond pas exactement — vérifie l'orthographe (espaces, accents). Les onglets attendus sont : La Fontaine, La baume, Pilles, Bousquet et Riou.
- **"Bloc 'Ligne N' introuvable"** : le libellé "Ligne 1" / "Ligne 2" / "Ligne 3" a été renommé ou supprimé dans le Sheet — l'appli s'appuie sur ce texte pour retrouver la bonne colonne.
- **"Plus de ligne libre"** : le tableau pré-construit (jusqu'à la ligne 34-41 selon l'onglet) est entièrement rempli — il faut l'étendre manuellement dans le Sheet (copier les formules Débit/mm vers le bas) avant de continuer à saisir depuis l'appli.
- **Échec de l'envoi** : vérifie la connexion internet du téléphone, et que l'URL dans `SCRIPT_URL` est bien complète (se terminant par `/exec`).
- Si tu modifies `Code.gs` plus tard, il faut redéployer une **nouvelle version** (Déployer > Gérer les déploiements > icône crayon > Version : Nouvelle version), sinon Google continue de servir l'ancienne version.

## Mise à jour : vérification, correction, journal

- **Écran « Vérifie »** avant chaque enregistrement (pluie, arrosage, heures, poids, vente) : résumé + alertes (doublon, date future ou ancienne, valeur aberrante, compteur incohérent, plus de 12 h par jour…). Alerte orange = on peut quand même enregistrer ; alerte rouge = bloquant. Pour la pluie : bouton « Remplacer l'existant ».
- **✏️ Corriger** (en haut de l'accueil) : les 12 dernières saisies de chaque module, avec Modifier / Supprimer. Le script vérifie que la ligne n'a pas changé entre-temps.
- **Onglet « Journal »** (créé automatiquement) : chaque modification / suppression y est tracée (avant / après).
- Les noms (personnes, lieux, tâches, clients) sont normalisés à l'enregistrement. « Brumale » est écrit aussi dans le commentaire ; les lignes dont le commentaire contient « brumale » sont roses.

### Déploiement
1. Apps Script : remplacer tout `Code.gs`, puis Déployer > Gérer les déploiements > Modifier > Nouvelle version.
2. Lancer une fois `colorerBrumale` (règle rose étendue au commentaire).
3. GitHub Desktop : remplacer `index.html` et `sw.js` (cache v9), commit + push, puis Ctrl+Maj+R.
