/**
 * À COLLER DANS : Google Sheet > Extensions > Apps Script (remplace tout le contenu existant)
 *
 * Structure de chaque bloc "Ligne N" dans les onglets La Fontaine / La baume / Pilles / Bousquet et Riou :
 * Date | Heure début | Heure fin | Durée (h) | Compteur début | Compteur fin
 * | Débit (m3/h) [formule] | M3 | mm [formule] | Pluie (mm)
 *
 * Deux types de saisie depuis l'appli :
 *  - ARROSAGE : remplit Date/Heure/Durée/Compteurs/M3 sur la ligne (parcelle+ligne) choisie.
 *               Écriture en NOIR.
 *  - PLUIE    : remplit seulement Date + "Pluie (mm)", mais sur les 11 blocs à la fois
 *               (toutes les lignes de toutes les parcelles). Écriture en ORANGE FONCÉ.
 *
 * Chaque saisie est insérée à la bonne place chronologique DANS LE MOIS EN COURS (segment
 * entre deux lignes TOTAL) : si une date est antérieure à une entrée déjà présente dans ce
 * segment, les lignes concernées sont décalées d'un cran vers le bas. Aucun décalage n'a lieu
 * au-delà d'une ligne TOTAL (pas de réordonnancement entre deux mois différents).
 *
 * La colonne Date est toujours affichée sans l'heure (format jj/mm/aaaa).
 *
 * MISE EN PLACE :
 *  1. Colle ce code dans Extensions > Apps Script.
 *  2. AVANT tout déploiement : sélectionne la fonction "ajouterColonnePluie" dans le menu
 *     déroulant en haut (à côté de "Exécuter" / "Debug"), puis clique "Exécuter" — une seule fois.
 *     Ça ajoute la colonne "Pluie (mm)" dans les 11 blocs. Vérifie ensuite le résultat dans
 *     "Affichage > Journal d'exécution" (ou l'icône horloge "Exécutions").
 *  3. Déployer > Gérer les déploiements > crayon > Version : Nouvelle version > Déployer.
 *     (ou "Nouveau déploiement" si c'est la première fois).
 *  4. Colle l'URL fournie dans SCRIPT_URL en haut de index.html.
 */

var PARCELLES = [
  { sheet: "La Fontaine", lignes: [1, 2, 3] },
  { sheet: "La baume", lignes: [1, 2, 3] },
  { sheet: "Pilles", lignes: [1, 2, 3] },
  { sheet: "Bousquet et Riou", lignes: [1, 2] }
];

var COULEUR_ARROSAGE = "#000000"; // noir
var COULEUR_PLUIE = "#B5540A";    // orange foncé
var FORMAT_DATE = "dd/mm/yyyy";   // date seule, sans heure

// Construit une date "pure" (sans dépendre d'une interprétation d'heure précise) à partir
// d'une chaîne "AAAA-MM-JJ", en fixant explicitement année/mois/jour plutôt qu'en passant par
// un parsing de chaîne ISO. Fixée à MIDI (et non minuit) : Apps Script (moteur V8) construit
// et lit les dates en UTC en interne, indépendamment du fuseau horaire réel du Sheet — avec
// minuit, un simple décalage de fuseau suffit à faire basculer l'affichage sur la veille.
// Midi laisse une marge de ±11h avant de risquer de changer de jour, ce qui couvre tous les
// fuseaux horaires réalistes.
function parseDateOnly(dateStr) {
  var parts = dateStr.split("-");
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10);
  var d = parseInt(parts[2], 10);
  return new Date(y, m - 1, d, 12, 0, 0);
}

// ---------- Structure de l'onglet "Heures" ----------
// Tableau brut (une ligne = une saisie) : Date | Personne | Lieux | Tache | Temps passé (h)
var HEURES_SHEET = "Heures";
var HEURES_COL_DATE = 2;     // B
var HEURES_COL_PERSONNE = 3; // C
var HEURES_COL_LIEU = 4;     // D
var HEURES_COL_TACHE = 5;    // E
var HEURES_COL_TEMPS = 6;    // F
// Colonne G : actuellement vide dans le Sheet (entre "Temps" et le tableau récapitulatif
// Tâche qui commence en H) — utilisée pour horodater discrètement chaque saisie, afin que
// l'appli puisse afficher "les dernières saisies" dans leur véritable ordre d'enregistrement,
// même quand elles sont insérées à une autre place pour respecter l'ordre chronologique des
// dates. Ça ne change rien à l'ordre des lignes dans le Sheet lui-même.
var HEURES_COL_HORODATAGE = 7; // G
var HEURES_DATA_START_ROW = 3;

// Liste de référence "Tache" (lignes), utilisée aussi comme lignes du tableau récapitulatif
// "Total tache" : colonne H, à partir de la ligne 4. Récapitulatif (Paul..Externe 3 + Total)
// sur 7 colonnes, de I à O.
var HEURES_TACHE_LIST_COL = 8;  // H
var HEURES_TACHE_LIST_START = 4;
var HEURES_TACHE_RECAP_WIDTH = 7; // I,J,K,L,M,N,O

// Liste de référence "Lieux" (lignes), utilisée aussi comme lignes du tableau récapitulatif
// "Total lieu" : colonne Q, à partir de la ligne 4. Récapitulatif sur 7 colonnes, de R à X.
var HEURES_LIEU_LIST_COL = 17; // Q
var HEURES_LIEU_LIST_START = 4;
var HEURES_LIEU_RECAP_WIDTH = 7; // R,S,T,U,V,W,X

// Personnes : liste FIXE à 6 emplacements (plus de saisie libre pour ce champ — chaque
// externe a désormais sa propre colonne/ligne dédiée, à choisir directement).
var HEURES_PERSONNES = ["Paul", "René", "Charles", "Externe 1", "Externe 2", "Externe 3"];
// Lignes du tableau récapitulatif mensuel par personne (colonne Z, à partir de la ligne 3),
// dans le même ordre que HEURES_PERSONNES.
var HEURES_PERSONNE_LIST_COL = 26; // Z
var HEURES_PERSONNE_LIST_START = 3;

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.mode === "pluie") {
      return handlePluie(data);
    }
    if (data.mode === "heures") {
      return handleHeures(data);
    }
    if (data.mode === "recolte") {
      return handleRecolte(data);
    }
    if (data.mode === "recolteVolume") {
      return handleRecolteVolume(data);
    }
    if (data.mode === "vente") {
      return handleVente(data);
    }
    return handleArrosage(data);
  } catch (err) {
    return jsonOutput({ ok: false, error: err.message });
  }
}

function doGet(e) {
  var action = e && e.parameter ? e.parameter.action : null;

  if (action === "bilans") {
    var data = obtenirTousLesBilans();
    var json = JSON.stringify(data);
    if (e.parameter.callback) {
      // JSONP : contourne le blocage CORS en lecture, puisqu'une balise <script> n'est
      // pas soumise aux mêmes restrictions qu'une requête fetch/XHR classique.
      return ContentService.createTextOutput(e.parameter.callback + "(" + json + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "listesHeures") {
    var listes = obtenirListesHeures();
    var json2 = JSON.stringify(listes);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + "(" + json2 + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json2).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "bilanHeures") {
    var bh = obtenirBilanHeures();
    var json3 = JSON.stringify(bh);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + "(" + json3 + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json3).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "dernieresSaisiesHeures") {
    var recentes = obtenirDernieresSaisiesHeures();
    var json4 = JSON.stringify(recentes);
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + "(" + json4 + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json4).setMimeType(ContentService.MimeType.JSON);
  }

  var actionsRecolte = {
    dernieresRecoltes: obtenirDernieresRecoltes,
    pointsRecolte: obtenirPointsRecolte,
    dernieresPesees: obtenirDernieresPesees,
    tableauPoids: obtenirTableauPoids,
    dernieresVentes: obtenirDernieresVentes,
    tableauVente: obtenirTableauVente,
    listesVente: obtenirListesVente,
    urlFichier: function () { return { url: SpreadsheetApp.getActiveSpreadsheet().getUrl() }; }
  };
  if (actionsRecolte[action]) {
    var json5 = JSON.stringify(actionsRecolte[action]());
    if (e.parameter.callback) {
      return ContentService.createTextOutput(e.parameter.callback + "(" + json5 + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json5).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("L'API fonctionne. Utilise une requête POST pour envoyer des données.");
}

// Libellés d'affichage pour chaque ligne, réutilisés côté consultation
var LABELS_LIGNE = {
  "La Fontaine": { 1: "Ligne 1", 2: "Ligne 2", 3: "Ligne 3" },
  "La baume": { 1: "Ligne 1 · Grande parcelle", 2: "Ligne 2 · Sud", 3: "Ligne 3 · Petite parcelle (bas)" },
  "Pilles": { 1: "Ligne 1", 2: "Ligne 2", 3: "Ligne 3" },
  "Bousquet et Riou": { 1: "Bousquet", 2: "Riou" }
};

// Ligne d'en-tête « Mois » du tableau de bilan d'un bloc : d'abord AU-DESSUS des données (tableaux
// en haut de l'onglet), sinon en dessous (ancienne disposition). -1 si introuvable.
function trouverEnteteBilan(sheet, block) {
  var col = block.blockCol;
  var max = Math.min(sheet.getMaxRows(), block.headerRow + 120);
  var vals = sheet.getRange(1, col, max, 1).getValues();
  var r;
  for (r = 0; r < block.headerRow - 1; r++) {
    if (normTxt(vals[r][0]) === "mois") return r + 1;
  }
  for (r = block.headerRow; r < max; r++) {
    if (normTxt(vals[r][0]) === "mois") return r + 1;
  }
  return -1;
}

// Récupère, pour les 11 blocs, le tableau de bilan mensuel déjà calculé (formules déjà
// en place dans le Sheet) : Mois | Pluie (mm) | Arrosage (mm) | cumul | Besoins (Etc).
// Toutes les lignes du tableau sont lues (mai à avril + la ligne TOTAL période estivale).
function obtenirTousLesBilans() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) return;

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) return;

      var headerRow2 = trouverEnteteBilan(sheet, block);
      if (headerRow2 === -1) return;

      var nb = Math.min(16, sheet.getMaxRows() - headerRow2);
      var rows = sheet.getRange(headerRow2 + 1, block.blockCol, nb, 5).getValues();
      var mois = [];
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var label = row[0] ? row[0].toString().trim() : "";
        if (!label) break; // fin du tableau
        mois.push({
          mois: label,
          total: /^total/i.test(label),
          pluie: Number(row[1]) || 0,
          arrosage: Number(row[2]) || 0,
          cumul: Number(row[3]) || 0,
          besoins: Number(row[4]) || 0
        });
      }

      var labels = LABELS_LIGNE[p.sheet] || {};
      out.push({
        onglet: p.sheet,
        ligne: ligneIndex,
        label: labels[ligneIndex] || ("Ligne " + ligneIndex),
        mois: mois
      });
    });
  });

  return out;
}

/* ---------- ARROSAGE (une seule parcelle + ligne) ---------- */

function handleArrosage(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = findSheetByName(ss, data.onglet);
  if (!sheet) {
    return jsonOutput({ ok: false, error: "Onglet introuvable : " + data.onglet });
  }

  var ligneIndex = parseInt(data.ligne, 10);
  if (!ligneIndex) {
    return jsonOutput({ ok: false, error: "Numéro de ligne manquant." });
  }

  var block = findLigneBlock(sheet, ligneIndex);
  if (!block) {
    return jsonOutput({ ok: false, error: "Bloc 'Ligne " + ligneIndex + "' introuvable dans l'onglet " + data.onglet });
  }

  var cols = blockColumns(block.blockCol);
  var dateValue = data.date ? parseDateOnly(data.date) : "";
  var targetRow = resolveTargetRow(sheet, block, cols, dateValue);
  if (targetRow === -1) {
    return jsonOutput({ ok: false, error: "Plus de ligne libre dans le tableau — il faut l'étendre manuellement." });
  }

  setValueColored(sheet.getRange(targetRow, cols.date), dateValue, COULEUR_ARROSAGE).setNumberFormat(FORMAT_DATE);
  setValueColored(sheet.getRange(targetRow, cols.duree), Number(data.duree), COULEUR_ARROSAGE);
  setValueColored(sheet.getRange(targetRow, cols.debut), Number(data.compteur_debut), COULEUR_ARROSAGE);
  setValueColored(sheet.getRange(targetRow, cols.fin), Number(data.compteur_fin), COULEUR_ARROSAGE);
  setValueColored(sheet.getRange(targetRow, cols.m3), Number(data.m3), COULEUR_ARROSAGE);

  // Heure début / Heure fin sont facultatives : rien n'est écrit si elles sont vides.
  if (data.heure_debut) {
    var hd = sheet.getRange(targetRow, cols.heureDebut);
    hd.setValue(parseHeure(data.heure_debut)).setNumberFormat("HH:mm").setFontColor(COULEUR_ARROSAGE);
  }
  if (data.heure_fin) {
    var hf = sheet.getRange(targetRow, cols.heureFin);
    hf.setValue(parseHeure(data.heure_fin)).setNumberFormat("HH:mm").setFontColor(COULEUR_ARROSAGE);
  }

  return jsonOutput({ ok: true, row: targetRow });
}

/* ---------- PLUIE (toutes les parcelles / lignes à la fois) ---------- */

function handlePluie(data) {
  var mm = Number(data.mm);
  if (!data.date || isNaN(mm)) {
    return jsonOutput({ ok: false, error: "Date ou mm manquant pour la pluie." });
  }
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateValue = parseDateOnly(data.date);
  var details = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { details.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { details.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var cols = blockColumns(block.blockCol);
      var targetRow = resolveTargetRow(sheet, block, cols, dateValue);
      if (targetRow === -1) { details.push(p.sheet + " Ligne " + ligneIndex + " : plus de ligne libre"); return; }

      setValueColored(sheet.getRange(targetRow, cols.date), dateValue, COULEUR_PLUIE).setNumberFormat(FORMAT_DATE);
      setValueColored(sheet.getRange(targetRow, cols.pluie), mm, COULEUR_PLUIE);
      details.push(p.sheet + " Ligne " + ligneIndex + " : ligne " + targetRow + " OK");
    });
  });

  return jsonOutput({ ok: true, details: details });
}

/* ================================================================
   HEURES — saisie du temps de travail (Nom, Date, Lieu, Tâche, Durée)
   ================================================================ */

function handleHeures(data) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) return jsonOutput({ ok: false, error: "Onglet 'Heures' introuvable." });

  if (!data.date || !data.nom || !data.lieu || !data.tache || data.duree === undefined || data.duree === "") {
    return jsonOutput({ ok: false, error: "Champs manquants pour la saisie heures." });
  }
  if (HEURES_PERSONNES.indexOf(data.nom) === -1) {
    return jsonOutput({ ok: false, error: "Personne inconnue : " + data.nom });
  }

  var dateValue = parseDateOnly(data.date);
  var duree = Number(data.duree);
  if (isNaN(duree) || duree <= 0) {
    return jsonOutput({ ok: false, error: "Durée invalide." });
  }

  // Si "lieu" ou "tache" est une valeur inédite (saisie via "Autre"), on l'ajoute à la liste
  // de référence correspondante (et donc aux lignes du tableau récapitulatif concerné).
  // "nom" est toujours l'un des 6 emplacements fixes (Paul/René/Charles/Externe 1/2/3) —
  // pas de liste à étendre pour ce champ.
  ajouterSiNouveau(sheet, HEURES_LIEU_LIST_COL, HEURES_LIEU_LIST_START, data.lieu, HEURES_LIEU_RECAP_WIDTH);
  ajouterSiNouveau(sheet, HEURES_TACHE_LIST_COL, HEURES_TACHE_LIST_START, data.tache, HEURES_TACHE_RECAP_WIDTH);

  var targetRow = resolveTargetRowHeures(sheet, dateValue);
  if (targetRow === -1) {
    return jsonOutput({ ok: false, error: "Plus de ligne libre dans le tableau des heures — il faut l'étendre manuellement." });
  }

  var dateCell = sheet.getRange(targetRow, HEURES_COL_DATE);
  dateCell.setValue(dateValue).setNumberFormat(FORMAT_DATE);
  sheet.getRange(targetRow, HEURES_COL_PERSONNE).setValue(data.nom);
  sheet.getRange(targetRow, HEURES_COL_LIEU).setValue(data.lieu);
  sheet.getRange(targetRow, HEURES_COL_TACHE).setValue(data.tache);
  sheet.getRange(targetRow, HEURES_COL_TEMPS).setValue(duree);
  sheet.getRange(targetRow, HEURES_COL_HORODATAGE).setValue(new Date()).setNumberFormat("dd/mm/yyyy hh:mm:ss");

  return jsonOutput({ ok: true, row: targetRow });
}

// Ajoute "valeur" à la liste verticale (colonne col, à partir de startRow) si elle n'y figure
// pas déjà. Recopie aussi les formules du récapitulatif (colonnes à droite, sur largeurRecap
// colonnes) depuis la ligne du dessus, pour que la nouvelle ligne soit immédiatement prise en
// compte dans les totaux.
function ajouterSiNouveau(sheet, col, startRow, valeur, largeurRecap) {
  var v = (valeur || "").toString().trim();
  if (!v) return;

  var maxScan = 200;
  var lastRow = startRow - 1;
  for (var r = startRow; r < startRow + maxScan; r++) {
    var cell = sheet.getRange(r, col).getValue();
    if (!cell || cell.toString().trim() === "") { lastRow = r - 1; break; }
    if (cell.toString().trim().toLowerCase() === v.toLowerCase()) return; // déjà présent
    lastRow = r;
  }

  var nouvelleRow = lastRow + 1;
  sheet.getRange(nouvelleRow, col).setValue(v);

  // Recopie les formules du récapitulatif (colonnes à droite de la liste) depuis la ligne
  // au-dessus, pour que les totaux de cette nouvelle ligne se calculent tout de suite.
  if (nouvelleRow > startRow) {
    var src = sheet.getRange(nouvelleRow - 1, col + 1, 1, largeurRecap);
    var dest = sheet.getRange(nouvelleRow, col + 1, 1, largeurRecap);
    src.copyTo(dest, { formulaOnly: true });
  }
}

// Trouve la ligne où écrire, en respectant l'ordre chronologique de l'ensemble du tableau
// (pas de segmentation par mois ici, contrairement aux tableaux d'irrigation — tout le
// journal des heures est une seule séquence continue).
function resolveTargetRowHeures(sheet, dateValue) {
  var maxScan = 2000;
  var nextFreeRow = -1;
  for (var r = HEURES_DATA_START_ROW; r < HEURES_DATA_START_ROW + maxScan; r++) {
    var v = sheet.getRange(r, HEURES_COL_DATE).getValue();
    if (!v || v === "") { nextFreeRow = r; break; }
  }
  if (nextFreeRow === -1) return -1;

  var insertionRow = nextFreeRow;
  for (var rr = HEURES_DATA_START_ROW; rr < nextFreeRow; rr++) {
    var d = sheet.getRange(rr, HEURES_COL_DATE).getValue();
    if (d instanceof Date && dateValue.getTime() < d.getTime()) {
      insertionRow = rr;
      break;
    }
  }

  if (insertionRow < nextFreeRow) {
    shiftBlockRowsDown(sheet, HEURES_COL_DATE, 6, insertionRow, nextFreeRow); // B à G = 6 colonnes (inclut l'horodatage)
  }

  return insertionRow;
}

// Renvoie les listes actuelles pour remplir les menus déroulants de l'appli : personnes
// (liste fixe à 6 emplacements), lieux et taches (listes de référence, extensibles via "Autre").
function obtenirListesHeures() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) return { personnes: HEURES_PERSONNES, lieux: [], taches: [] };

  var lieux = listeVerticale(sheet, HEURES_LIEU_LIST_COL, HEURES_LIEU_LIST_START);
  var taches = listeVerticale(sheet, HEURES_TACHE_LIST_COL, HEURES_TACHE_LIST_START);

  return { personnes: HEURES_PERSONNES, lieux: lieux, taches: taches };
}

function listeVerticale(sheet, col, startRow) {
  var maxScan = 200;
  var out = [];
  for (var r = startRow; r < startRow + maxScan; r++) {
    var v = sheet.getRange(r, col).getValue();
    if (!v || v.toString().trim() === "") break;
    out.push(v.toString().trim());
  }
  return out;
}

// Bilan pour la consultation dans l'appli : priorité au total par personne (calculé
// directement depuis le journal, indépendamment des colonnes fixes du tableau récapitulatif),
// puis totaux par tâche et par lieu.
function obtenirBilanHeures() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) return { parPersonne: [], parTache: [], parLieu: [] };

  var data = lireBlocBorne(sheet, HEURES_DATA_START_ROW, HEURES_COL_DATE, 5, 2000); // B:F (borné à la taille réelle de la feuille)

  var parPersonne = {}, parTache = {}, parLieu = {};
  data.forEach(function (row) {
    var date = row[0], personne = (row[1] || "").toString().trim(), lieu = (row[2] || "").toString().trim();
    var tache = (row[3] || "").toString().trim(), temps = Number(row[4]) || 0;
    if (!date || !personne || temps <= 0) return;

    parPersonne[personne] = (parPersonne[personne] || 0) + temps;
    if (lieu) parLieu[lieu] = (parLieu[lieu] || 0) + temps;
    if (tache) parTache[tache] = (parTache[tache] || 0) + temps;
  });

  function versListe(obj) {
    return Object.keys(obj).sort().map(function (k) { return { label: k, total: obj[k] }; })
      .sort(function (a, b) { return b.total - a.total; });
  }

  return { parPersonne: versListe(parPersonne), parTache: versListe(parTache), parLieu: versListe(parLieu) };
}

// Renvoie les 5 dernières lignes (par position, donc les plus récemment ajoutées au bas du
// tableau) du journal B:F, pour que l'appli puisse afficher un contrôle visuel juste après
// une saisie — utile puisque l'appli ne peut pas lire la réponse du serveur (CORS).
function obtenirDernieresSaisiesHeures() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) return [];

  var maxScan = 2000;
  var lastRow = -1;
  for (var r = HEURES_DATA_START_ROW; r < HEURES_DATA_START_ROW + maxScan; r++) {
    var v = sheet.getRange(r, HEURES_COL_DATE).getValue();
    if (!v || v === "") break;
    lastRow = r;
  }
  if (lastRow === -1) return [];

  var nbLignes = lastRow - HEURES_DATA_START_ROW + 1;
  var data = sheet.getRange(HEURES_DATA_START_ROW, HEURES_COL_DATE, nbLignes, 6).getValues(); // B:G

  var lignes = data.map(function (row) {
    var horodatage = row[5] instanceof Date ? row[5].getTime() : 0; // lignes saisies avant
    // l'ajout de cette colonne : pas d'horodatage, elles passent naturellement en dernier.
    return {
      date: row[0] instanceof Date ? formatDateFR(row[0]) : (row[0] || ""),
      personne: row[1] || "",
      lieu: row[2] || "",
      tache: row[3] || "",
      temps: Number(row[4]) || 0,
      _horodatage: horodatage
    };
  });

  // Ordre de SAISIE réel (le plus récemment enregistré en premier), pas l'ordre des lignes
  // dans le tableau — qui lui reste chronologique par date de travail, inchangé sur le Sheet.
  lignes.sort(function (a, b) { return b._horodatage - a._horodatage; });

  return lignes.slice(0, 5).map(function (l) {
    return { date: l.date, personne: l.personne, lieu: l.lieu, tache: l.tache, temps: l.temps };
  });
}

function formatDateFR(d) {
  var jj = ("0" + d.getDate()).slice(-2);
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  return jj + "/" + mm + "/" + d.getFullYear();
}

/* ---------- Diagnostic ponctuel : fuseau horaire du classeur ---------- */
// À exécuter pour vérifier ce qui explique un éventuel décalage d'heure (ex. colonne
// "Horodatage" de l'onglet Heures affichant une heure incohérente). Le fuseau horaire du
// classeur (visible aussi dans Fichier > Paramètres du classeur > Fuseau horaire) doit
// normalement être "Europe/Paris" pour une exploitation en France. S'il affiche autre chose,
// c'est très probablement la cause du décalage — il suffit de le corriger dans ce menu.
/* ================================================================
   RÉCOLTE — onglet "Récolte pointage" (une ligne par truffe pointée)
   Colonnes : A Truffes No | B Date | C Parcelle | D Poids | E Qualité | F Image
              | G Localisation | H Commentaires | I Latitude | J Longitude
   (E et F restent à remplir à la main ; I et J sont des colonnes techniques
   « lat / lon » en texte, utilisées par la carte de l'appli et pour Google My Maps.)
   + onglet "Récolte volume" (poids total après récolte, par parcelle)
   ================================================================ */
var RECOLTE_SHEET = "Récolte pointage";
var RECOLTE_DATA_START_ROW = 3;
var RECOLTE_COL_NO = 1;       // A
var RECOLTE_COL_DATE = 2;     // B
var RECOLTE_COL_PARCELLE = 3; // C
var RECOLTE_COL_POIDS = 4;    // D
var RECOLTE_COL_LOC = 7;      // G
var RECOLTE_COL_COMMENT = 8;  // H
var RECOLTE_COL_LAT = 9;      // I
var RECOLTE_COL_LON = 10;     // J
var RECOLTE_NB_COLS = 10;     // A à J (pour le décalage chronologique)
var RECOLTE_PARCELLES = ["La Fontaine", "La Baume", "Bousquet", "Riou", "Les Pilles"];
var RECOLTE_POIDS = ["<10 g", "20 g", "50 g", "100 g", ">100 g"];
var VOLUME_SHEET = "Récolte volume";

function normTxt(t) {
  return (t || "").toString().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim().toLowerCase();
}

// Recherche d'onglet insensible à la casse ET aux accents ("Récolte" = "recolte").
function trouverOngletSouple(ss, nom) {
  var voulu = normTxt(nom), trouve = null;
  ss.getSheets().forEach(function (s) { if (normTxt(s.getName()) === voulu) trouve = s; });
  return trouve;
}

// Onglet du pointage des truffes : « Récolte pointage » (ancien nom « Récolte » accepté).
function ongletPointage(ss) {
  return trouverOngletSouple(ss, RECOLTE_SHEET) || trouverOngletSouple(ss, "Récolte");
}

// Lit jusqu'à nbMax lignes à partir de startRow, sans jamais dépasser la taille réelle de la
// feuille (une feuille Google fait souvent 1000 lignes : lire au-delà provoque une erreur).
function lireBlocBorne(sheet, startRow, col, nbCols, nbMax) {
  var dispo = sheet.getMaxRows() - startRow + 1;
  var n = Math.max(1, Math.min(nbMax, dispo));
  return sheet.getRange(startRow, col, n, nbCols).getValues();
}

// Première ligne libre (date vide) dans le bloc lu ; ajoute des lignes à la feuille si le
// tableau a rempli toute la grille.
function premiereLigneLibre(sheet, startRow, valeursDate, nbMax) {
  for (var i = 0; i < valeursDate.length; i++) {
    if (valeursDate[i] === "" || valeursDate[i] === null) return startRow + i;
  }
  if (valeursDate.length < nbMax) {
    sheet.insertRowsAfter(sheet.getMaxRows(), 200);
    return startRow + valeursDate.length;
  }
  return -1;
}

function handleRecolte(data) {
  var sheet = ongletPointage(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet) return jsonOutput({ ok: false, error: "Onglet 'Récolte pointage' introuvable." });

  if (RECOLTE_PARCELLES.indexOf(data.parcelle) === -1) return jsonOutput({ ok: false, error: "Parcelle inconnue : " + data.parcelle });
  if (RECOLTE_POIDS.indexOf(data.poids) === -1) return jsonOutput({ ok: false, error: "Poids inconnu : " + data.poids });
  if (!data.date) return jsonOutput({ ok: false, error: "Date manquante." });

  var dateValue = parseDateOnly(data.date);

  var NB_MAX = 3000;
  var ab = lireBlocBorne(sheet, RECOLTE_DATA_START_ROW, 1, 2, NB_MAX); // colonnes A (n°) et B (date)
  var maxNo = 0;
  ab.forEach(function (row) { var no = Number(row[0]); if (no > maxNo) maxNo = no; });
  var nextFreeRow = premiereLigneLibre(sheet, RECOLTE_DATA_START_ROW, ab.map(function (row) { return row[1]; }), NB_MAX);
  if (nextFreeRow === -1) return jsonOutput({ ok: false, error: "Tableau Récolte plein." });

  // Ordre chronologique : la saisie va à la bonne place selon sa date (cas rare, car la date
  // est celle du jour ; une récolte du jour s'ajoute normalement à la suite).
  var insertionRow = nextFreeRow;
  for (var k = 0; k < nextFreeRow - RECOLTE_DATA_START_ROW && k < ab.length; k++) {
    var d = ab[k][1];
    if (d instanceof Date && dateValue.getTime() < d.getTime()) { insertionRow = RECOLTE_DATA_START_ROW + k; break; }
  }
  if (insertionRow < nextFreeRow) {
    shiftBlockRowsDown(sheet, 1, RECOLTE_NB_COLS, insertionRow, nextFreeRow);
  }

  var r = insertionRow;
  sheet.getRange(r, RECOLTE_COL_NO).setValue(maxNo + 1);
  sheet.getRange(r, RECOLTE_COL_DATE).setValue(dateValue).setNumberFormat(FORMAT_DATE);
  sheet.getRange(r, RECOLTE_COL_PARCELLE).setValue(data.parcelle);
  sheet.getRange(r, RECOLTE_COL_POIDS).setNumberFormat("@").setValue(data.poids); // texte (« <10 g »)
  sheet.getRange(r, RECOLTE_COL_COMMENT).setValue((data.commentaire || "").toString().trim());

  var lat = Number(data.lat), lon = Number(data.lon);
  var cellLoc = sheet.getRange(r, RECOLTE_COL_LOC);
  if (data.lat !== null && data.lat !== undefined && data.lat !== "" && isFinite(lat) && isFinite(lon)) {
    var txt = lat.toFixed(6) + ", " + lon.toFixed(6);
    var url = "https://www.google.com/maps/search/?api=1&query=" + lat.toFixed(6) + "," + lon.toFixed(6);
    // Texte « lat, lon » cliquable : ouvre Google Maps avec une épingle sur la truffe.
    cellLoc.setRichTextValue(SpreadsheetApp.newRichTextValue().setText(txt).setLinkUrl(url).build());
    sheet.getRange(r, RECOLTE_COL_LAT).setNumberFormat("@").setValue(lat.toFixed(6));
    sheet.getRange(r, RECOLTE_COL_LON).setNumberFormat("@").setValue(lon.toFixed(6));
  } else {
    cellLoc.setValue("");
    sheet.getRange(r, RECOLTE_COL_LAT).setValue("");
    sheet.getRange(r, RECOLTE_COL_LON).setValue("");
  }

  // En-têtes des colonnes techniques Latitude / Longitude (créés une seule fois, même style que H2)
  var hLat = sheet.getRange(2, RECOLTE_COL_LAT);
  if (!hLat.getValue()) {
    sheet.getRange(2, RECOLTE_COL_COMMENT).copyTo(sheet.getRange(2, RECOLTE_COL_LAT, 1, 2), { formatOnly: true });
    hLat.setValue("Latitude");
    sheet.getRange(2, RECOLTE_COL_LON).setValue("Longitude");
  }

  return jsonOutput({ ok: true, row: r, no: maxNo + 1 });
}

// Pour l'écran de contrôle : nombre total de truffes + les 5 dernières lignes du tableau
// (la date étant celle du jour, ce sont les plus récentes). Le total sert aussi à l'appli
// pour CONFIRMER qu'une saisie vient bien d'arriver dans le Sheet.
function obtenirDernieresRecoltes() {
  var sheet = ongletPointage(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet) return { total: 0, rows: [] };
  var vals = lireBlocBorne(sheet, RECOLTE_DATA_START_ROW, 1, RECOLTE_COL_COMMENT, 3000);
  var lignes = [];
  for (var i = 0; i < vals.length; i++) {
    if (vals[i][1] === "" || vals[i][1] === null) break;
    lignes.push({
      no: vals[i][0],
      date: vals[i][1] instanceof Date ? formatDateFR(vals[i][1]) : vals[i][1],
      parcelle: vals[i][2] || "",
      poids: vals[i][3] || "",
      gps: !!vals[i][6],
      commentaire: vals[i][7] || ""
    });
  }
  return { total: lignes.length, rows: lignes.slice(-5).reverse() };
}

// Tous les points géolocalisés, pour la carte de l'appli (rafraîchie en direct).
function obtenirPointsRecolte() {
  var sheet = ongletPointage(SpreadsheetApp.getActiveSpreadsheet());
  if (!sheet) return [];
  var vals = lireBlocBorne(sheet, RECOLTE_DATA_START_ROW, 1, RECOLTE_NB_COLS, 3000);
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var row = vals[i];
    if (row[1] === "" || row[1] === null) break;
    var lat = parseFloat(row[RECOLTE_COL_LAT - 1]), lon = parseFloat(row[RECOLTE_COL_LON - 1]);
    if (!isFinite(lat) || !isFinite(lon)) continue; // truffe sans position : pas de point
    out.push({
      no: row[0],
      date: row[1] instanceof Date ? formatDateFR(row[1]) : row[1],
      parcelle: row[2] || "",
      poids: row[3] || "",
      commentaire: row[RECOLTE_COL_COMMENT - 1] || "",
      lat: lat,
      lon: lon
    });
  }
  return out;
}

/* ---------- Poids après récolte : onglet "Récolte volume" ---------- */
// Mise en page : une ligne par DATE ; colonne « Date » ; puis, pour chaque parcelle, un groupe de
// colonnes [poids brut non lavé | poids lavé non trié | poids net vendable | commentaire] sous le
// nom de la parcelle (cellules fusionnées) ; puis, à droite, les colonnes de total par ligne :
// « Total poids brut », « Total lavé non trié », « Total poids net vendable » et « % de perte »
// (= 1 − net vendable / lavé non trié). Au-dessus des en-têtes (ligne 1), les TOTAUX / MOYENNE de
// chaque colonne de total, sur TOUTE la colonne (de la 1re saisie jusqu'en bas, nouvelles saisies
// comprises). Les colonnes sont repérées d'après les EN-TÊTES, pas d'après leur position.
// « verifierOngletVolume » montre comment les en-têtes sont lus (rien n'est écrit).
var PARCELLES_VOLUME = [
  { key: "fontaine", label: "La Fontaine", re: /fontaine/ },
  { key: "baume", label: "La Baume", re: /baume/ },
  { key: "bousquet", label: "Bousquet", re: /bousquet/ },
  { key: "riou", label: "Riou", re: /riou/ },
  { key: "pilles", label: "Les Pilles", re: /pilles/ }
];
var TYPES_POIDS = ["brut", "lave", "net"];
var LIBELLES_TYPES = { brut: "brut non lavé", lave: "lavé non trié", net: "net vendable" };
var NB_MAX_LIGNES = 3000;

function cleParcelleTexte(texte) {
  var t = normTxt(texte), trouvees = [];
  PARCELLES_VOLUME.forEach(function (p) { if (p.re.test(t)) trouvees.push(p.key); });
  return trouvees.length === 1 ? trouvees[0] : null; // 0 ou 2 parcelles dans le texte : non exploitable
}
function cleParcelleLabel(label) {
  var k = null;
  PARCELLES_VOLUME.forEach(function (p) { if (p.label === label) k = p.key; });
  return k;
}
// Type de poids d'après un en-tête (déjà passé par normTxt : minuscules, sans accents).
// L'ordre compte : « net vendable (lavé et trié) » contient « lavé », « brut non lavé » aussi.
function sousTypeTexte(t) {
  if (/\bnet\b|vendable/.test(t)) return "net";
  if (/\bbrut\b/.test(t)) return "brut";
  if (/\blave\b/.test(t)) return "lave";
  return null;
}
function jourStr(d, tz) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }

// Lit les en-têtes de l'onglet et renvoie { dateCol, startRow, groups, totals, perteCol, totalsRow, ... }
function analyserOngletVolume(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var nbRows = Math.min(10, Math.max(sheet.getMaxRows(), 1));
  var vals = sheet.getRange(1, 1, nbRows, lastCol).getValues();
  // on ne garde que le TEXTE des en-têtes (les totaux chiffrés de la ligne 1 n'en font pas partie)
  function txt(r, c) { return typeof vals[r][c] === "string" ? normTxt(vals[r][c]) : ""; }

  var dateCol = 0, dateRow = 0, r, c;
  for (r = 0; r < nbRows && !dateCol; r++) {
    for (c = 0; c < lastCol && !dateCol; c++) {
      if (/^date/.test(txt(r, c))) { dateCol = c + 1; dateRow = r + 1; }
    }
  }
  if (!dateCol) return { error: "Onglet '" + VOLUME_SHEET + "' : en-tête « Date » introuvable dans les 10 premières lignes." };

  // dernière ligne d'en-têtes = dernière ligne (avant les données) qui contient un intitulé
  var lastHeaderRow = dateRow;
  for (r = dateRow; r < nbRows; r++) {
    if (vals[r][dateCol - 1] instanceof Date) break;
    var aIntitule = false;
    for (c = 0; c < lastCol; c++) {
      var t0 = txt(r, c);
      if (t0 && (/brut|\bnet\b|vendable|poids|total|commentaire|date|lave/.test(t0) || cleParcelleTexte(t0))) aIntitule = true;
    }
    if (aIntitule) lastHeaderRow = r + 1;
  }

  // groupe (parcelle) de chaque colonne : le nom est au-dessus, sur la 1re colonne du groupe
  // (cellule fusionnée : les colonnes suivantes sont vides et héritent du groupe)
  var groupAt = [], topGroupRow = 0;
  for (c = 0; c < lastCol; c++) groupAt[c] = null;
  for (r = 0; r < lastHeaderRow; r++) {
    var cur = null;
    for (c = 0; c < lastCol; c++) {
      var t = txt(r, c);
      if (t) {
        cur = /total/.test(t) ? null : cleParcelleTexte(t);
        if (cur && (!topGroupRow || r + 1 < topGroupRow)) topGroupRow = r + 1;
      }
      if (cur && !groupAt[c]) groupAt[c] = cur;
    }
  }

  var groups = {}, totals = {}, units = {}, commentCols = {}, commentCol = 0, perteCol = 0;
  var dernierCol = dateCol, ignorees = [];
  for (c = 0; c < lastCol; c++) {
    var label = "";
    for (r = 0; r < lastHeaderRow; r++) label += " " + txt(r, c);
    label = label.trim();
    if (!label) continue;
    units[c + 1] = /\bkg\b|kilo/.test(label) ? "kg" : "g";
    if (/commentaire|remarque/.test(label)) {
      if (groupAt[c]) commentCols[groupAt[c]] = c + 1; else if (!commentCol) commentCol = c + 1;
      dernierCol = Math.max(dernierCol, c + 1);
      continue;
    }
    if (/perte/.test(label)) { if (!perteCol) perteCol = c + 1; dernierCol = Math.max(dernierCol, c + 1); continue; }
    var sub = sousTypeTexte(label);
    if (/total/.test(label)) {
      if (sub && !totals[sub]) { totals[sub] = c + 1; dernierCol = Math.max(dernierCol, c + 1); }
      continue;
    }
    var g = groupAt[c];
    if (!g || !sub) continue;
    dernierCol = Math.max(dernierCol, c + 1);
    groups[g] = groups[g] || {};
    if (!groups[g][sub]) { groups[g][sub] = c + 1; continue; }
    // deux colonnes du même type pour une parcelle (ex. « poids brut » et « poids brut non lavé ») :
    // on garde celle qui dit « non lavé » ; l'autre est ignorée (et signalée par verifierOngletVolume).
    var anc = groups[g][sub], ancLabel = "";
    for (r = 0; r < lastHeaderRow; r++) ancLabel += " " + txt(r, anc - 1);
    if (/non lave/.test(label) && !/non lave/.test(ancLabel)) { groups[g][sub] = c + 1; ignorees.push({ parcelle: g, col: anc }); }
    else ignorees.push({ parcelle: g, col: c + 1 });
  }

  var ordre = Object.keys(groups).sort(function (a, b) {
    return Math.min.apply(null, TYPES_POIDS.map(function (t) { return groups[a][t] || 9999; })) -
           Math.min.apply(null, TYPES_POIDS.map(function (t) { return groups[b][t] || 9999; }));
  });
  return {
    dateCol: dateCol, headerRow: lastHeaderRow, startRow: lastHeaderRow + 1, groups: groups, totals: totals,
    perteCol: perteCol, totalsRow: topGroupRow > 1 ? topGroupRow - 1 : 0, ignorees: ignorees,
    units: units, commentCols: commentCols, commentCol: commentCol, dernierCol: dernierCol, ordre: ordre
  };
}

// Onglet + analyse. Si l'onglet n'existe pas et creer = true, il est créé avec la même mise en
// page que celle décrite ci-dessus.
function infosOngletVolume(creer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = trouverOngletSouple(ss, VOLUME_SHEET);
  if (!sheet) {
    if (!creer) return null;
    sheet = ss.insertSheet(VOLUME_SHEET);
    var ligne2 = [""], ligne3 = ["Date"];
    PARCELLES_VOLUME.forEach(function (p) {
      ligne2.push(p.label, "", "", "");
      ligne3.push("poids brut non lavé", "Poids lavé non trié", "Poids net vendable (lavé et trié)", "Commentaire");
    });
    ligne2.push("Total poids brut en g", "Total lavé non trié en g", "Total poids net vendable en g", "% de perte");
    ligne3.push("", "", "", "");
    sheet.getRange(2, 1, 1, ligne2.length).setValues([ligne2]);
    sheet.getRange(3, 1, 1, ligne3.length).setValues([ligne3]);
    sheet.getRange(2, 1, 2, ligne2.length).setBackground("#D9D9D9").setFontWeight("bold");
  }
  var info = analyserOngletVolume(sheet);
  if (info.error) return creer ? info : null;
  info.sheet = sheet;
  info.tz = ss.getSpreadsheetTimeZone();
  return info;
}

function colsDuType(info, type) {
  var cols = [];
  info.ordre.forEach(function (g) { if (info.groups[g] && info.groups[g][type]) cols.push(info.groups[g][type]); });
  return cols;
}

// Formules attendues pour la ligne `row` : total brut / lavé / net = somme des cellules de toutes les
// parcelles ; % de perte = 1 − net / lavé (vide tant que l'un des deux est à 0).
function formulesLigneVolume(info, row) {
  var f = {};
  TYPES_POIDS.forEach(function (type) {
    if (!info.totals[type]) return;
    var refs = colsDuType(info, type).map(function (c) { return columnToLetter(c) + row; });
    if (refs.length) f[info.totals[type]] = "=" + refs.join("+");
  });
  if (info.perteCol && info.totals.lave && info.totals.net) {
    var w = columnToLetter(info.totals.lave) + row, x = columnToLetter(info.totals.net) + row;
    f[info.perteCol] = "=IF(AND(" + w + ">0;" + x + ">0);1-(" + x + "/" + w + ");\"\")";
  }
  return f;
}

// Remet d'aplomb TOUTES les formules de total, sans jamais écraser un total saisi à la main :
//  - lignes de données (et lignes pré-remplies en dessous) : total brut / lavé / net / % de perte ;
//  - ligne du haut : somme (ou moyenne pour le % de perte) de TOUTE la colonne, de la 1re ligne de
//    données jusqu'en bas (« V4:V »), donc nouvelles saisies comprises.
function assurerTotauxVolume(info) {
  var sheet = info.sheet, start = info.startRow;
  var tcols = Object.keys(info.totals).map(function (t) { return info.totals[t]; });
  if (info.perteCol) tcols.push(info.perteCol);
  if (!tcols.length) return;
  var nb = Math.max(1, Math.min(NB_MAX_LIGNES, sheet.getMaxRows() - start + 1));
  var rng = sheet.getRange(start, 1, nb, info.dernierCol);
  var vals = rng.getValues(), forms = rng.getFormulas();
  var last = -1, i;
  for (i = 0; i < nb; i++) {
    var d = vals[i][info.dateCol - 1];
    var hasF = tcols.some(function (c) { return forms[i][c - 1] !== ""; });
    if ((d !== "" && d !== null) || hasF) last = i;
  }
  for (i = 0; i <= last; i++) {
    var row = start + i, attendu = formulesLigneVolume(info, row);
    Object.keys(attendu).forEach(function (cs) {
      var c = Number(cs), actuel = forms[i][c - 1], valeur = vals[i][c - 1];
      if (actuel === attendu[cs]) return;
      if (actuel === "" && valeur !== "" && valeur !== null) {
        // total tapé à la main : on le garde tant qu'aucune parcelle n'a de poids de ce type
        var type = null;
        TYPES_POIDS.forEach(function (t) { if (info.totals[t] === c) type = t; });
        if (type) {
          var aSaisie = colsDuType(info, type).some(function (cc) { return vals[i][cc - 1] !== "" && vals[i][cc - 1] !== null; });
          if (!aSaisie) return;
        }
      }
      sheet.getRange(row, c).setFormula(attendu[cs]);
    });
  }
  if (info.totalsRow) {
    var tr = info.totalsRow;
    TYPES_POIDS.forEach(function (t) {
      if (!info.totals[t]) return;
      var L = columnToLetter(info.totals[t]), f = "=SUM(" + L + start + ":" + L + ")";
      var cell = sheet.getRange(tr, info.totals[t]);
      if (cell.getFormula() !== f) cell.setFormula(f);
    });
    if (info.perteCol) {
      var P = columnToLetter(info.perteCol), fp = "=IFERROR(AVERAGE(" + P + start + ":" + P + ");\"\")";
      var cp = sheet.getRange(tr, info.perteCol);
      if (cp.getFormula() !== fp) cp.setFormula(fp);
    }
  }
}

function handleRecolteVolume(data) {
  var info = infosOngletVolume(true);
  if (info.error) return jsonOutput({ ok: false, error: info.error });

  var key = cleParcelleLabel(data.parcelle);
  if (!key) return jsonOutput({ ok: false, error: "Parcelle inconnue : " + data.parcelle });
  if (TYPES_POIDS.indexOf(data.type) === -1) return jsonOutput({ ok: false, error: "Type de poids inconnu : " + data.type });
  var grp = info.groups[key];
  if (!grp || !grp[data.type]) {
    return jsonOutput({ ok: false, error: "Colonne « poids " + LIBELLES_TYPES[data.type] + " » introuvable pour " + data.parcelle + " (voir verifierOngletVolume)." });
  }
  var grammes = Number(data.poids);
  if (!isFinite(grammes) || grammes <= 0) return jsonOutput({ ok: false, error: "Poids invalide." });
  if (!data.date) return jsonOutput({ ok: false, error: "Date manquante." });

  var sheet = info.sheet, tz = info.tz, start = info.startRow;
  var jour = jourStr(parseDateOnly(data.date), tz);
  var bloc = lireBlocBorne(sheet, start, 1, info.dernierCol, NB_MAX_LIGNES);

  // Une ligne par date : on cherche d'abord la ligne de ce jour ; sinon la place chronologique.
  var rowSame = -1, insertion = -1, nextFree = -1;
  for (var i = 0; i < bloc.length; i++) {
    var d = bloc[i][info.dateCol - 1];
    if (d === "" || d === null) { nextFree = start + i; break; }
    if (d instanceof Date) {
      var j = jourStr(d, tz);
      if (j === jour && rowSame === -1) rowSame = start + i;
      if (j > jour && insertion === -1) insertion = start + i;
    }
  }
  var target;
  if (rowSame !== -1) {
    target = rowSame;
  } else {
    if (nextFree === -1) {
      if (bloc.length < NB_MAX_LIGNES) { sheet.insertRowsAfter(sheet.getMaxRows(), 200); nextFree = start + bloc.length; }
      else return jsonOutput({ ok: false, error: "Tableau plein." });
    }
    target = insertion !== -1 ? insertion : nextFree;
    if (insertion !== -1) shiftBlockRowsDown(sheet, 1, info.dernierCol, insertion, nextFree);
    sheet.getRange(target, info.dateCol).setValue(parseDateOnly(data.date)).setNumberFormat(FORMAT_DATE);
  }

  // Poids : plusieurs saisies de la même parcelle / même type le même jour se CUMULENT.
  var col = grp[data.type], unite = info.units[col] || "g";
  var cellule = sheet.getRange(target, col);
  var avant = cellule.getValue();
  var ajout = unite === "kg" ? grammes / 1000 : grammes;
  var cumul = (typeof avant === "number" ? avant : 0) + ajout;
  cellule.setValue(unite === "kg" ? Math.round(cumul * 1000) / 1000 : Math.round(cumul));

  // Commentaire : colonne « Commentaire » de CETTE parcelle (sinon note sur la cellule du poids)
  var commentaire = (data.commentaire || "").toString().trim();
  if (commentaire) {
    var colCom = info.commentCols[key] || info.commentCol;
    if (colCom) {
      var cc = sheet.getRange(target, colCom), ancien = (cc.getValue() || "").toString();
      cc.setValue(ancien ? ancien + " ; " + commentaire : commentaire);
    } else {
      var ancienneNote = cellule.getNote() || "";
      cellule.setNote((ancienneNote ? ancienneNote + "\n" : "") + commentaire);
    }
  }

  // Totaux : formules de la ligne saisie, des lignes décalées, et de la ligne du haut (toute la colonne)
  assurerTotauxVolume(info);
  SpreadsheetApp.flush();

  return jsonOutput({ ok: true, row: target, cumul: cumul });
}

function nombreOuNull(v) { return (typeof v === "number" && isFinite(v)) ? v : null; }

// Lit tout le tableau et le met en forme pour l'appli (poids toujours en grammes).
function lireTableauPoids(info) {
  var sheet = info.sheet;
  SpreadsheetApp.flush();
  var vals = lireBlocBorne(sheet, info.startRow, 1, info.dernierCol, NB_MAX_LIGNES);
  function enG(v, col) {
    var n = Number(v);
    if (!isFinite(n) || v === "" || v === null) return 0;
    return (info.units[col] === "kg") ? Math.round(n * 1000) : n;
  }
  var parcelles = info.ordre.map(function (g) {
    var label = g;
    PARCELLES_VOLUME.forEach(function (p) { if (p.key === g) label = p.label; });
    return { key: g, label: label };
  });
  var rows = [], totaux = {}, somme = { brut: 0, lave: 0, net: 0 }, signature = 0;
  parcelles.forEach(function (p) { totaux[p.key] = { brut: 0, lave: 0, net: 0, pLave: 0, pNet: 0 }; });

  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][info.dateCol - 1];
    if (d === "" || d === null) break;
    var ligne = { date: d instanceof Date ? formatDateFR(d) : d, v: {}, totalBrut: 0, totalLave: 0, totalNet: 0 };
    var calc = { brut: 0, lave: 0, net: 0 };
    parcelles.forEach(function (p) {
      var grp = info.groups[p.key], o = { brut: 0, lave: 0, net: 0, commentaire: "" };
      TYPES_POIDS.forEach(function (t) {
        var g = grp[t] ? enG(vals[i][grp[t] - 1], grp[t]) : 0;
        o[t] = g;
        totaux[p.key][t] += g;
        calc[t] += g;
        signature += g;
      });
      var colCom = info.commentCols[p.key];
      if (colCom) o.commentaire = (vals[i][colCom - 1] || "").toString();
      // pour la perte par parcelle : seulement les dates où le lavé ET le net vendable sont renseignés
      if (o.lave > 0 && o.net > 0) { totaux[p.key].pLave += o.lave; totaux[p.key].pNet += o.net; }
      if (o.brut || o.lave || o.net || o.commentaire) ligne.v[p.key] = o;
    });
    // total de la ligne : colonne « Total … » du Sheet (y compris un total tapé à la main), sinon somme
    var tb = info.totals.brut ? nombreOuNull(vals[i][info.totals.brut - 1]) : null;
    var tl = info.totals.lave ? nombreOuNull(vals[i][info.totals.lave - 1]) : null;
    var tn = info.totals.net ? nombreOuNull(vals[i][info.totals.net - 1]) : null;
    ligne.totalBrut = tb !== null ? tb : calc.brut;
    ligne.totalLave = tl !== null ? tl : calc.lave;
    ligne.totalNet = tn !== null ? tn : calc.net;
    ligne.perte = (ligne.totalLave > 0 && ligne.totalNet > 0) ? 1 - ligne.totalNet / ligne.totalLave : null;
    somme.brut += ligne.totalBrut; somme.lave += ligne.totalLave; somme.net += ligne.totalNet;
    rows.push(ligne);
  }

  // Totaux / moyenne de la ligne du haut du Sheet (ce que l'utilisateur voit en haut des colonnes)
  var haut = { brut: null, lave: null, net: null, perte: null };
  if (info.totalsRow) {
    TYPES_POIDS.forEach(function (t) {
      if (info.totals[t]) haut[t] = nombreOuNull(sheet.getRange(info.totalsRow, info.totals[t]).getValue());
    });
    if (info.perteCol) haut.perte = nombreOuNull(sheet.getRange(info.totalsRow, info.perteCol).getValue());
  }
  var totalBrut = haut.brut !== null ? haut.brut : somme.brut;
  var totalLave = haut.lave !== null ? haut.lave : somme.lave;
  var totalNet = haut.net !== null ? haut.net : somme.net;
  var perte = haut.perte !== null ? haut.perte : ((totalLave > 0 && totalNet > 0) ? 1 - totalNet / totalLave : null);
  return {
    unite: "g", parcelles: parcelles, rows: rows, totaux: totaux,
    totalBrut: totalBrut, totalLave: totalLave, totalNet: totalNet, perte: perte, signature: signature
  };
}

// Écran de contrôle : les 5 dernières dates + une « signature » (somme de tous les poids) que
// l'appli compare avant / après pour CONFIRMER qu'une saisie est bien arrivée dans le Sheet.
function obtenirDernieresPesees() {
  var info = infosOngletVolume(false);
  if (!info) return { signature: 0, parcelles: [], rows: [] };
  var t = lireTableauPoids(info);
  return { signature: t.signature, parcelles: t.parcelles, rows: t.rows.slice(-5).reverse() };
}

// Tableau complet pour l'écran de consultation.
function obtenirTableauPoids() {
  var info = infosOngletVolume(false);
  if (!info) return { unite: "g", parcelles: [], rows: [], totaux: {}, totalBrut: 0, totalLave: 0, totalNet: 0, perte: null, signature: 0 };
  var t = lireTableauPoids(info);
  t.rows = t.rows.slice().reverse(); // la plus récente en premier
  return t;
}

// À exécuter pour vérifier, SANS rien écrire, comment les en-têtes de « Récolte volume » sont
// interprétés (colonnes de chaque parcelle, colonnes de total, unités).
function verifierOngletVolume() {
  var info = infosOngletVolume(false);
  var lignes = [];
  if (!info) {
    lignes.push("Onglet '" + VOLUME_SHEET + "' introuvable (il sera créé à la première saisie) ou en-têtes non reconnus.");
    var sh = trouverOngletSouple(SpreadsheetApp.getActiveSpreadsheet(), VOLUME_SHEET);
    if (sh) { var a = analyserOngletVolume(sh); if (a.error) lignes.push("Détail : " + a.error); }
  } else {
    lignes.push("Colonne Date : " + columnToLetter(info.dateCol) + " — données à partir de la ligne " + info.startRow);
    PARCELLES_VOLUME.forEach(function (p) {
      var g = info.groups[p.key];
      if (!g) { lignes.push("⚠ " + p.label + " : AUCUN groupe de colonnes reconnu"); return; }
      lignes.push(p.label + " : " + TYPES_POIDS.map(function (t) {
        return LIBELLES_TYPES[t] + "=" + (g[t] ? columnToLetter(g[t]) : "—");
      }).join(" · ") + " · commentaire=" + (info.commentCols[p.key] ? columnToLetter(info.commentCols[p.key]) : "— (⚠ aucune colonne « Commentaire » : le commentaire sera mis en note sur la cellule du poids)"));
    });
    info.ignorees.forEach(function (x) {
      lignes.push("⚠ Colonne " + columnToLetter(x.col) + " (" + x.parcelle + ") en double d'un autre poids de même type : IGNORÉE par l'appli et par les totaux. Renomme-la (ex. « Commentaire ») ou supprime-la.");
    });
    TYPES_POIDS.forEach(function (t) {
      lignes.push("Total " + LIBELLES_TYPES[t] + " : " + (info.totals[t] ? "colonne " + columnToLetter(info.totals[t]) : "⚠ colonne NON reconnue (en-tête « Total … »)"));
    });
    lignes.push("% de perte : " + (info.perteCol ? "colonne " + columnToLetter(info.perteCol) + " (= 1 − net vendable / lavé non trié)" : "⚠ non reconnue"));
    lignes.push("Ligne des totaux / moyenne : " + (info.totalsRow ? "ligne " + info.totalsRow + " (formules sur toute la colonne, à partir de la ligne " + info.startRow + ")" : "⚠ non trouvée (elle doit être au-dessus de la ligne des noms de parcelles)"));
  }
  Logger.log(lignes.join("\n"));
  return lignes;
}

// À exécuter UNE fois (facultatif : l'appli le fait aussi à chaque saisie) : réécrit les formules
// de total de toutes les lignes et de la ligne du haut dans « Récolte volume ».
function reparerTotauxVolume() {
  var info = infosOngletVolume(false);
  if (!info) { Logger.log("Onglet introuvable."); return; }
  assurerTotauxVolume(info);
  Logger.log("Totaux de « Récolte volume » réécrits.");
}

/* ---------- Ventes : onglet "Vente" ---------- */
// Mise en page : une ligne par DATE (ordre chronologique). Pour chaque qualité (Brisure, Morceaux,
// 2e choix, 1er choix, Extra) : 2 colonnes [Prix/Kg | Poids g]. Puis 3 colonnes calculées par
// date : « Total vente » (€), « Prix de vente moyen » (€/kg), « Total poids vendu » (g). Au-dessus
// des en-têtes, le tableau des TOTAUX / MOYENNE (en rouge) de chaque colonne, sur toute la colonne.
// Plusieurs ventes de la même qualité le même jour : poids additionnés, prix = moyenne pondérée.
var VENTE_SHEET = "Vente";
var QUALITES_VENTE = [
  { key: "brisure", label: "Brisure", re: /brisure/ },
  { key: "morceaux", label: "Morceaux", re: /morceau/ },
  { key: "choix2", label: "2e choix", re: /\b2\s*(e|eme|nd|nde)?\s*choix|deuxieme/ },
  { key: "choix1", label: "1er choix", re: /\b1\s*(er|ere|ier)?\s*choix|premier/ },
  { key: "extra", label: "Extra", re: /extra/ }
];

function cleQualiteTexte(t) {
  var trouvees = [];
  QUALITES_VENTE.forEach(function (q) { if (q.re.test(t)) trouvees.push(q.key); });
  return trouvees.length === 1 ? trouvees[0] : null;
}
function cleQualiteLabel(label) {
  var k = null;
  QUALITES_VENTE.forEach(function (q) { if (q.label === label || q.key === label) k = q.key; });
  return k;
}

function analyserOngletVente(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var nbRows = Math.min(12, Math.max(sheet.getMaxRows(), 1));
  var vals = sheet.getRange(1, 1, nbRows, lastCol).getValues();
  function txt(r, c) { return typeof vals[r][c] === "string" ? normTxt(vals[r][c]) : ""; }
  var r, c;

  var qualCols = {}, groupRow = 0;
  for (r = 0; r < nbRows; r++) {
    for (c = 1; c < lastCol; c++) {
      var k = cleQualiteTexte(txt(r, c));
      if (k && !qualCols[k]) { qualCols[k] = c + 1; if (!groupRow) groupRow = r + 1; }
    }
  }
  var manque = QUALITES_VENTE.filter(function (q) { return !qualCols[q.key]; }).map(function (q) { return q.label; });
  if (manque.length) return { error: "Onglet '" + VENTE_SHEET + "' : qualité(s) introuvable(s) dans les en-têtes : " + manque.join(", ") + "." };

  var headerRow = 0;
  for (r = groupRow; r < nbRows && !headerRow; r++) {
    for (c = 0; c < lastCol; c++) if (/^prix/.test(txt(r, c))) { headerRow = r + 1; break; }
  }
  if (!headerRow) return { error: "Onglet '" + VENTE_SHEET + "' : ligne d'en-têtes « Prix/Kg » / « Poids » introuvable." };

  var cols = {}, dernierQ = 0;
  for (var qi = 0; qi < QUALITES_VENTE.length; qi++) {
    var key = QUALITES_VENTE[qi].key, pc = qualCols[key];
    if (!/^prix/.test(txt(headerRow - 1, pc - 1)) || !/^poids/.test(txt(headerRow - 1, pc))) {
      return { error: "Onglet '" + VENTE_SHEET + "' : sous « " + QUALITES_VENTE[qi].label + " » les colonnes doivent s'appeler « Prix… » puis « Poids… » (ligne " + headerRow + ")." };
    }
    cols[key] = { prix: pc, poids: pc + 1 };
    dernierQ = Math.max(dernierQ, pc + 1);
  }

  var colEuros = 0, colMoy = 0, colG = 0, clientCol = 0, listCol = 0;
  for (c = dernierQ; c < lastCol; c++) {
    var label = "";
    for (r = 0; r < headerRow; r++) label += " " + txt(r, c);
    label = label.trim();
    if (!label) continue;
    if (/client/.test(label)) { if (/liste/.test(label)) { if (!listCol) listCol = c + 1; } else if (!clientCol) clientCol = c + 1; }
    else if (/moyen/.test(label)) { if (!colMoy) colMoy = c + 1; }
    else if (/total/.test(label) && /vente|€|eur/.test(label)) { if (!colEuros) colEuros = c + 1; }
    else if (/total/.test(label) && /\bg\b|poids|vendu/.test(label)) { if (!colG) colG = c + 1; }
  }
  if (!colEuros || !colMoy || !colG) {
    return { error: "Onglet '" + VENTE_SHEET + "' : colonnes « Total vente (€) », « Prix de vente moyen » et « Total poids vendu » non reconnues." };
  }

  // ligne du tableau des totaux : la ligne (au-dessus des en-têtes) qui contient des formules
  var totalsRow = 0;
  var probe = cols.brisure.poids;
  for (r = groupRow + 1; r < headerRow && !totalsRow; r++) {
    if (sheet.getRange(r, probe).getFormula() !== "") totalsRow = r;
  }
  if (!totalsRow) totalsRow = Math.max(headerRow - 2, groupRow + 1);

  return {
    dateCol: 1, groupRow: groupRow, headerRow: headerRow, startRow: headerRow + 1, totalsRow: totalsRow,
    cols: cols, colEuros: colEuros, colMoy: colMoy, colG: colG, clientCol: clientCol, listCol: listCol,
    dernierCol: Math.max(colEuros, colMoy, colG, dernierQ, clientCol)
  };
}

function infosOngletVente(creer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = trouverOngletSouple(ss, VENTE_SHEET);
  if (!sheet) {
    if (!creer) return null;
    sheet = ss.insertSheet(VENTE_SHEET);
    var l1 = [""], l2 = [""], l5 = [""];
    QUALITES_VENTE.forEach(function (q) { l1.push(q.label, ""); l2.push("moyenne", "total"); l5.push("Prix/Kg", "Poids g"); });
    l1.push("", "", ""); l2.push("total €", "moyenne €", "total g"); l5.push("Total vente", "Prix de vente moyen", "Total poids vendu");
    sheet.getRange(1, 1, 1, l1.length).setValues([l1]);
    sheet.getRange(2, 1, 1, l2.length).setValues([l2]);
    sheet.getRange(5, 1, 1, l5.length).setValues([l5]);
    sheet.getRange(1, 1, 3, l1.length).setFontColor("#980000").setFontWeight("bold");
    sheet.getRange(5, 1, 1, l5.length).setFontWeight("bold");
  }
  var info = analyserOngletVente(sheet);
  if (info.error) return creer ? info : null;
  info.sheet = sheet;
  info.tz = ss.getSpreadsheetTimeZone();
  return info;
}

// Clients : colonne « Client » (à droite de « Total poids vendu ») et liste des noms proposés dans
// l'appli (colonne « Liste clients », éditable à la main : une ligne = un nom). Créées si absentes.
var CLIENTS_PAR_DEFAUT = ["Particulier", "Restaurateur", "Grossiste", "Marché", "Autre professionnel"];

function lireListeClients(info) {
  if (!info.listCol) return [];
  var n = Math.max(1, Math.min(300, info.sheet.getMaxRows() - info.startRow + 1));
  var v = info.sheet.getRange(info.startRow, info.listCol, n, 1).getValues(), out = [];
  for (var i = 0; i < v.length; i++) { var t = (v[i][0] || "").toString().trim(); if (t) out.push(t); }
  return out;
}

function assurerColonnesClient(info) {
  var sheet = info.sheet, modifie = false;
  function colLibre(depuis) {
    var col = depuis;
    while ((sheet.getRange(info.headerRow, col).getValue() || "") !== "") col++;
    return col;
  }
  if (!info.clientCol) {
    var cc = colLibre(Math.max(info.colEuros, info.colMoy, info.colG) + 1);
    sheet.getRange(info.headerRow, cc).setValue("Client").setFontWeight("bold");
    info.clientCol = cc; modifie = true;
  }
  if (!info.listCol) {
    var lc = colLibre(info.clientCol + 2);
    sheet.getRange(info.headerRow, lc).setValue("Liste clients").setFontWeight("bold");
    info.listCol = lc; modifie = true;
  }
  if (!lireListeClients({ sheet: sheet, startRow: info.startRow, listCol: info.listCol }).length) {
    sheet.getRange(info.startRow, info.listCol, CLIENTS_PAR_DEFAUT.length, 1)
      .setValues(CLIENTS_PAR_DEFAUT.map(function (x) { return [x]; }));
  }
  if (!modifie) return info;
  var neuf = analyserOngletVente(sheet);
  if (neuf.error) return neuf;
  neuf.sheet = sheet; neuf.tz = info.tz;
  return neuf;
}

function ajouterClientSiNouveau(info, nom) {
  if (!nom) return;
  var liste = lireListeClients(info), k = normTxt(nom);
  for (var i = 0; i < liste.length; i++) if (normTxt(liste[i]) === k) return;
  var n = Math.max(1, Math.min(300, info.sheet.getMaxRows() - info.startRow + 1));
  var v = info.sheet.getRange(info.startRow, info.listCol, n, 1).getValues();
  for (var j = 0; j < v.length; j++) {
    if (!(v[j][0] || "").toString().trim()) { info.sheet.getRange(info.startRow + j, info.listCol).setValue(nom); return; }
  }
}

function obtenirListesVente() {
  var info = infosOngletVente(true);
  if (info.error) return { clients: CLIENTS_PAR_DEFAUT.slice() };
  info = assurerColonnesClient(info);
  if (info.error) return { clients: CLIENTS_PAR_DEFAUT.slice() };
  return { clients: lireListeClients(info) };
}

// Formules attendues pour une ligne de date : total € = Σ prix × poids / 1000 ; moyenne €/kg ; total g
function formulesLigneVente(info, row) {
  var termes = [], poids = [];
  QUALITES_VENTE.forEach(function (q) {
    var p = columnToLetter(info.cols[q.key].prix) + row, w = columnToLetter(info.cols[q.key].poids) + row;
    termes.push("(" + p + "*" + w + ")");
    poids.push(w);
  });
  var f = {};
  var E = columnToLetter(info.colEuros) + row, G = columnToLetter(info.colG) + row;
  f[info.colEuros] = "=(" + termes.join("+") + ")/1000";
  f[info.colMoy] = "=IF(" + G + ">0;" + E + "/(" + G + "/1000);\"\")";
  f[info.colG] = "=" + poids.join("+");
  return f;
}

// Remet d'aplomb les formules des lignes de date et le tableau des totaux du haut (sur TOUTE la
// colonne, à partir de la 1re ligne de données).
function assurerTotauxVente(info) {
  var sheet = info.sheet, start = info.startRow;
  var nb = Math.max(1, Math.min(NB_MAX_LIGNES, sheet.getMaxRows() - start + 1));
  var rng = sheet.getRange(start, 1, nb, info.dernierCol);
  var vals = rng.getValues(), forms = rng.getFormulas();
  var fcols = [info.colEuros, info.colMoy, info.colG], last = -1, i;
  for (i = 0; i < nb; i++) {
    var d = vals[i][0];
    var hasF = fcols.some(function (c) { return forms[i][c - 1] !== ""; });
    if ((d !== "" && d !== null) || hasF) last = i;
  }
  for (i = 0; i <= last; i++) {
    var row = start + i, attendu = formulesLigneVente(info, row);
    Object.keys(attendu).forEach(function (cs) {
      var c = Number(cs);
      if (forms[i][c - 1] !== attendu[cs]) sheet.getRange(row, c).setFormula(attendu[cs]);
    });
  }
  var tr = info.totalsRow;
  function poser(col, f) {
    var cell = sheet.getRange(tr, col);
    if (cell.getFormula() !== f) cell.setFormula(f);
  }
  QUALITES_VENTE.forEach(function (q) {
    var P = columnToLetter(info.cols[q.key].prix), W = columnToLetter(info.cols[q.key].poids);
    poser(info.cols[q.key].prix, "=IFERROR(AVERAGE(" + P + start + ":" + P + ");\"\")");
    poser(info.cols[q.key].poids, "=SUM(" + W + start + ":" + W + ")");
  });
  var LE = columnToLetter(info.colEuros), LM = columnToLetter(info.colMoy), LG = columnToLetter(info.colG);
  poser(info.colEuros, "=SUM(" + LE + start + ":" + LE + ")");
  poser(info.colMoy, "=IFERROR(AVERAGE(" + LM + start + ":" + LM + ");\"\")");
  poser(info.colG, "=SUM(" + LG + start + ":" + LG + ")");
}

// data = { date: "AAAA-MM-JJ", lignes: [ { qualite: "1er choix", prix: 350, poids: 291 }, ... ] }
function handleVente(data) {
  var info = infosOngletVente(true);
  if (info.error) return jsonOutput({ ok: false, error: info.error });
  info = assurerColonnesClient(info);
  if (info.error) return jsonOutput({ ok: false, error: info.error });
  var client = (data.client || "").toString().trim();
  if (!data.date) return jsonOutput({ ok: false, error: "Date manquante." });
  var lignes = data.lignes || [];
  if (!lignes.length) return jsonOutput({ ok: false, error: "Aucune qualité renseignée." });
  var propres = [];
  for (var li = 0; li < lignes.length; li++) {
    var q = cleQualiteLabel(lignes[li].qualite);
    var prix = Number(lignes[li].prix), poids = Number(lignes[li].poids);
    if (!q) return jsonOutput({ ok: false, error: "Qualité inconnue : " + lignes[li].qualite });
    if (!isFinite(prix) || prix <= 0) return jsonOutput({ ok: false, error: "Prix invalide pour " + lignes[li].qualite + "." });
    if (!isFinite(poids) || poids <= 0) return jsonOutput({ ok: false, error: "Poids invalide pour " + lignes[li].qualite + "." });
    propres.push({ key: q, prix: prix, poids: poids });
  }

  var sheet = info.sheet, tz = info.tz, start = info.startRow;
  var jour = jourStr(parseDateOnly(data.date), tz);
  var bloc = lireBlocBorne(sheet, start, 1, info.dernierCol, NB_MAX_LIGNES);
  var rowSame = -1, insertion = -1, nextFree = -1;
  for (var i = 0; i < bloc.length; i++) {
    var d = bloc[i][0];
    if (d === "" || d === null) { nextFree = start + i; break; }
    if (d instanceof Date) {
      var j = jourStr(d, tz);
      // une ligne par DATE et par CLIENT : même jour mais autre client = nouvelle ligne
      if (j === jour && rowSame === -1 && normTxt(bloc[i][info.clientCol - 1]) === normTxt(client)) rowSame = start + i;
      if (j > jour && insertion === -1) insertion = start + i;
    }
  }
  var target;
  if (rowSame !== -1) {
    target = rowSame;
  } else {
    if (nextFree === -1) {
      if (bloc.length < NB_MAX_LIGNES) { sheet.insertRowsAfter(sheet.getMaxRows(), 200); nextFree = start + bloc.length; }
      else return jsonOutput({ ok: false, error: "Tableau plein." });
    }
    target = insertion !== -1 ? insertion : nextFree;
    if (insertion !== -1) shiftBlockRowsDown(sheet, 1, info.dernierCol, insertion, nextFree);
    sheet.getRange(target, 1).setValue(parseDateOnly(data.date)).setNumberFormat(FORMAT_DATE);
    if (client) sheet.getRange(target, info.clientCol).setValue(client);
  }
  ajouterClientSiNouveau(info, client);

  propres.forEach(function (p) {
    var cp = sheet.getRange(target, info.cols[p.key].prix), cw = sheet.getRange(target, info.cols[p.key].poids);
    var ancienP = Number(cp.getValue()), ancienW = Number(cw.getValue());
    if (isFinite(ancienW) && ancienW > 0) {
      var w = ancienW + p.poids;
      var moy = ((isFinite(ancienP) ? ancienP : 0) * ancienW + p.prix * p.poids) / w;
      cp.setValue(Math.round(moy * 100) / 100);
      cw.setValue(Math.round(w));
    } else {
      cp.setValue(Math.round(p.prix * 100) / 100);
      cw.setValue(Math.round(p.poids));
    }
  });

  assurerTotauxVente(info);
  sheet.getRange(target, info.colEuros, 1, 2).setNumberFormat("#,##0.00");
  SpreadsheetApp.flush();
  return jsonOutput({ ok: true, row: target });
}

function lireTableauVente(info) {
  var sheet = info.sheet;
  SpreadsheetApp.flush();
  var vals = lireBlocBorne(sheet, info.startRow, 1, info.dernierCol, NB_MAX_LIGNES);
  var qualites = QUALITES_VENTE.map(function (q) { return { key: q.key, label: q.label }; });
  var rows = [], signature = 0, somme = { euros: 0, g: 0 };
  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][0];
    if (d === "" || d === null) break;
    var ligne = { date: d instanceof Date ? formatDateFR(d) : d, client: info.clientCol ? (vals[i][info.clientCol - 1] || "").toString() : "", q: {}, euros: 0, moyenne: null, g: 0 };
    var eurosCalc = 0, gCalc = 0;
    QUALITES_VENTE.forEach(function (q) {
      var p = Number(vals[i][info.cols[q.key].prix - 1]), w = Number(vals[i][info.cols[q.key].poids - 1]);
      if (isFinite(w) && w > 0) {
        var pp = isFinite(p) ? p : 0;
        ligne.q[q.key] = { prix: pp, poids: w };
        eurosCalc += pp * w / 1000; gCalc += w; signature += w;
      }
    });
    var e = nombreOuNull(vals[i][info.colEuros - 1]), m = nombreOuNull(vals[i][info.colMoy - 1]), g = nombreOuNull(vals[i][info.colG - 1]);
    ligne.euros = e !== null ? e : eurosCalc;
    ligne.g = g !== null ? g : gCalc;
    ligne.moyenne = m !== null ? m : (ligne.g > 0 ? ligne.euros / (ligne.g / 1000) : null);
    somme.euros += ligne.euros; somme.g += ligne.g;
    rows.push(ligne);
  }
  // Tableau rouge du haut, tel qu'il est dans le Sheet
  var tr = info.totalsRow, haut = { q: {} };
  QUALITES_VENTE.forEach(function (q) {
    haut.q[q.key] = {
      moyenne: nombreOuNull(sheet.getRange(tr, info.cols[q.key].prix).getValue()),
      total: nombreOuNull(sheet.getRange(tr, info.cols[q.key].poids).getValue())
    };
  });
  haut.euros = nombreOuNull(sheet.getRange(tr, info.colEuros).getValue());
  haut.moyenne = nombreOuNull(sheet.getRange(tr, info.colMoy).getValue());
  haut.g = nombreOuNull(sheet.getRange(tr, info.colG).getValue());
  if (haut.euros === null) haut.euros = somme.euros;
  if (haut.g === null) haut.g = somme.g;
  return { qualites: qualites, totaux: haut, rows: rows, signature: signature };
}

function obtenirDernieresVentes() {
  var info = infosOngletVente(false);
  if (!info) return { signature: 0, qualites: [], rows: [] };
  var t = lireTableauVente(info);
  return { signature: t.signature, qualites: t.qualites, rows: t.rows.slice(-5).reverse() };
}

function obtenirTableauVente() {
  var info = infosOngletVente(false);
  if (!info) return { qualites: [], totaux: { q: {}, euros: 0, moyenne: null, g: 0 }, rows: [], signature: 0 };
  var t = lireTableauVente(info);
  t.rows = t.rows.slice().reverse();
  return t;
}

function verifierOngletVente() {
  var info = infosOngletVente(false);
  var lignes = [];
  if (!info) {
    lignes.push("Onglet '" + VENTE_SHEET + "' introuvable (il sera créé à la première vente) ou en-têtes non reconnus.");
    var sh = trouverOngletSouple(SpreadsheetApp.getActiveSpreadsheet(), VENTE_SHEET);
    if (sh) { var a = analyserOngletVente(sh); if (a.error) lignes.push("Détail : " + a.error); }
  } else {
    lignes.push("Données à partir de la ligne " + info.startRow + " · tableau des totaux (rouge) : ligne " + info.totalsRow);
    QUALITES_VENTE.forEach(function (q) {
      lignes.push(q.label + " : prix/kg=" + columnToLetter(info.cols[q.key].prix) + " · poids=" + columnToLetter(info.cols[q.key].poids));
    });
    lignes.push("Client : " + (info.clientCol ? "colonne " + columnToLetter(info.clientCol) : "(sera créée à la première vente)") + " · liste des clients : " + (info.listCol ? "colonne " + columnToLetter(info.listCol) : "(sera créée)"));
    lignes.push("Total vente € : " + columnToLetter(info.colEuros) + " · prix moyen : " + columnToLetter(info.colMoy) + " · total poids : " + columnToLetter(info.colG));
  }
  Logger.log(lignes.join("\n"));
  return lignes;
}

// Facultatif (l'appli le fait à chaque vente) : réécrit toutes les formules de l'onglet Vente.
function reparerTotauxVente() {
  var info = infosOngletVente(false);
  if (!info) { Logger.log("Onglet introuvable."); return; }
  assurerTotauxVente(info);
  Logger.log("Totaux de « Vente » réécrits.");
}

/* ---------- Nettoyage ponctuel : effacer une ligne de test du journal Heures ---------- */
// Efface uniquement le CONTENU de B3:G3 (Date à Horodatage) — ne supprime ni ne décale aucune
// cellule, donc aucun risque de conflit avec les cellules fusionnées des tableaux
// récapitulatifs qui partagent la ligne 3 plus loin (in-têtes Paul/René/Charles...).
function effacerLigneTestHeures() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) { Logger.log("Onglet 'Heures' introuvable."); return; }
  sheet.getRange(3, HEURES_COL_DATE, 1, 6).clearContent(); // B3:G3
  Logger.log("Contenu de B3:G3 effacé (rien supprimé ni décalé).");
}

function verifierFuseauHoraire() {
  var tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var maintenant = new Date();
  Logger.log("Fuseau horaire actuel du classeur : " + tz);
  Logger.log("Heure 'actuelle' vue par le script (new Date()) : " + maintenant);
  Logger.log("Cette même heure, affichée explicitement en Europe/Paris : " + Utilities.formatDate(maintenant, "Europe/Paris", "dd/MM/yyyy HH:mm:ss"));
  if (tz !== "Europe/Paris") {
    Logger.log("⚠️ Le fuseau du classeur n'est PAS 'Europe/Paris' — c'est très probablement la cause du décalage d'heure observé. Corrige-le dans Fichier > Paramètres du classeur > Fuseau horaire, puis relance cette fonction pour confirmer.");
  } else {
    Logger.log("✓ Le fuseau du classeur est déjà correctement réglé sur Europe/Paris.");
  }
  return tz;
}

/* ---------- Migration : formules des tableaux récapitulatifs de l'onglet Heures ---------- */
// À exécuter une seule fois (et à nouveau après tout ajout manuel d'une nouvelle ligne dans
// les listes Tache/Lieux, si jamais la recopie automatique n'avait pas eu lieu). Pose les
// formules SUMIFS dans les 3 tableaux récapitulatifs déjà présents dans l'onglet "Heures" :
// Total par tâche (H:O), Total par lieu (Q:X), Total par personne et par mois (Z:AM).
function configurerTableauxHeures() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HEURES_SHEET);
  if (!sheet) { Logger.log("Onglet 'Heures' introuvable."); return; }

  var results = [];
  var plageFin = HEURES_DATA_START_ROW + 2000;
  var colDateL = columnToLetter(HEURES_COL_DATE);
  var colPersL = columnToLetter(HEURES_COL_PERSONNE);
  var colLieuL = columnToLetter(HEURES_COL_LIEU);
  var colTacheL = columnToLetter(HEURES_COL_TACHE);
  var colTempsL = columnToLetter(HEURES_COL_TEMPS);
  var refDate = "$" + colDateL + "$" + HEURES_DATA_START_ROW + ":$" + colDateL + "$" + plageFin;
  var refPers = "$" + colPersL + "$" + HEURES_DATA_START_ROW + ":$" + colPersL + "$" + plageFin;
  var refLieu = "$" + colLieuL + "$" + HEURES_DATA_START_ROW + ":$" + colLieuL + "$" + plageFin;
  var refTache = "$" + colTacheL + "$" + HEURES_DATA_START_ROW + ":$" + colTacheL + "$" + plageFin;
  var refTemps = "$" + colTempsL + "$" + HEURES_DATA_START_ROW + ":$" + colTempsL + "$" + plageFin;

  // Les 6 personnes ont chacune leur propre colonne, avec correspondance EXACTE au nom
  // (plus de seau générique "Externe" — Externe 1/2/3 sont des emplacements à part entière).
  var PERSONNE_COLS_TACHE = [9, 10, 11, 12, 13, 14];  // I,J,K,L,M,N
  var PERSONNE_COLS_LIEU = [18, 19, 20, 21, 22, 23];  // R,S,T,U,V,W

  // ---------- Total par tâche (liste en H, récap en I:O) ----------
  var taches = listeVerticale(sheet, HEURES_TACHE_LIST_COL, HEURES_TACHE_LIST_START);
  taches.forEach(function (nom, idx) {
    var r = HEURES_TACHE_LIST_START + idx;
    var refLigne = "$H" + r;
    HEURES_PERSONNES.forEach(function (personne, pIdx) {
      var col = PERSONNE_COLS_TACHE[pIdx];
      var colLettre = columnToLetter(col);
      sheet.getRange(r, col).setFormula("=SUMIFS(" + refTemps + ";" + refPers + ";" + colLettre + "$3;" + refTache + ";" + refLigne + ")");
    });
    sheet.getRange(r, 15).setFormula("=SUM(I" + r + ":N" + r + ")"); // O = Total
  });
  results.push("Total par tâche : " + taches.length + " ligne(s).");

  // ---------- Total par lieu (liste en Q, récap en R:X) ----------
  var lieux = listeVerticale(sheet, HEURES_LIEU_LIST_COL, HEURES_LIEU_LIST_START);
  lieux.forEach(function (nom, idx) {
    var r = HEURES_LIEU_LIST_START + idx;
    var refLigne = "$Q" + r;
    HEURES_PERSONNES.forEach(function (personne, pIdx) {
      var col = PERSONNE_COLS_LIEU[pIdx];
      var colLettre = columnToLetter(col);
      sheet.getRange(r, col).setFormula("=SUMIFS(" + refTemps + ";" + refPers + ";" + colLettre + "$3;" + refLieu + ";" + refLigne + ")");
    });
    sheet.getRange(r, 24).setFormula("=SUM(R" + r + ":W" + r + ")"); // X = Total
  });
  results.push("Total par lieu : " + lieux.length + " ligne(s).");

  // ---------- Total par personne et par mois (liste en Z, mois en AA:AL, total AM) ----------
  for (var m = 1; m <= 12; m++) {
    var colMois = 26 + m; // AA=27 -> m=1(Janvier) ... AL=38 -> m=12(Décembre)
    var debut = "DATE(2026;" + m + ";1)";
    var finExclusive = m < 12 ? "DATE(2026;" + (m + 1) + ";1)" : "DATE(2027;1;1)";

    HEURES_PERSONNES.forEach(function (personne, pIdx) {
      var r = HEURES_PERSONNE_LIST_START + pIdx;
      var refLigne = "$Z" + r;
      sheet.getRange(r, colMois).setFormula("=SUMIFS(" + refTemps + ";" + refPers + ";" + refLigne + ";" + refDate + ";\">=\"&" + debut + ";" + refDate + ";\"<\"&" + finExclusive + ")");
    });
  }
  // Colonne TOTAL (AM = 39)
  HEURES_PERSONNES.forEach(function (personne, pIdx) {
    var r = HEURES_PERSONNE_LIST_START + pIdx;
    sheet.getRange(r, 39).setFormula("=SUM(AA" + r + ":AL" + r + ")");
  });
  results.push("Total par personne et par mois : " + HEURES_PERSONNES.length + " ligne(s) x 12 mois.");

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Fonction de migration à exécuter UNE SEULE FOIS ---------- */

function ajouterColonnePluie() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    // Traite les lignes de la droite vers la gauche (3 puis 2 puis 1) pour que l'insertion
    // dans un bloc ne décale pas la position des blocs pas encore traités à sa gauche.
    var lignesDesc = p.lignes.slice().sort(function (a, b) { return b - a; });

    lignesDesc.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable, colonne NON ajoutée"); return; }

      var colMm = block.blockCol + 8;
      var headerApresMm = sheet.getRange(block.headerRow, colMm + 1).getValue();
      if (headerApresMm && headerApresMm.toString().trim().toLowerCase() === "pluie (mm)") {
        results.push(p.sheet + " Ligne " + ligneIndex + " : déjà présente, ignorée");
        return;
      }

      sheet.insertColumnsAfter(colMm, 1);
      sheet.getRange(block.headerRow, colMm + 1).setValue("Pluie (mm)");
      results.push(p.sheet + " Ligne " + ligneIndex + " : colonne ajoutée en colonne " + columnToLetter(colMm + 1) + ", ligne d'en-tête " + block.headerRow);
    });
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Fonction de correction ponctuelle : Heure début/fin manquantes sur Ligne 3
   + repositionnement de la colonne Pluie (mm) si elle a été insérée au mauvais endroit ---------- */
// À exécuter une seule fois si "Ligne 3" de La Fontaine / La baume / Pilles n'a pas reçu
// les colonnes Heure début / Heure fin lors de l'insertion manuelle initiale, ET que
// "ajouterColonnePluie" a déjà été exécutée entre-temps (donc mal positionnée sur ces blocs).
function corrigerLigne3EtColonnePluie() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ["La Fontaine", "La baume", "Pilles"];
  var results = [];

  tabs.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { results.push(name + " : onglet introuvable"); return; }

    var block = findLigneBlock(sheet, 3);
    if (!block) { results.push(name + " Ligne 3 : bloc introuvable"); return; }

    var colAfterDate = block.blockCol + 1;
    var existingHeure = sheet.getRange(block.headerRow, colAfterDate).getValue();
    var alreadyHasHeure = existingHeure && existingHeure.toString().trim().toLowerCase() === "heure début";

    if (alreadyHasHeure) {
      results.push(name + " Ligne 3 : colonnes Heure déjà présentes, rien à faire.");
      return;
    }

    // Cherche où se trouve l'éventuelle colonne "Pluie (mm)" déjà insérée (probablement
    // décalée, puisqu'elle a été placée en supposant un bloc à 9 colonnes alors qu'il n'en avait que 7).
    var oldPluieCol = -1;
    for (var c = block.blockCol; c <= block.blockCol + 15; c++) {
      var v = sheet.getRange(block.headerRow, c).getValue();
      if (v && v.toString().trim().toLowerCase() === "pluie (mm)") { oldPluieCol = c; break; }
    }

    // Insère les 2 colonnes Heure début / Heure fin juste après Date
    sheet.insertColumnsAfter(block.blockCol, 2);
    sheet.getRange(block.headerRow, block.blockCol + 1).setValue("Heure début");
    sheet.getRange(block.headerRow, block.blockCol + 2).setValue("Heure fin");
    results.push(name + " Ligne 3 : colonnes Heure ajoutées en " + columnToLetter(block.blockCol + 1) + " et " + columnToLetter(block.blockCol + 2) + ".");

    if (oldPluieCol === -1) {
      results.push(name + " Ligne 3 : aucune colonne 'Pluie (mm)' trouvée — rien à repositionner (peut-être pas encore créée pour ce bloc).");
      return;
    }

    // L'insertion ci-dessus a décalé "Pluie (mm)" de +2
    var newPluieCol = oldPluieCol + 2;
    var correctPluieCol = block.blockCol + 9; // position correcte une fois Heure ajoutée

    if (newPluieCol <= correctPluieCol) {
      results.push(name + " Ligne 3 : colonne Pluie (mm) déjà à la bonne position, rien à repositionner.");
      return;
    }

    var nbColsASupprimer = newPluieCol - correctPluieCol;

    // Vérifie que les colonnes à supprimer (entre "mm" et l'actuelle "Pluie (mm)") sont bien vides,
    // pour ne jamais supprimer de vraies données par erreur.
    var toutEstVide = true;
    for (var cc = correctPluieCol; cc < newPluieCol && toutEstVide; cc++) {
      var nbLignes = Math.max(sheet.getLastRow() - block.headerRow + 1, 1);
      var colValues = sheet.getRange(block.headerRow, cc, nbLignes, 1).getValues();
      for (var rr = 0; rr < colValues.length; rr++) {
        if (colValues[rr][0] !== "" && colValues[rr][0] !== null) { toutEstVide = false; break; }
      }
    }

    if (toutEstVide) {
      sheet.deleteColumns(correctPluieCol, nbColsASupprimer);
      results.push(name + " Ligne 3 : colonne Pluie (mm) repositionnée correctement en colonne " + columnToLetter(correctPluieCol) + ".");
    } else {
      results.push(name + " Ligne 3 : ATTENTION — des colonnes entre 'mm' et 'Pluie (mm)' contiennent des données non vides. Repositionnement automatique annulé pour ne rien effacer par erreur : vérifie et corrige ce bloc manuellement.");
    }
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Fonction de correction ponctuelle : mise en forme de "Pluie (mm)" sur Ligne 3 ---------- */
// À exécuter si la colonne "Pluie (mm)" de Ligne 3 (La Fontaine / La baume / Pilles) existe
// et contient les bonnes valeurs, mais n'a pas la même mise en forme (bordures, couleurs,
// largeur) que le reste du tableau. Ne touche à aucune valeur, seulement au format.
function corrigerFormatColonnePluieLigne3() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tabs = ["La Fontaine", "La baume", "Pilles"];
  var results = [];

  tabs.forEach(function (name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet) { results.push(name + " : onglet introuvable"); return; }

    var block = findLigneBlock(sheet, 3);
    if (!block) { results.push(name + " Ligne 3 : bloc introuvable"); return; }

    var cols = blockColumns(block.blockCol);
    var lastRow = Math.max(sheet.getLastRow(), block.headerRow);
    var nbRows = lastRow - block.headerRow + 1;

    var sourceRange = sheet.getRange(block.headerRow, cols.mm, nbRows, 1);
    var destRange = sheet.getRange(block.headerRow, cols.pluie, nbRows, 1);

    // Copie uniquement la mise en forme (bordures, couleurs, police...) de la colonne "mm"
    // vers la colonne "Pluie (mm)" — les valeurs déjà présentes ne sont pas modifiées.
    sourceRange.copyTo(destRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
    sheet.setColumnWidth(cols.pluie, sheet.getColumnWidth(cols.mm));

    results.push(name + " Ligne 3 : mise en forme de 'Pluie (mm)' alignée sur 'mm'.");
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Import ponctuel : pluie Météociel Tallard, août 2026 ---------- */
// À exécuter une seule fois. Insère les jours où les précipitations 24h dépassaient 5 mm
// (source : meteociel.fr, station Tallard 05170001, données jusqu'au 25/08/2026 inclus),
// en réutilisant exactement la même logique que la saisie "Pluie" depuis l'appli
// (ordre chronologique respecté, écriture en orange, jamais sur une ligne TOTAL).
function importPluieMeteocielAout2026() {
  var entries = [
    { date: "2026-08-20", mm: 41.3 },
    { date: "2026-08-21", mm: 7.4 },
    { date: "2026-08-24", mm: 15.1 }
  ];

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  entries.forEach(function (entry) {
    var dateValue = parseDateOnly(entry.date);

    PARCELLES.forEach(function (p) {
      var sheet = ss.getSheetByName(p.sheet);
      if (!sheet) { results.push(entry.date + " — " + p.sheet + " : onglet introuvable"); return; }

      p.lignes.forEach(function (ligneIndex) {
        var block = findLigneBlock(sheet, ligneIndex);
        if (!block) { results.push(entry.date + " — " + p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

        var cols = blockColumns(block.blockCol);
        var targetRow = resolveTargetRow(sheet, block, cols, dateValue);
        if (targetRow === -1) { results.push(entry.date + " — " + p.sheet + " Ligne " + ligneIndex + " : plus de ligne libre"); return; }

        setValueColored(sheet.getRange(targetRow, cols.date), dateValue, COULEUR_PLUIE).setNumberFormat(FORMAT_DATE);
        setValueColored(sheet.getRange(targetRow, cols.pluie), entry.mm, COULEUR_PLUIE);
        results.push(entry.date + " (" + entry.mm + " mm) — " + p.sheet + " Ligne " + ligneIndex + " : ligne " + targetRow + " OK");
      });
    });
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Correction ponctuelle : nettoyer les dates déjà écrites (heure résiduelle) ---------- */
// À exécuter si des dates déjà saisies affichent une heure (ex. "03/05/2026 15:00:00").
// Ne touche qu'aux cellules contenant une vraie date (jamais aux textes comme "TOTAL Juin") :
// remet l'heure à minuit et réapplique le format jj/mm/aaaa, sans modifier la couleur du texte.
function normaliserDatesExistantes() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var cols = blockColumns(block.blockCol);
      var nbLignes = 300;
      var range = sheet.getRange(block.dataStartRow, cols.date, nbLignes, 1);
      var values = range.getValues();
      var corrections = 0;

      for (var i = 0; i < values.length; i++) {
        var v = values[i][0];
        if (v instanceof Date) {
          var propre = new Date(v.getFullYear(), v.getMonth(), v.getDate());
          if (propre.getTime() !== v.getTime()) {
            values[i][0] = propre;
            corrections++;
          }
        }
      }

      if (corrections > 0) {
        range.setValues(values);
        range.setNumberFormat(FORMAT_DATE);
      }
      results.push(p.sheet + " Ligne " + ligneIndex + " : " + corrections + " date(s) corrigée(s).");
    });
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Tableaux récapitulatifs (Mois | Pluie | Arrosage | cumul | Besoins) EN HAUT des onglets ---------- */
// Disposition (identique dans les 4 onglets) : lignes 1 à 14 = un tableau par ligne d'arrosage, placé
// juste au-dessus de ses données (mêmes colonnes que le bloc) :
//   ligne 1 : en-têtes · lignes 2 à 6 : mai à septembre · ligne 7 : TOTAL période estivale ·
//   lignes 8 à 14 : octobre à avril (saison suivante).
// Chaque case est une formule qui additionne TOUTES les données du bloc (de la 1re saisie jusqu'en bas
// de l'onglet, nouvelles saisies comprises) : pluie = colonne « Pluie (mm) », arrosage = colonne « mm ».
var MOIS_BILAN = [
  { nom: "Mai", num: 5, an: 2026 }, { nom: "Juin", num: 6, an: 2026 }, { nom: "Juillet", num: 7, an: 2026 },
  { nom: "Août", num: 8, an: 2026 }, { nom: "Septembre", num: 9, an: 2026 },
  { nom: "octobre", num: 10, an: 2026 }, { nom: "novembre", num: 11, an: 2026 }, { nom: "décembre", num: 12, an: 2026 },
  { nom: "janvier", num: 1, an: 2027 }, { nom: "février", num: 2, an: 2027 }, { nom: "mars", num: 3, an: 2027 },
  { nom: "avril", num: 4, an: 2027 }
];
var BESOINS_PAR_DEFAUT = [59, 78, 94, 85, 55]; // mai à septembre (Etc = Kc x ETP)
var NB_LIGNES_RECAP = 15; // 14 lignes de tableau + 1 ligne vide avant le titre de la parcelle
var COULEURS_LIGNE_RECAP = { 1: "#FCE5CD", 2: "#CFE2F3", 3: "#D9EAD3" };

// Cellule « Surface couverte (ha) » du bloc, en référence absolue (ex. $B$18), ou null.
function cellSurfaceBloc(sheet, block) {
  var n = Math.max(1, block.headerRow - 1);
  var labels = sheet.getRange(1, 1, n, 1).getValues();
  for (var r = 0; r < n; r++) {
    if (/surface/i.test((labels[r][0] || "").toString())) {
      var v = sheet.getRange(r + 1, block.blockCol).getValue();
      if (typeof v === "number" && v > 0) return "$" + columnToLetter(block.blockCol) + "$" + (r + 1);
    }
  }
  return null;
}

// Dernière ligne contenant une donnée (date, m3 ou pluie) dans un bloc.
function derniereLigneBloc(sheet, block) {
  var cols = blockColumns(block.blockCol);
  var n = Math.max(1, Math.min(1000, sheet.getMaxRows() - block.dataStartRow + 1));
  var v = sheet.getRange(block.dataStartRow, cols.date, n, cols.pluie - cols.date + 1).getValues();
  var last = block.dataStartRow - 1;
  for (var i = 0; i < n; i++) {
    for (var j = 0; j < v[i].length; j++) {
      if (v[i][j] !== "" && v[i][j] !== null) { last = block.dataStartRow + i; break; }
    }
  }
  return last;
}

// Remet la formule de la colonne « mm » : mm = (M3 / surface) / 10, avec la surface lue dans la
// cellule « Surface couverte (ha) » du bloc (plus de surface écrite en dur dans les formules).
function reecrireFormulesMm(sheet, block) {
  var surf = cellSurfaceBloc(sheet, block);
  if (!surf) return "surface introuvable, formules mm inchangées";
  var cols = blockColumns(block.blockCol);
  var last = derniereLigneBloc(sheet, block);
  if (last < block.dataStartRow) return "aucune donnée";
  var I = columnToLetter(cols.m3), f = [];
  for (var r = block.dataStartRow; r <= last; r++) f.push(["=(" + I + r + "/" + surf + ")/10"]);
  sheet.getRange(block.dataStartRow, cols.mm, f.length, 1).setFormulas(f);
  return f.length + " formule(s) mm";
}

// Écrit (ou réécrit) le tableau récapitulatif d'un bloc. besoins = 5 valeurs (mai à septembre).
function ecrireRecapBloc(sheet, block, ligneIndex, besoins) {
  var c0 = block.blockCol, ds = block.dataStartRow, cols = blockColumns(c0);
  var L = columnToLetter;
  var dateR = L(cols.date) + ds + ":" + L(cols.date);
  var pluieR = L(cols.pluie) + ds + ":" + L(cols.pluie);
  var mmR = L(cols.mm) + ds + ":" + L(cols.mm);
  var tint = COULEURS_LIGNE_RECAP[ligneIndex] || "#F1EEE3";

  var head = sheet.getRange(1, c0, 1, 5);
  head.setValues([["Mois", "Pluie (mm)", "Arrosage (mm)", "cumul pluie+arrosage", "Besoins de plante, sans pluie Etc (mm) = Kc x ETP"]]);
  head.setBackground("#073763").setFontColor("#FFFFFF").setFontWeight("bold");

  MOIS_BILAN.forEach(function (m, i) {
    var row = i < 5 ? 2 + i : 3 + i; // la ligne 7 est le total de la période estivale
    var deb = "DATE(" + m.an + ";" + m.num + ";1)";
    var fin = m.num === 12 ? "DATE(" + (m.an + 1) + ";1;1)" : "DATE(" + m.an + ";" + (m.num + 1) + ";1)";
    var crit = dateR + ";\">=\"&" + deb + ";" + dateR + ";\"<\"&" + fin;
    sheet.getRange(row, c0).setValue(m.nom).setFontWeight("normal").setBackground(null);
    sheet.getRange(row, c0 + 1).setFormula("=SUMIFS(" + pluieR + ";" + crit + ")").setBackground(tint);
    sheet.getRange(row, c0 + 2).setFormula("=SUMIFS(" + mmR + ";" + crit + ")").setBackground(tint);
    sheet.getRange(row, c0 + 3).setFormula("=" + L(c0 + 1) + row + "+" + L(c0 + 2) + row);
    if (i < 5) sheet.getRange(row, c0 + 4).setValue(besoins && besoins[i] !== "" && besoins[i] !== null ? besoins[i] : BESOINS_PAR_DEFAUT[i]);
    sheet.getRange(row, c0 + 1, 1, 4).setNumberFormat("0");
  });

  sheet.getRange(7, c0).setValue("TOTAL periode estivale").setFontWeight("bold");
  for (var k = 1; k <= 4; k++) {
    sheet.getRange(7, c0 + k).setFormula("=SUM(" + L(c0 + k) + "2:" + L(c0 + k) + "6)").setFontWeight("bold").setNumberFormat("0");
  }

  // Code couleur dynamique sur Cumul / Besoins (vert : besoin couvert ; jaune : déficit jusqu'à 20 % ;
  // rouge : déficit de plus de 20 %). Les mois sans besoin renseigné ne sont pas colorés.
  var cu = L(c0 + 3), be = L(c0 + 4);
  var rng = sheet.getRange(2, c0 + 3, 13, 2);
  var VERT = "#B7E1CD", JAUNE = "#FCE8B2", ROUGE = "#F4C7C3";
  var regles = sheet.getConditionalFormatRules().filter(function (rule) {
    return !rule.getRanges().some(function (rg) {
      return rg.getRow() <= 14 && rg.getColumn() <= c0 + 4 && rg.getLastColumn() >= c0 + 3;
    });
  });
  [["=AND($" + be + "2>0;$" + cu + "2>=$" + be + "2)", VERT],
   ["=AND($" + be + "2>0;$" + cu + "2<$" + be + "2;$" + cu + "2>=$" + be + "2*0.8)", JAUNE],
   ["=AND($" + be + "2>0;$" + cu + "2<$" + be + "2*0.8)", ROUGE]].forEach(function (x) {
    regles.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(x[0]).setBackground(x[1]).setRanges([rng]).build());
  });
  sheet.setConditionalFormatRules(regles);
}

// À EXÉCUTER UNE SEULE FOIS (après avoir fait une copie du fichier : Fichier > Créer une copie).
// - Onglets dont le tableau est encore sous les données (Bousquet et Riou, La baume, Pilles) : le
//   tableau est supprimé du milieu des données (les saisies qui le suivaient remontent), puis 15 lignes
//   sont insérées en haut et les tableaux récapitulatifs y sont reconstruits (mai à avril).
// - La Fontaine (déjà remontée par tes soins) : les formules sont réécrites sur le même modèle.
// Dans tous les cas : les formules « mm » sont remises d'aplomb (surface lue dans « Surface couverte »)
// et les tableaux couvrent toutes les données saisies (de la 1re ligne de données jusqu'en bas).
function mettreAJourTableauxRecap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  PARCELLES.forEach(function (p) {
    var sheet = findSheetByName(ss, p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }
    var blocks = [];
    p.lignes.forEach(function (li) { var b = findLigneBlock(sheet, li); if (b) blocks.push({ li: li, b: b }); });
    if (!blocks.length) { results.push(p.sheet + " : blocs introuvables, rien modifié."); return; }

    var enteteRow = trouverEnteteBilan(sheet, blocks[0].b);
    var dejaEnHaut = enteteRow !== -1 && enteteRow < blocks[0].b.headerRow;
    var besoinsParLigne = {};

    if (enteteRow !== -1) {
      blocks.forEach(function (x) {
        var vals = sheet.getRange(enteteRow + 1, x.b.blockCol + 4, 5, 1).getValues();
        besoinsParLigne[x.li] = vals.map(function (v) { return v[0]; });
      });
    }

    if (!dejaEnHaut) {
      if (enteteRow === -1) { results.push(p.sheet + " : tableau récapitulatif introuvable, rien modifié."); return; }
      // sécurité : le tableau doit faire exactement 7 lignes (en-tête, 5 mois, TOTAL)
      var finLabel = normTxt(sheet.getRange(enteteRow + 6, blocks[0].b.blockCol).getValue());
      if (finLabel.indexOf("total") !== 0) { results.push(p.sheet + " : disposition inattendue sous les données, rien modifié."); return; }
      sheet.deleteRows(enteteRow, 7);
      sheet.insertRowsBefore(1, NB_LIGNES_RECAP);
      sheet.getRange(1, 1, NB_LIGNES_RECAP, sheet.getMaxColumns()).clear();
      // les repères de blocs ont bougé : on les relit
      blocks = [];
      p.lignes.forEach(function (li) { var b = findLigneBlock(sheet, li); if (b) blocks.push({ li: li, b: b }); });
      results.push(p.sheet + " : tableaux déplacés en haut (anciennes lignes " + enteteRow + " à " + (enteteRow + 6) + " retirées des données).");
    }

    blocks.forEach(function (x) {
      ecrireRecapBloc(sheet, x.b, x.li, besoinsParLigne[x.li]);
      var m = reecrireFormulesMm(sheet, x.b);
      results.push(p.sheet + " Ligne " + x.li + " : tableau mai→avril écrit ; " + m + ".");
    });
  });
  Logger.log(results.join("\n"));
  return results;
}

// Contrôle (n'écrit rien) : où le script trouve les blocs, les tableaux et le début des données.
function verifierTableauxRecap() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), out = [];
  PARCELLES.forEach(function (p) {
    var sheet = findSheetByName(ss, p.sheet);
    if (!sheet) { out.push(p.sheet + " : onglet introuvable"); return; }
    p.lignes.forEach(function (li) {
      var b = findLigneBlock(sheet, li);
      if (!b) { out.push(p.sheet + " Ligne " + li + " : bloc introuvable"); return; }
      var h = trouverEnteteBilan(sheet, b);
      out.push(p.sheet + " Ligne " + li + " : colonne " + columnToLetter(b.blockCol) + ", en-têtes ligne " + b.headerRow +
        ", données dès la ligne " + b.dataStartRow + ", tableau récapitulatif " +
        (h === -1 ? "INTROUVABLE" : (h < b.headerRow ? "en haut (ligne " + h + ")" : "sous les données (ligne " + h + ") → lance mettreAJourTableauxRecap")) +
        ", surface " + (cellSurfaceBloc(sheet, b) || "non trouvée"));
    });
  });
  Logger.log(out.join("\n"));
  return out;
}

/* ---------- Fonctions utilitaires partagées ---------- */

function findSheetByName(ss, name) {
  var wanted = (name || "").trim().toLowerCase();
  var found = null;
  ss.getSheets().forEach(function (s) {
    if (s.getName().trim().toLowerCase() === wanted) found = s;
  });
  return found;
}

// Repère la colonne de bloc en cherchant le libellé "Ligne N..." dans les 6 premières lignes,
// puis la ligne d'en-tête "Date" juste en dessous.
function findLigneBlock(sheet, ligneIndex) {
  // Le libellé « Ligne N » est cherché dans les 40 premières lignes : depuis que les tableaux
  // récapitulatifs sont en haut de l'onglet, il se trouve plus bas (ex. ligne 17 à La Fontaine).
  var scanRows = Math.min(40, sheet.getMaxRows());
  var scanCols = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, scanRows, scanCols).getValues();
  var ligneRegex = new RegExp("^Ligne\\s*" + ligneIndex + "(\\s|\\(|$)", "i");

  var blockCol = -1, ligneRow = -1;
  for (var r = 0; r < scanRows && blockCol === -1; r++) {
    for (var c = 0; c < scanCols; c++) {
      var v = values[r][c];
      if (v && ligneRegex.test(v.toString().trim())) {
        blockCol = c + 1;
        ligneRow = r + 1;
        break;
      }
    }
  }
  if (blockCol === -1) return null;

  var headerRow = -1;
  for (var rr = ligneRow; rr <= ligneRow + 10; rr++) {
    var cellVal = sheet.getRange(rr, blockCol).getValue();
    if (cellVal && cellVal.toString().trim().toLowerCase() === "date") {
      headerRow = rr;
      break;
    }
  }
  if (headerRow === -1) return null;

  return { blockCol: blockCol, ligneRow: ligneRow, headerRow: headerRow, dataStartRow: headerRow + 2 };
}

// Colonnes du bloc (relatives à blockCol) :
// +0 Date | +1 Heure début | +2 Heure fin | +3 Durée (h) | +4 Compteur début | +5 Compteur fin
// | +6 Débit (formule) | +7 M3 | +8 mm (formule) | +9 Pluie (mm)
function blockColumns(blockCol) {
  return {
    date: blockCol,
    heureDebut: blockCol + 1,
    heureFin: blockCol + 2,
    duree: blockCol + 3,
    debut: blockCol + 4,
    fin: blockCol + 5,
    debit: blockCol + 6,
    m3: blockCol + 7,
    mm: blockCol + 8,
    pluie: blockCol + 9
  };
}

// Une ligne est considérée libre si M3 ET Date sont vides. Les lignes TOTAL ont toujours
// une formule (donc non vides) dans M3, donc automatiquement ignorées. Vérifier aussi Date
// (et pas seulement M3) évite de réutiliser une ligne déjà occupée par une saisie Pluie
// (qui remplit Date + Pluie(mm) mais laisse M3 vide).
function findNextFreeRow(sheet, dataStartRow, colM3, colDate) {
  var maxRows = sheet.getMaxRows();
  var n = Math.min(maxRows - dataStartRow + 1, 1000);
  if (n >= 1) {
    var m3 = sheet.getRange(dataStartRow, colM3, n, 1).getValues();
    var dt = sheet.getRange(dataStartRow, colDate, n, 1).getValues();
    for (var i = 0; i < n; i++) {
      var m3Blank = (m3[i][0] === "" || m3[i][0] === null);
      var dateBlank = (dt[i][0] === "" || dt[i][0] === null);
      if (m3Blank && dateBlank) return dataStartRow + i;
    }
  }
  // Tableau plein : on ajoute des lignes en bas de l'onglet (sans dépasser ~1000 lignes de données)
  if (maxRows - dataStartRow + 1 < 1000) {
    sheet.insertRowsAfter(maxRows, 50);
    return maxRows + 1;
  }
  return -1;
}

function copyFormulaIfMissing(sheet, targetRow, col, dataStartRow) {
  var cell = sheet.getRange(targetRow, col);
  if (!cell.getFormula() && targetRow > dataStartRow) {
    sheet.getRange(targetRow - 1, col).copyTo(cell, { formulaOnly: true });
  }
}

// Repère le début du "segment" (mois) en cours : la ligne juste après la dernière ligne TOTAL
// rencontrée avant nextFreeRow (ou dataStartRow s'il n'y en a aucune). Le repérage se fait en
// détectant une formule SUM dans la colonne M3 (les lignes TOTAL ont toujours ce type de formule).
function findSegmentStart(sheet, dataStartRow, colM3, nextFreeRow) {
  var segStart = dataStartRow;
  for (var r = dataStartRow; r < nextFreeRow; r++) {
    var formula = sheet.getRange(r, colM3).getFormula();
    if (formula && formula.indexOf("SUM") !== -1) {
      segStart = r + 1;
    }
  }
  return segStart;
}

// Détermine la ligne où écrire une nouvelle saisie, en respectant l'ordre chronologique
// à l'intérieur du mois (segment) en cours. Si la date saisie est antérieure à une date déjà
// présente dans le segment, les lignes concernées sont décalées d'un cran vers le bas (jamais
// au-delà d'une ligne TOTAL, pour ne jamais perturber un mois déjà clos).
function resolveTargetRow(sheet, block, cols, dateValue) {
  var nextFreeRow = findNextFreeRow(sheet, block.dataStartRow, cols.m3, cols.date);
  if (nextFreeRow === -1) return -1;

  if (!(dateValue instanceof Date)) {
    copyFormulaIfMissing(sheet, nextFreeRow, cols.debit, block.dataStartRow);
    copyFormulaIfMissing(sheet, nextFreeRow, cols.mm, block.dataStartRow);
    return nextFreeRow;
  }

  var segStart = findSegmentStart(sheet, block.dataStartRow, cols.m3, nextFreeRow);

  var insertionRow = nextFreeRow;
  for (var r = segStart; r < nextFreeRow; r++) {
    var v = sheet.getRange(r, cols.date).getValue();
    if (v instanceof Date && dateValue.getTime() < v.getTime()) {
      insertionRow = r;
      break;
    }
  }

  if (insertionRow < nextFreeRow) {
    var blockWidth = cols.pluie - cols.date + 1;
    shiftBlockRowsDown(sheet, cols.date, blockWidth, insertionRow, nextFreeRow);
  }

  copyFormulaIfMissing(sheet, insertionRow, cols.debit, block.dataStartRow);
  copyFormulaIfMissing(sheet, insertionRow, cols.mm, block.dataStartRow);

  return insertionRow;
}

// Décale vers le bas, d'une ligne, tout le contenu du bloc (valeurs + formules + mise en forme)
// entre fromRow et toRow (inclus/exclu comme un intervalle [fromRow, toRow]). Ne touche qu'aux
// colonnes de CE bloc — les blocs voisins (autres lignes de la même parcelle) ne sont pas affectés.
function shiftBlockRowsDown(sheet, blockCol, blockWidth, fromRow, toRow) {
  for (var r = toRow; r > fromRow; r--) {
    var src = sheet.getRange(r - 1, blockCol, 1, blockWidth);
    var dest = sheet.getRange(r, blockCol, 1, blockWidth);
    src.copyTo(dest);
  }
  sheet.getRange(fromRow, blockCol, 1, blockWidth).clearContent();
}

function setValueColored(range, value, color) {
  return range.setValue(value).setFontColor(color);
}

function parseHeure(hhmm) {
  var parts = hhmm.split(":");
  var h = parseInt(parts[0], 10) || 0;
  var m = parseInt(parts[1], 10) || 0;
  return new Date(1899, 11, 30, h, m, 0);
}

function columnToLetter(col) {
  var letter = "";
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
