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
    tableauPoids: obtenirTableauPoids
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

// Récupère, pour les 11 blocs, le tableau de bilan mensuel déjà calculé (formules déjà
// en place dans le Sheet) : Mois | Pluie (mm) | Arrosage (mm) | cumul | Besoins (Etc).
function obtenirTousLesBilans() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) return;

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) return;

      var colMois = block.blockCol;
      var headerRow2 = -1;
      for (var r = block.headerRow; r < block.headerRow + 90; r++) {
        var v = sheet.getRange(r, colMois).getValue();
        if (v && v.toString().trim().toLowerCase() === "mois") { headerRow2 = r; break; }
      }
      if (headerRow2 === -1) return;

      var rows = sheet.getRange(headerRow2 + 1, colMois, 6, 5).getValues(); // 5 mois + TOTAL
      var mois = rows.map(function (row) {
        return {
          mois: row[0] ? row[0].toString() : "",
          pluie: Number(row[1]) || 0,
          arrosage: Number(row[2]) || 0,
          cumul: Number(row[3]) || 0,
          besoins: Number(row[4]) || 0
        };
      });

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
   RÉCOLTE — onglet "Récolte" (une ligne par truffe pointée)
   Colonnes : A Truffes No | B Date | C Parcelle | D Poids | E Qualité | F Image
              | G Localisation | H Commentaires | I Latitude | J Longitude
   (E et F restent à remplir à la main ; I et J sont des colonnes techniques
   « lat / lon » en texte, utilisées par la carte de l'appli et pour Google My Maps.)
   + onglet "Récolte volume" (poids total après récolte, par parcelle)
   ================================================================ */
var RECOLTE_SHEET = "Récolte";
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
  var sheet = trouverOngletSouple(SpreadsheetApp.getActiveSpreadsheet(), RECOLTE_SHEET);
  if (!sheet) return jsonOutput({ ok: false, error: "Onglet 'Récolte' introuvable." });

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
  var sheet = trouverOngletSouple(SpreadsheetApp.getActiveSpreadsheet(), RECOLTE_SHEET);
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
  var sheet = trouverOngletSouple(SpreadsheetApp.getActiveSpreadsheet(), RECOLTE_SHEET);
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
// Structure attendue : une ligne par DATE ; colonne « Date » ; puis, pour chaque parcelle, un
// groupe de colonnes (Poids récolté | Poids brut | Poids net vendable) ; puis des colonnes de
// total (Total brut, Total net vendable…). Les colonnes sont repérées d'après les EN-TÊTES (nom
// de la parcelle au-dessus, intitulé du poids en dessous — cellules fusionnées acceptées), pas
// d'après leur position : on peut les déplacer. Utiliser « verifierOngletVolume » pour voir
// comment les en-têtes sont interprétés avant de saisir quoi que ce soit.
var PARCELLES_VOLUME = [
  { key: "fontaine", label: "La Fontaine", re: /fontaine/ },
  { key: "baume", label: "La Baume", re: /baume/ },
  { key: "bousquet", label: "Bousquet", re: /bousquet/ },
  { key: "riou", label: "Riou", re: /riou/ },
  { key: "pilles", label: "Les Pilles", re: /pilles/ }
];
var TYPES_POIDS = ["recolte", "brut", "net"];
var LIBELLES_TYPES = { recolte: "récolté", brut: "brut", net: "net vendable" };

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
function sousTypeTexte(t) {
  if (/\bbrut\b/.test(t)) return "brut";
  if (/\bnet\b|vendable/.test(t)) return "net";
  if (/recolt|poids|volume/.test(t)) return "recolte";
  return null;
}
function jourStr(d, tz) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }

// Lit les en-têtes de l'onglet et renvoie { dateCol, startRow, groups, totals, units, ... }
function analyserOngletVolume(sheet) {
  var lastCol = Math.max(sheet.getLastColumn(), 1);
  var nbRows = Math.min(10, Math.max(sheet.getMaxRows(), 1));
  var vals = sheet.getRange(1, 1, nbRows, lastCol).getValues();

  var dateCol = 0, dateRow = 0, r, c;
  for (r = 0; r < nbRows && !dateCol; r++) {
    for (c = 0; c < lastCol && !dateCol; c++) {
      if (typeof vals[r][c] === "string" && /^date/.test(normTxt(vals[r][c]))) { dateCol = c + 1; dateRow = r + 1; }
    }
  }
  if (!dateCol) return { error: "Onglet '" + VOLUME_SHEET + "' : en-tête « Date » introuvable dans les 10 premières lignes." };

  // dernière ligne d'en-têtes = dernière ligne (avant les données) qui contient un intitulé
  var lastHeaderRow = dateRow;
  for (r = dateRow; r < nbRows; r++) {
    if (vals[r][dateCol - 1] instanceof Date) break;
    var aIntitule = vals[r].some(function (v) {
      var t = normTxt(v);
      return t && (/brut|\bnet\b|vendable|recolt|poids|total|date/.test(t) || cleParcelleTexte(t));
    });
    if (aIntitule) lastHeaderRow = r + 1;
  }

  // groupe (parcelle) de chaque colonne : le nom est au-dessus, sur la 1re colonne du groupe
  var groupAt = [];
  for (c = 0; c < lastCol; c++) groupAt[c] = null;
  for (r = 0; r < lastHeaderRow; r++) {
    var cur = null;
    for (c = 0; c < lastCol; c++) {
      var t = normTxt(vals[r][c]);
      if (t) cur = /total/.test(t) ? null : cleParcelleTexte(t);
      if (cur && !groupAt[c]) groupAt[c] = cur;
    }
  }

  var groups = {}, totals = {}, units = {}, commentCol = 0, dernierCol = dateCol, doublon = null;
  var colsDuGroupe = {};
  for (c = 0; c < lastCol; c++) {
    var label = "";
    for (r = 0; r < lastHeaderRow; r++) label += " " + normTxt(vals[r][c]);
    label = label.trim();
    if (!label) continue;
    units[c + 1] = /\bkg\b|kilo/.test(label) ? "kg" : "g";
    if (/commentaire|remarque/.test(label)) { commentCol = c + 1; dernierCol = Math.max(dernierCol, c + 1); continue; }
    var sub = sousTypeTexte(label);
    if (/total/.test(label)) {
      if (sub && !totals[sub]) { totals[sub] = c + 1; dernierCol = Math.max(dernierCol, c + 1); }
      continue;
    }
    var g = groupAt[c];
    if (!g) continue;
    dernierCol = Math.max(dernierCol, c + 1);
    (colsDuGroupe[g] = colsDuGroupe[g] || []).push(c + 1);
    if (sub) {
      groups[g] = groups[g] || {};
      if (groups[g][sub]) doublon = "Parcelle « " + g + " » : deux colonnes « " + LIBELLES_TYPES[sub] + " ».";
      else groups[g][sub] = c + 1;
    }
  }
  // repli : un groupe de 3 colonnes sans intitulés reconnus = récolté | brut | net, dans cet ordre
  Object.keys(colsDuGroupe).forEach(function (g) {
    if (!groups[g] && colsDuGroupe[g].length === 3) {
      groups[g] = { recolte: colsDuGroupe[g][0], brut: colsDuGroupe[g][1], net: colsDuGroupe[g][2] };
    }
  });
  if (doublon) return { error: doublon };

  var ordre = Object.keys(colsDuGroupe).sort(function (a, b) { return colsDuGroupe[a][0] - colsDuGroupe[b][0]; });
  return {
    dateCol: dateCol, headerRow: lastHeaderRow, startRow: lastHeaderRow + 1, groups: groups, totals: totals,
    units: units, commentCol: commentCol, dernierCol: dernierCol, ordre: ordre
  };
}

// Onglet + analyse. Si l'onglet n'existe pas et creer = true, il est créé avec la structure
// décrite ci-dessus (en-têtes sur 2 lignes, 5 parcelles, totaux brut / net vendable).
function infosOngletVolume(creer) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = trouverOngletSouple(ss, VOLUME_SHEET);
  if (!sheet) {
    if (!creer) return null;
    sheet = ss.insertSheet(VOLUME_SHEET);
    var ligne1 = ["Date"], ligne2 = [""];
    PARCELLES_VOLUME.forEach(function (p) {
      ligne1.push(p.label, "", "");
      ligne2.push("Poids récolté (g)", "Poids brut (g)", "Poids net vendable (g)");
    });
    ligne1.push("Total brut (g)", "Total net vendable (g)");
    ligne2.push("", "");
    sheet.getRange(1, 1, 1, ligne1.length).setValues([ligne1]);
    sheet.getRange(2, 1, 1, ligne2.length).setValues([ligne2]);
    sheet.getRange(1, 1, 2, ligne1.length).setBackground("#D9D9D9").setFontWeight("bold");
  }
  var info = analyserOngletVolume(sheet);
  if (info.error) return creer ? info : null;
  info.sheet = sheet;
  info.tz = ss.getSpreadsheetTimeZone();
  return info;
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
  var NB_MAX = 3000;
  var bloc = lireBlocBorne(sheet, start, 1, info.dernierCol, NB_MAX);

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
      if (bloc.length < NB_MAX) { sheet.insertRowsAfter(sheet.getMaxRows(), 200); nextFree = start + bloc.length; }
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

  var commentaire = (data.commentaire || "").toString().trim();
  if (commentaire) {
    if (info.commentCol) {
      var cc = sheet.getRange(target, info.commentCol), ancien = (cc.getValue() || "").toString();
      cc.setValue(ancien ? ancien + " ; " + commentaire : commentaire);
    } else {
      var ancienneNote = cellule.getNote() || "";
      cellule.setNote((ancienneNote ? ancienneNote + "\n" : "") + commentaire); // pas de colonne commentaire : note sur la cellule
    }
  }

  // Totaux : formules dans les colonnes « Total … » de la ligne saisie, et des lignes qui ont
  // été décalées d'un cran vers le bas (leurs formules sont réécrites, sans dépendre de la
  // façon dont la copie ajuste les références).
  ecrireTotauxLigne(sheet, info, target);
  if (rowSame === -1 && insertion !== -1) {
    for (var rr = insertion + 1; rr <= nextFree; rr++) ecrireTotauxLigne(sheet, info, rr);
  }

  return jsonOutput({ ok: true, row: target, cumul: cumul });
}
// Total brut / net / récolté d'une ligne = somme des cellules correspondantes de toutes les parcelles.
function ecrireTotauxLigne(sheet, info, row) {
  Object.keys(info.totals).forEach(function (type) {
    var refs = [];
    info.ordre.forEach(function (g) {
      if (info.groups[g] && info.groups[g][type]) refs.push(columnToLetter(info.groups[g][type]) + row);
    });
    if (refs.length) sheet.getRange(row, info.totals[type]).setFormula("=" + refs.join("+"));
  });
}

// Lit tout le tableau et le met en forme pour l'appli (poids toujours en grammes).
function lireTableauPoids(info) {
  var sheet = info.sheet, tz = info.tz;
  var vals = lireBlocBorne(sheet, info.startRow, 1, info.dernierCol, 3000);
  function enG(v, col) {
    var n = Number(v);
    if (!isFinite(n) || v === "" || v === null) return 0;
    return (info.units[col] === "kg") ? Math.round(n * 1000) : n;
  }
  var parcelles = info.ordre.filter(function (g) { return info.groups[g]; }).map(function (g) {
    var label = g;
    PARCELLES_VOLUME.forEach(function (p) { if (p.key === g) label = p.label; });
    return { key: g, label: label };
  });
  var rows = [], totaux = {}, totalBrut = 0, totalNet = 0, totalRecolte = 0, signature = 0;
  parcelles.forEach(function (p) { totaux[p.key] = { recolte: 0, brut: 0, net: 0 }; });

  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][info.dateCol - 1];
    if (d === "" || d === null) break;
    var ligne = { date: d instanceof Date ? formatDateFR(d) : d, v: {}, totalBrut: 0, totalNet: 0 };
    parcelles.forEach(function (p) {
      var grp = info.groups[p.key], o = {}, aDesDonnees = false;
      TYPES_POIDS.forEach(function (t) {
        var g = grp[t] ? enG(vals[i][grp[t] - 1], grp[t]) : 0;
        o[t] = g;
        if (g) aDesDonnees = true;
        totaux[p.key][t] += g;
        signature += g;
      });
      if (aDesDonnees) ligne.v[p.key] = o;
      ligne.totalBrut += o.brut;
      ligne.totalNet += o.net;
    });
    totalBrut += ligne.totalBrut; totalNet += ligne.totalNet;
    parcelles.forEach(function (p) { totalRecolte += (ligne.v[p.key] ? ligne.v[p.key].recolte : 0); });
    rows.push(ligne);
  }
  return { unite: "g", parcelles: parcelles, rows: rows, totaux: totaux, totalBrut: totalBrut, totalNet: totalNet, totalRecolte: totalRecolte, signature: signature };
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
  if (!info) return { unite: "g", parcelles: [], rows: [], totaux: {}, totalBrut: 0, totalNet: 0, totalRecolte: 0, signature: 0 };
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
      lignes.push(p.label + " : récolté=" + (g.recolte ? columnToLetter(g.recolte) : "—") + " · brut=" + (g.brut ? columnToLetter(g.brut) : "—") + " · net=" + (g.net ? columnToLetter(g.net) : "—") + " (unité " + (info.units[g.recolte || g.brut || g.net] || "g") + ")");
    });
    TYPES_POIDS.forEach(function (t) { if (info.totals[t]) lignes.push("Total " + LIBELLES_TYPES[t] + " : colonne " + columnToLetter(info.totals[t])); });
    if (!info.totals.brut || !info.totals.net) lignes.push("⚠ Colonne de total brut et/ou net non reconnue (en-tête contenant « total » + « brut » / « net »).");
    lignes.push(info.commentCol ? "Commentaires : colonne " + columnToLetter(info.commentCol) : "Pas de colonne Commentaires : le commentaire est mis en note sur la cellule du poids.");
  }
  Logger.log(lignes.join("\n"));
  return lignes;
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

/* ---------- Reproduction du bilan mensuel (Pluie/Arrosage/cumul/Besoins) sous chaque ligne ---------- */
// À exécuter une seule fois. Reproduit sous les 10 blocs restants (tous sauf La Fontaine
// Ligne 1, qui contient déjà le modèle d'origine construit à la main) le même tableau :
// Mois | Pluie (mm) | Arrosage (mm) | cumul pluie+arrosage | Besoins de plante Etc (mm).
// Pluie (mm) et Arrosage (mm) sont des SUMIFS calculés sur les colonnes Date/Pluie(mm)/mm
// du bloc concerné, pour chaque mois de mai à septembre 2026. Le fond des colonnes
// Pluie (mm) et Arrosage (mm) reprend la couleur de la ligne (orange clair Ligne 1,
// bleu clair Ligne 2, vert clair Ligne 3). Les besoins (Etc) sont les mêmes 5 valeurs pour
// toutes les parcelles/lignes, comme dans le modèle d'origine.
function ajouterBilanMensuel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  var mois = [
    { nom: "Mai", num: 5, besoin: 59 },
    { nom: "Juin", num: 6, besoin: 78 },
    { nom: "Juillet", num: 7, besoin: 94 },
    { nom: "Août", num: 8, besoin: 85 },
    { nom: "Septembre", num: 9, besoin: 55 }
  ];

  var COULEURS_LIGNE = { 1: "#FCE5CD", 2: "#CFE2F3", 3: "#D9EAD3" };
  var COULEUR_HEADER_BG = "#073763";
  var COULEUR_HEADER_TXT = "#FFFFFF";

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      if (p.sheet === "La Fontaine" && ligneIndex === 1) {
        results.push(p.sheet + " Ligne 1 : déjà existant (modèle d'origine), ignoré.");
        return;
      }

      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var cols = blockColumns(block.blockCol);
      var colMois = block.blockCol;         // même colonne que "Date"
      var colPluieBilan = block.blockCol + 1;
      var colArrosageBilan = block.blockCol + 2;
      var colCumul = block.blockCol + 3;
      var colBesoins = block.blockCol + 4;

      // Cherche une zone de 7 lignes x 5 colonnes libre, en partant d'un point de départ
      // estimé (même écart que le modèle d'origine), et en descendant si besoin.
      var startRow = block.headerRow + 38;
      var found = false;
      for (var t = 0; t < 30; t++) {
        var libre = true;
        for (var rr = startRow; rr < startRow + 7 && libre; rr++) {
          var rowVals = sheet.getRange(rr, colMois, 1, 5).getValues()[0];
          for (var cc = 0; cc < rowVals.length; cc++) {
            if (rowVals[cc] !== "" && rowVals[cc] !== null) { libre = false; break; }
          }
        }
        if (libre) { found = true; break; }
        startRow += 1;
      }
      if (!found) {
        results.push(p.sheet + " Ligne " + ligneIndex + " : aucune zone libre trouvée, tableau NON ajouté.");
        return;
      }

      var headerRow2 = startRow;
      var headers = ["Mois", "Pluie (mm)", "Arrosage (mm)", "cumul pluie+arrosage", "Besoins de plante, sans pluie Etc (mm) = Kc x ETP"];
      var headerRange = sheet.getRange(headerRow2, colMois, 1, 5);
      headerRange.setValues([headers]);
      headerRange.setBackground(COULEUR_HEADER_BG).setFontColor(COULEUR_HEADER_TXT).setFontWeight("bold");

      var couleurLigne = COULEURS_LIGNE[ligneIndex] || "#FFFFFF";
      var colDateLettre = columnToLetter(cols.date);
      var colPluieLettre = columnToLetter(cols.pluie);
      var colMmLettre = columnToLetter(cols.mm);
      var plageDebut = block.dataStartRow;
      var plageFin = block.dataStartRow + 300; // fenêtre large mais bornée (évite les colonnes entières)
      var dateRange = colDateLettre + plageDebut + ":" + colDateLettre + plageFin;
      var pluieRange = colPluieLettre + plageDebut + ":" + colPluieLettre + plageFin;
      var mmRange = colMmLettre + plageDebut + ":" + colMmLettre + plageFin;

      mois.forEach(function (m, idx) {
        var r = headerRow2 + 1 + idx;
        // Point-virgule comme séparateur d'arguments : ce classeur est en locale française,
        // où Sheets attend ";" entre les arguments d'une fonction (la virgule sert de
        // séparateur décimal). Une formule construite avec des virgules provoque une
        // "Erreur d'analyse de formule".
        var debut = "DATE(2026;" + m.num + ";1)";
        var finExclusive = "DATE(2026;" + (m.num + 1) + ";1)"; // 1er du mois suivant (borne exclusive)

        sheet.getRange(r, colMois).setValue(m.nom);

        // SUMPRODUCT + IFERROR : ignore les cellules en erreur (ex. #REF! isolés plus loin
        // dans le tableau) au lieu de faire planter tout le calcul, contrairement à SUMIFS
        // sur une plage contenant ne serait-ce qu'une seule cellule en erreur.
        // IFERROR englobe ici TOUTE l'opération (comparaisons de dates ET valeur à sommer),
        // pas seulement la lecture des plages : ça protège aussi contre une comparaison
        // texte/date invalide (ex. la cellule "TOTAL Juin" mélangée aux vraies dates), qui
        // provoquerait sinon une #ERROR! non capturée par un IFERROR posé uniquement sur la plage.
        // La borne haute est exclusive (< 1er du mois suivant) plutôt que <= dernier jour, pour
        // rester correct même si une saisie du dernier jour du mois garde une heure résiduelle.
        var pluieCell = sheet.getRange(r, colPluieBilan);
        pluieCell.setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + pluieRange + ";0))");
        pluieCell.setBackground(couleurLigne);

        var arrosageCell = sheet.getRange(r, colArrosageBilan);
        arrosageCell.setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + mmRange + ";0))");
        arrosageCell.setBackground(couleurLigne);

        var cumulCell = sheet.getRange(r, colCumul);
        cumulCell.setFormula("=" + columnToLetter(colPluieBilan) + r + "+" + columnToLetter(colArrosageBilan) + r);

        sheet.getRange(r, colBesoins).setValue(m.besoin);
      });

      var totalRow = headerRow2 + 1 + mois.length;
      var firstDataRow2 = headerRow2 + 1;
      var lastDataRow2 = headerRow2 + mois.length;

      sheet.getRange(totalRow, colMois).setValue("TOTAL").setFontWeight("bold");
      sheet.getRange(totalRow, colPluieBilan).setFormula("=SUM(" + columnToLetter(colPluieBilan) + firstDataRow2 + ":" + columnToLetter(colPluieBilan) + lastDataRow2 + ")").setFontWeight("bold");
      sheet.getRange(totalRow, colArrosageBilan).setFormula("=SUM(" + columnToLetter(colArrosageBilan) + firstDataRow2 + ":" + columnToLetter(colArrosageBilan) + lastDataRow2 + ")").setFontWeight("bold");
      sheet.getRange(totalRow, colCumul).setFormula("=SUM(" + columnToLetter(colCumul) + firstDataRow2 + ":" + columnToLetter(colCumul) + lastDataRow2 + ")").setFontWeight("bold");
      sheet.getRange(totalRow, colBesoins).setFormula("=SUM(" + columnToLetter(colBesoins) + firstDataRow2 + ":" + columnToLetter(colBesoins) + lastDataRow2 + ")").setFontWeight("bold");

      results.push(p.sheet + " Ligne " + ligneIndex + " : tableau ajouté à partir de la ligne " + headerRow2 + ".");
    });
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Correction ponctuelle : formules du bilan mensuel déjà créé ---------- */
// À exécuter si "ajouterBilanMensuel" a déjà été lancée et que les cellules Pluie (mm) /
// Arrosage (mm) affichent une erreur de calcul. Repère chaque tableau déjà en place
// (recherche du mot "Mois" dans la bonne colonne) et remplace uniquement les formules
// Pluie/Arrosage/TOTAL par une version insensible aux erreurs isolées (IFERROR).
// Ne touche ni aux en-têtes, ni aux couleurs, ni aux valeurs "Besoins" déjà en place.
function corrigerFormulesBilanMensuel() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  var mois = [
    { nom: "Mai", num: 5 },
    { nom: "Juin", num: 6 },
    { nom: "Juillet", num: 7 },
    { nom: "Août", num: 8 },
    { nom: "Septembre", num: 9 }
  ];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      if (p.sheet === "La Fontaine" && ligneIndex === 1) {
        results.push(p.sheet + " Ligne 1 : modèle d'origine, ignoré.");
        return;
      }

      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var cols = blockColumns(block.blockCol);
      var colMois = block.blockCol;
      var colPluieBilan = block.blockCol + 1;
      var colArrosageBilan = block.blockCol + 2;
      var colCumul = block.blockCol + 3;

      // Retrouve la ligne d'en-tête du bilan déjà créé (recherche "Mois" à partir de la
      // zone estimée par ajouterBilanMensuel, en descendant si besoin).
      var headerRow2 = -1;
      var scanFrom = block.headerRow + 30;
      for (var r = scanFrom; r < scanFrom + 60; r++) {
        var v = sheet.getRange(r, colMois).getValue();
        if (v && v.toString().trim().toLowerCase() === "mois") { headerRow2 = r; break; }
      }
      if (headerRow2 === -1) {
        results.push(p.sheet + " Ligne " + ligneIndex + " : tableau de bilan introuvable, rien à corriger.");
        return;
      }

      var colDateLettre = columnToLetter(cols.date);
      var colPluieLettre = columnToLetter(cols.pluie);
      var colMmLettre = columnToLetter(cols.mm);
      var plageDebut = block.dataStartRow;
      var plageFin = block.dataStartRow + 300;
      var dateRange = colDateLettre + plageDebut + ":" + colDateLettre + plageFin;
      var pluieRange = colPluieLettre + plageDebut + ":" + colPluieLettre + plageFin;
      var mmRange = colMmLettre + plageDebut + ":" + colMmLettre + plageFin;

      mois.forEach(function (m, idx) {
        var r = headerRow2 + 1 + idx;
        var debut = "DATE(2026;" + m.num + ";1)";
        var finExclusive = "DATE(2026;" + (m.num + 1) + ";1)"; // 1er du mois suivant (borne exclusive)

        sheet.getRange(r, colPluieBilan).setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + pluieRange + ";0))");
        sheet.getRange(r, colArrosageBilan).setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + mmRange + ";0))");
        sheet.getRange(r, colCumul).setFormula("=" + columnToLetter(colPluieBilan) + r + "+" + columnToLetter(colArrosageBilan) + r);
      });

      var totalRow = headerRow2 + 1 + mois.length;
      var firstDataRow2 = headerRow2 + 1;
      var lastDataRow2 = headerRow2 + mois.length;
      sheet.getRange(totalRow, colPluieBilan).setFormula("=SUM(" + columnToLetter(colPluieBilan) + firstDataRow2 + ":" + columnToLetter(colPluieBilan) + lastDataRow2 + ")");
      sheet.getRange(totalRow, colArrosageBilan).setFormula("=SUM(" + columnToLetter(colArrosageBilan) + firstDataRow2 + ":" + columnToLetter(colArrosageBilan) + lastDataRow2 + ")");
      sheet.getRange(totalRow, colCumul).setFormula("=SUM(" + columnToLetter(colCumul) + firstDataRow2 + ":" + columnToLetter(colCumul) + lastDataRow2 + ")");

      results.push(p.sheet + " Ligne " + ligneIndex + " : formules corrigées (tableau trouvé à la ligne " + headerRow2 + ").");
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

/* ---------- Mise en forme ponctuelle : arrondir l'affichage des données ---------- */
// À exécuter une seule fois. Applique un format "0" (arrondi à l'entier, sans décimale
// affichée) sur les colonnes de données de chaque bloc (Durée, Compteur début, Compteur
// fin, Débit, M3, mm, Pluie) ainsi que sur les tableaux de bilan mensuel (Pluie/Arrosage/
// cumul/Besoins), pour les 4 onglets. Ne touche ni aux dates, ni aux heures (qui gardent
// leur propre format), ni aux paramètres de parcelle (Surface/Longueur/Vitesse), pour ne
// pas perdre leur précision. Ceci change uniquement l'AFFICHAGE, pas les valeurs stockées
// ni les formules — les calculs restent faits avec la pleine précision.
function arrondirAffichageDonnees() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var cols = blockColumns(block.blockCol);
      var plageFin = block.dataStartRow + 300;
      var nbLignes = plageFin - block.dataStartRow + 1;

      // Colonnes numériques du bloc principal (Date et Heures gardent leur propre format)
      [cols.duree, cols.debut, cols.fin, cols.debit, cols.m3, cols.mm, cols.pluie].forEach(function (c) {
        sheet.getRange(block.dataStartRow, c, nbLignes, 1).setNumberFormat("0");
      });

      // Tableau de bilan mensuel associé, si présent (repéré par le mot "Mois")
      var colMois = block.blockCol;
      for (var r = block.headerRow; r < block.headerRow + 90; r++) {
        var v = sheet.getRange(r, colMois).getValue();
        if (v && v.toString().trim().toLowerCase() === "mois") {
          var colPluieBilan = block.blockCol + 1;
          sheet.getRange(r + 1, colPluieBilan, 6, 4).setNumberFormat("0");
          break;
        }
      }

      results.push(p.sheet + " Ligne " + ligneIndex + " : affichage arrondi appliqué.");
    });
  });

  Logger.log(results.join("\n"));
  return results;
}

/* ---------- Remplacement ponctuel : bilan mensuel de La Fontaine Ligne 1 ---------- */
// À exécuter une seule fois. Repère le tableau construit à la main sous La Fontaine
// Ligne 1 (recherche du mot "Mois") et le remplace entièrement par la version générée
// automatiquement (mêmes formules SUMPRODUCT+IFERROR, mêmes couleurs) que les 10 autres
// blocs. Écrase tout ce qui se trouvait dans cette zone (en-têtes, valeurs, formules).
function remplacerBilanFontaineLigne1() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];

  var mois = [
    { nom: "Mai", num: 5, besoin: 59 },
    { nom: "Juin", num: 6, besoin: 78 },
    { nom: "Juillet", num: 7, besoin: 94 },
    { nom: "Août", num: 8, besoin: 85 },
    { nom: "Septembre", num: 9, besoin: 55 }
  ];
  var COULEURS_LIGNE = { 1: "#FCE5CD", 2: "#CFE2F3", 3: "#D9EAD3" };
  var COULEUR_HEADER_BG = "#073763";
  var COULEUR_HEADER_TXT = "#FFFFFF";

  var sheet = ss.getSheetByName("La Fontaine");
  if (!sheet) { Logger.log("La Fontaine : onglet introuvable"); return; }

  var ligneIndex = 1;
  var block = findLigneBlock(sheet, ligneIndex);
  if (!block) { Logger.log("La Fontaine Ligne 1 : bloc introuvable"); return; }

  var cols = blockColumns(block.blockCol);
  var colMois = block.blockCol;
  var colPluieBilan = block.blockCol + 1;
  var colArrosageBilan = block.blockCol + 2;
  var colCumul = block.blockCol + 3;
  var colBesoins = block.blockCol + 4;

  // Repère le tableau existant (recherche "Mois" à partir de la zone estimée)
  var headerRow2 = -1;
  var scanFrom = block.headerRow + 30;
  for (var r = scanFrom; r < scanFrom + 60; r++) {
    var v = sheet.getRange(r, colMois).getValue();
    if (v && v.toString().trim().toLowerCase() === "mois") { headerRow2 = r; break; }
  }
  if (headerRow2 === -1) {
    Logger.log("La Fontaine Ligne 1 : tableau existant introuvable, rien remplacé.");
    return;
  }

  // Efface entièrement l'ancienne zone (valeurs + mise en forme) avant de reconstruire
  var zone = sheet.getRange(headerRow2, colMois, 7, 5);
  zone.clearContent();
  zone.clearFormat();

  var headerRange = sheet.getRange(headerRow2, colMois, 1, 5);
  headerRange.setValues([["Mois", "Pluie (mm)", "Arrosage (mm)", "cumul pluie+arrosage", "Besoins de plante, sans pluie Etc (mm) = Kc x ETP"]]);
  headerRange.setBackground(COULEUR_HEADER_BG).setFontColor(COULEUR_HEADER_TXT).setFontWeight("bold");

  var couleurLigne = COULEURS_LIGNE[ligneIndex];
  var colDateLettre = columnToLetter(cols.date);
  var colPluieLettre = columnToLetter(cols.pluie);
  var colMmLettre = columnToLetter(cols.mm);
  var plageDebut = block.dataStartRow;
  var plageFin = block.dataStartRow + 300;
  var dateRange = colDateLettre + plageDebut + ":" + colDateLettre + plageFin;
  var pluieRange = colPluieLettre + plageDebut + ":" + colPluieLettre + plageFin;
  var mmRange = colMmLettre + plageDebut + ":" + colMmLettre + plageFin;

  mois.forEach(function (m, idx) {
    var r = headerRow2 + 1 + idx;
    var debut = "DATE(2026;" + m.num + ";1)";
    var finExclusive = "DATE(2026;" + (m.num + 1) + ";1)";

    sheet.getRange(r, colMois).setValue(m.nom);

    var pluieCell = sheet.getRange(r, colPluieBilan);
    pluieCell.setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + pluieRange + ";0))");
    pluieCell.setBackground(couleurLigne).setNumberFormat("0");

    var arrosageCell = sheet.getRange(r, colArrosageBilan);
    arrosageCell.setFormula("=SUMPRODUCT(IFERROR((" + dateRange + ">=" + debut + ")*(" + dateRange + "<" + finExclusive + ")*" + mmRange + ";0))");
    arrosageCell.setBackground(couleurLigne).setNumberFormat("0");

    var cumulCell = sheet.getRange(r, colCumul);
    cumulCell.setFormula("=" + columnToLetter(colPluieBilan) + r + "+" + columnToLetter(colArrosageBilan) + r).setNumberFormat("0");

    sheet.getRange(r, colBesoins).setValue(m.besoin).setNumberFormat("0");
  });

  var totalRow = headerRow2 + 1 + mois.length;
  var firstDataRow2 = headerRow2 + 1;
  var lastDataRow2 = headerRow2 + mois.length;

  sheet.getRange(totalRow, colMois).setValue("TOTAL").setFontWeight("bold");
  sheet.getRange(totalRow, colPluieBilan).setFormula("=SUM(" + columnToLetter(colPluieBilan) + firstDataRow2 + ":" + columnToLetter(colPluieBilan) + lastDataRow2 + ")").setFontWeight("bold").setNumberFormat("0");
  sheet.getRange(totalRow, colArrosageBilan).setFormula("=SUM(" + columnToLetter(colArrosageBilan) + firstDataRow2 + ":" + columnToLetter(colArrosageBilan) + lastDataRow2 + ")").setFontWeight("bold").setNumberFormat("0");
  sheet.getRange(totalRow, colCumul).setFormula("=SUM(" + columnToLetter(colCumul) + firstDataRow2 + ":" + columnToLetter(colCumul) + lastDataRow2 + ")").setFontWeight("bold").setNumberFormat("0");
  sheet.getRange(totalRow, colBesoins).setFormula("=SUM(" + columnToLetter(colBesoins) + firstDataRow2 + ":" + columnToLetter(colBesoins) + lastDataRow2 + ")").setFontWeight("bold").setNumberFormat("0");

  Logger.log("La Fontaine Ligne 1 : ancien tableau (ligne " + headerRow2 + ") remplacé par la version générée.");
}

/* ---------- Code couleur ponctuel : Cumul / Besoins des tableaux de bilan ---------- */
// À exécuter une seule fois. Applique une mise en forme conditionnelle DYNAMIQUE (elle se
// recalcule automatiquement si les données changent) sur les colonnes Cumul et Besoins de
// chacun des 11 tableaux de bilan :
//  - Vert  : cumul >= besoins
//  - Jaune : cumul entre 80% et 100% du besoin (déficit de 0 à 20%)
//  - Rouge : cumul < 80% du besoin (déficit de plus de 20%)
function appliquerCodeCouleurBilan() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var results = [];
  var VERT = "#B7E1CD", JAUNE = "#FCE8B2", ROUGE = "#F4C7C3";

  PARCELLES.forEach(function (p) {
    var sheet = ss.getSheetByName(p.sheet);
    if (!sheet) { results.push(p.sheet + " : onglet introuvable"); return; }

    p.lignes.forEach(function (ligneIndex) {
      var block = findLigneBlock(sheet, ligneIndex);
      if (!block) { results.push(p.sheet + " Ligne " + ligneIndex + " : bloc introuvable"); return; }

      var colMois = block.blockCol;
      var headerRow2 = -1;
      for (var r = block.headerRow; r < block.headerRow + 90; r++) {
        var v = sheet.getRange(r, colMois).getValue();
        if (v && v.toString().trim().toLowerCase() === "mois") { headerRow2 = r; break; }
      }
      if (headerRow2 === -1) { results.push(p.sheet + " Ligne " + ligneIndex + " : tableau de bilan introuvable"); return; }

      var colCumulBilan = block.blockCol + 3;
      var colBesoinsBilan = block.blockCol + 4;
      var firstDataRow2 = headerRow2 + 1;

      var rangeCumul = sheet.getRange(firstDataRow2, colCumulBilan, 6, 1);   // colonne Cumul seule
      var rangeBesoins = sheet.getRange(firstDataRow2, colBesoinsBilan, 6, 1); // colonne Besoins seule

      // Retire TOUTE règle déjà posée touchant cette zone (peu importe sa largeur — ça
      // couvre aussi bien les nouvelles règles à 1 colonne que les résidus d'une éventuelle
      // ancienne tentative à 2 colonnes), pour ne jamais laisser une ancienne règle prioritaire
      // masquer les nouvelles.
      var rulesExistantes = sheet.getConditionalFormatRules();
      var rulesConservees = rulesExistantes.filter(function (rule) {
        return !rule.getRanges().some(function (rg) {
          return rg.getRow() === firstDataRow2 &&
                 (rg.getColumn() === colCumulBilan || rg.getColumn() === colBesoinsBilan);
        });
      });

      // Références relatives simples (sans $) : la plage appliquée n'ayant qu'une seule
      // colonne, la ligne s'ajuste automatiquement pour chaque cellule, comme un remplissage
      // vers le bas classique.
      var colCumulLettre = columnToLetter(colCumulBilan);
      var colBesoinsLettre = columnToLetter(colBesoinsBilan);
      var ref = firstDataRow2;

      var formuleVert = "=" + colCumulLettre + ref + ">=" + colBesoinsLettre + ref;
      var formuleJaune = "=AND(" + colCumulLettre + ref + "<" + colBesoinsLettre + ref + ";" + colCumulLettre + ref + ">=" + colBesoinsLettre + ref + "*0.8)";
      var formuleRouge = "=" + colCumulLettre + ref + "<" + colBesoinsLettre + ref + "*0.8";

      [rangeCumul, rangeBesoins].forEach(function (rng) {
        var ruleVert = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formuleVert).setBackground(VERT).setRanges([rng]).build();
        var ruleJaune = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formuleJaune).setBackground(JAUNE).setRanges([rng]).build();
        var ruleRouge = SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied(formuleRouge).setBackground(ROUGE).setRanges([rng]).build();
        rulesConservees.push(ruleVert, ruleJaune, ruleRouge);
      });

      sheet.setConditionalFormatRules(rulesConservees);

      results.push(p.sheet + " Ligne " + ligneIndex + " : code couleur appliqué (lignes " + firstDataRow2 + " à " + (firstDataRow2 + 5) + ").");
    });
  });

  Logger.log(results.join("\n"));
  return results;
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
  var scanRows = 6;
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
  var maxScan = 300;
  for (var rowI = dataStartRow; rowI < dataStartRow + maxScan; rowI++) {
    var m3v = sheet.getRange(rowI, colM3).getValue();
    var dv = sheet.getRange(rowI, colDate).getValue();
    var m3Blank = (m3v === "" || m3v === null);
    var dateBlank = (dv === "" || dv === null);
    if (m3Blank && dateBlank) return rowI;
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
