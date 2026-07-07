/**
 * ============================================================
 *  BACKEND — Google Apps Script (v3)
 *  Anti-doublons : si nom + prénom + école + série existent
 *  déjà, la ligne est mise à jour au lieu d'être dupliquée,
 *  et le compteur "Nb de calculs" est incrémenté.
 * ============================================================
 */

var HEADERS = [
  "Horodatage", "Nom", "Prénom", "École", "Série",
  "Math 2nde", "Math 1ère", "Math Tle", "Math Bac",
  "Phys 2nde", "Phys 1ère", "Phys Tle", "Phys Bac",
  "Fr 2nde", "Fr 1ère", "Fr Tle", "Fr Bac",
  "Ang 2nde", "Ang 1ère", "Ang Tle", "Ang Bac",
  "SVT 2nde", "SVT 1ère", "SVT Tle", "SVT Bac",
  "MT 2nde", "MT 1ère", "MT Tle", "MT Bac",
  "ScEco 2nde", "ScEco 1ère", "ScEco Tle", "ScEco Bac",
  "MGM Math", "MGM Phys", "MGM Fr", "MGM Ang", "MGM SVT", "MGM MT", "MGM ScEco",
  "Résultats (JSON)", "Meilleure filière", "Meilleure moyenne",
  "Nb de calculs"
];

var SUBJECT_ORDER = ["math", "phys", "fr", "ang", "svt", "mt", "sceco"];

function doPost(e) {
  // Verrou : évite les conflits si deux étudiants envoient en même temps
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    // Construire la ligne de données (sans le compteur)
    var row = [
      new Date(),
      clean_(data.nom),
      clean_(data.prenom),
      clean_(data.ecole),
      clean_(data.serie)
    ];
    SUBJECT_ORDER.forEach(function (s) {
      var n = (data.notes && data.notes[s]) || {};
      row.push(num_(n.n2), num_(n.n1), num_(n.tle), num_(n.bac));
    });
    SUBJECT_ORDER.forEach(function (s) {
      row.push(data.mgm && data.mgm[s] != null ? round2_(data.mgm[s]) : "");
    });
    var res = data.resultats || [];
    var best = res.length ? res[0] : { filiere: "", moyenne: "" };
    row.push(JSON.stringify(res), best.filiere, best.moyenne);

    // --- Recherche d'un doublon : nom + prénom + école + série ---
    var key = norm_(data.nom) + "|" + norm_(data.prenom) + "|" +
              norm_(data.ecole) + "|" + norm_(data.serie);
    var existingRow = findRow_(sheet, key);

    if (existingRow > 0) {
      // Déjà présent : on met à jour la ligne et on incrémente le compteur
      var countCell = sheet.getRange(existingRow, HEADERS.length);
      var count = Number(countCell.getValue()) || 1;
      row.push(count + 1);
      sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    } else {
      // Nouvel étudiant : nouvelle ligne, compteur = 1
      row.push(1);
      sheet.appendRow(row);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok", updated: existingRow > 0 }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// Permet de tester l'URL dans un navigateur
function doGet() {
  return ContentService.createTextOutput("API INP-HB : service actif ✓ (v3 anti-doublons)");
}

/* ---------- utilitaires ---------- */

// Cherche une ligne existante avec la même clé. Retourne son numéro, ou 0.
function findRow_(sheet, key) {
  var last = sheet.getLastRow();
  if (last < 2) return 0;
  // Colonnes 2 à 5 : Nom, Prénom, École, Série
  var values = sheet.getRange(2, 2, last - 1, 4).getValues();
  for (var i = 0; i < values.length; i++) {
    var k = norm_(values[i][0]) + "|" + norm_(values[i][1]) + "|" +
            norm_(values[i][2]) + "|" + norm_(values[i][3]);
    if (k === key) return i + 2; // +2 : ligne d'en-tête + index qui part de 0
  }
  return 0;
}

// Normalise un texte : MAJUSCULES, sans accents, espaces multiples réduits
// → "  Lycée Simone  Ehivet " devient "LYCEE SIMONE EHIVET"
function norm_(v) {
  return String(v == null ? "" : v)
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .toUpperCase()
    .replace(/\s+/g, " ")
    .trim();
}

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Etudiants") || ss.insertSheet("Etudiants");
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function clean_(v) {
  return String(v == null ? "" : v).substring(0, 200);
}

function num_(v) {
  var n = parseFloat(v);
  return isNaN(n) ? "" : n;
}

function round2_(v) {
  return Math.round(parseFloat(v) * 100) / 100;
}