/**
 * ============================================================
 *  BACKEND — Google Apps Script
 *  Reçoit les données de l'application web et les ajoute
 *  dans une Google Sheet stockée dans ton Google Drive.
 * ============================================================
 */

// En-têtes de la feuille (créés automatiquement à la 1ère utilisation)
var HEADERS = [
  "Horodatage", "Nom", "Prénom", "École", "Série",
  // Notes brutes par matière (2nde / 1ère / Tle / Bac)
  "Math 2nde", "Math 1ère", "Math Tle", "Math Bac",
  "Phys 2nde", "Phys 1ère", "Phys Tle", "Phys Bac",
  "Fr 2nde", "Fr 1ère", "Fr Tle", "Fr Bac",
  "Ang 2nde", "Ang 1ère", "Ang Tle", "Ang Bac",
  "SVT 2nde", "SVT 1ère", "SVT Tle", "SVT Bac",
  "MT 2nde", "MT 1ère", "MT Tle", "MT Bac",
  "ScEco 2nde", "ScEco 1ère", "ScEco Tle", "ScEco Bac",
  // Moyennes générales par matière
  "MGM Math", "MGM Phys", "MGM Fr", "MGM Ang", "MGM SVT", "MGM MT", "MGM ScEco",
  // Résultats par filière (JSON) + meilleure filière
  "Résultats (JSON)", "Meilleure filière", "Meilleure moyenne"
];

var SUBJECT_ORDER = ["math", "phys", "fr", "ang", "svt", "mt", "sceco"];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = getSheet_();

    var row = [
      new Date(),
      clean_(data.nom),
      clean_(data.prenom),
      clean_(data.ecole),
      clean_(data.serie)
    ];

    // Notes brutes : 4 colonnes par matière, vide si la matière
    // n'existe pas pour cette série
    SUBJECT_ORDER.forEach(function (s) {
      var n = (data.notes && data.notes[s]) || {};
      row.push(num_(n.n2), num_(n.n1), num_(n.tle), num_(n.bac));
    });

    // Moyennes générales par matière
    SUBJECT_ORDER.forEach(function (s) {
      row.push(data.mgm && data.mgm[s] != null ? round2_(data.mgm[s]) : "");
    });

    // Résultats par filière
    var res = data.resultats || [];
    var best = res.length ? res[0] : { filiere: "", moyenne: "" };
    row.push(JSON.stringify(res), best.filiere, best.moyenne);

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Permet de tester l'URL dans un navigateur
function doGet() {
  return ContentService.createTextOutput("API INP-HB : service actif ✓");
}

/* ---------- utilitaires ---------- */

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