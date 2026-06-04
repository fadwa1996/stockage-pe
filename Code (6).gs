var SHEET_ID = '1qurZM2b4xEJg-UtVtiAgtkvZjKypa2j-Db-YpUeSbcs';

function doGet(e) {
  var type     = (e && e.parameter && e.parameter.type)     || 'commandes';
  var callback = (e && e.parameter && e.parameter.callback) || 'callback';
  var data = [];
  try {
    if      (type === 'stocks')       { data = getStocks(); }
    else if (type === 'commandes')    { data = getCommandes(); }
    else if (type === 'pp')           { data = getGenericStock('Stock PRODUIT FINI PP'); }
    else if (type === 'monofilament') { data = getMonofilament(); }
  } catch (err) { data = []; }
  var output = ContentService.createTextOutput(callback + '(' + JSON.stringify(data) + ')');
  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}

function getCommandes() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Commandes') || ss.getSheets()[1] || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    var rows  = sheet.getDataRange().getValues();
    var result = [];
    // En-têtes ligne 3 (index 2), données à partir de ligne 4 (index 3)
    // C=index 2 (Date), D=index 3 (Designation), E=index 4 (Qte), F=index 5 (Reste)
    for (var i = 3; i < rows.length; i++) {
      var row  = rows[i];
      var date  = trim(String(row[2]));
      var desig = trim(String(row[3]));
      var qte   = parseNum(row[4]);
      var reste = parseNum(row[5]);
      if (!desig && !qte) continue;
      result.push({ client: '', date: formatDateVal(row[2]), desig: desig, qte: qte, reste: reste });
    }
    return result;
  } catch(err) { return []; }
}

function getStocks() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Stocks PE') || ss.getSheetByName('Stocks') || ss.getSheets()[0];
    var lastRow = sheet.getLastRow();
    if (lastRow < 7) return [];
    var rows  = sheet.getDataRange().getValues();
    var result = [];
    for (var i = 6; i < rows.length; i++) {
      var row      = rows[i];
      var ref      = trim(String(row[3]));
      var nbr      = parseNum(row[4]);
      var poids    = parseNum(row[5]);
      var remarque = trim(String(row[6] || ''));
      if (!ref) continue;
      result.push({ ref: ref, nbr: nbr, poids: poids, min: 60, remarque: remarque });
    }
    return result;
  } catch(err) { return []; }
}

function getGenericStock(sheetName) {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 5) return [];
    var rows = sheet.getDataRange().getValues();
    var result = [];
    // En-têtes ligne 4 (index 3), données à partir de ligne 5 (index 4)
    // C=index 2 (Designation), D=index 3 (Qte), E=index 4 (Poids)
    for (var i = 4; i < rows.length; i++) {
      var row  = rows[i];
      var ref  = trim(String(row[2]));
      var nbr  = parseNum(row[3]);
      var poids = parseNum(row[4]);
      if (!ref) continue;
      result.push({ ref: ref, nbr: nbr, poids: poids });
    }
    return result;
  } catch(err) { return []; }
}

function getMonofilament() {
  try {
    var ss    = SpreadsheetApp.openById(SHEET_ID);
    var sheet = ss.getSheetByName('Stock Monofilament');
    if (!sheet) return [];
    var lastRow = sheet.getLastRow();
    if (lastRow < 4) return [];
    var rows = sheet.getDataRange().getValues();
    var result = [];
    // Détecte les sous-rubriques (lignes fusionnées/titre) vs lignes de données
    // Col C=index 2, D=index 3, E=index 4
    for (var i = 3; i < rows.length; i++) {
      var row  = rows[i];
      var colC = trim(String(row[2]));
      var colD = row[3];
      var colE = row[4];
      if (!colC) continue;
      // Ligne sous-rubrique : pas de valeur en D et E (ou D/E = 0 et texte long en C)
      var isSub = (!parseNum(colD) && !parseNum(colE)) || (colC.length > 20 && !parseNum(colD));
      if (isSub) {
        result.push({ ref: colC, nbr: 0, poids: 0, isSub: true });
      } else {
        result.push({ ref: colC, nbr: parseNum(colD), poids: parseNum(colE), isSub: false });
      }
    }
    return result;
  } catch(err) { return []; }
}

function trim(s) { return String(s || '').trim(); }

function parseNum(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/\s/g,'').replace(',','.'));
  return isNaN(n) ? 0 : n;
}

function isDateLike(v) {
  if (v instanceof Date) return true;
  return /^\d{1,2}[\/\-\.]\d{1,2}[\/\-\.]\d{4}$/.test(String(v).trim());
}

function formatDateVal(v) {
  if (!v) return '';
  if (v instanceof Date) {
    var d = v.getDate();
    var m = v.getMonth() + 1;
    var y = v.getFullYear();
    return (d < 10 ? '0' + d : d) + '/' + (m < 10 ? '0' + m : m) + '/' + y;
  }
  return String(v).trim();
}
