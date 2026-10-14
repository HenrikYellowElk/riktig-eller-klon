// Resultatmottagare för "Riktig eller klon" (Google Apps Script, webbapp).
// Installation (Henrik, ≈2 minuter):
// 1. Gå till https://script.google.com, Nytt projekt, klistra in denna fil som Code.gs.
// 2. Distribuera → Ny distribution → Typ: Webbapp → Kör som: Jag → Åtkomst: Alla → Distribuera.
// 3. Kopiera webbappens URL (slutar på /exec) och klistra in den i web_demo/clips.json som "resultsUrl".
// Resultaten hamnar i ett kalkylark "Riktig eller klon – resultat" i din Drive (skapas automatiskt).

const SHEET_NAME = 'Riktig eller klon – resultat';

function sheet_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('sheetId');
  let ss = id ? SpreadsheetApp.openById(id) : null;
  if (!ss) { ss = SpreadsheetApp.create(SHEET_NAME); props.setProperty('sheetId', ss.getId()); }
  const sh = ss.getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(['time', 'name', 'score', 'answered', 'total', 'clone', 'hidden', 'rows']);
  return sh;
}

function doPost(e) {
  const rec = JSON.parse(e.postData.contents);
  const sh = sheet_();
  sh.appendRow([rec.time, String(rec.name).slice(0, 40), rec.score, rec.answered, rec.total, rec.clone || '', !!rec.hidden, JSON.stringify(rec.rows || [])]);
  return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const sh = sheet_();
  const rows = sh.getDataRange().getValues().slice(1);
  const results = rows.map(r => ({time: r[0], name: r[1], score: r[2], answered: r[3], total: r[4], clone: r[5], hidden: r[6] === true || r[6] === 'TRUE'}));
  return ContentService.createTextOutput(JSON.stringify({results: results.slice(-500)})).setMimeType(ContentService.MimeType.JSON);
}
