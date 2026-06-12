/**
 * Sprint Timer – backend (Google Apps Script Web App)
 *
 * Nasazení:
 *   1) Otevři Google Sheet → Rozšíření → Apps Script
 *   2) Vlož tento kód, ulož
 *   3) Deploy → New deployment → typ "Web app"
 *        Execute as: Me
 *        Who has access: Anyone
 *   4) Zkopíruj /exec URL do index.html (GOOGLE_SCRIPT_URL)
 *   5) Nastav SHARED_SECRET shodně s index.html
 *
 * POZOR: po každé změně kódu udělej Deploy → Manage deployments → edit
 *        → Version: New version, jinak běží pod stejným URL stará verze.
 */

const STARTOVKA_SHEET = "Startovka";
const RESULTS_SHEET    = "Results";
const SHARED_SECRET    = "SEM_DEJ_TAJNY_RETEZEC"; // musí se shodovat s index.html

// Hlavička listu Results (vytvoří se automaticky, když chybí)
const RESULTS_HEADER = [
  "Čas zápisu", "Startovní číslo", "Jméno", "Kategorie", "Čas", "Čas (ms)", "ID jízdy"
];

/** Test v prohlížeči – ověří, že je nasazení živé. */
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "OK", service: "Sprint Timer" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  // Souběžné zápisy z více mobilů serializujeme zámkem
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const data = JSON.parse(e.postData.contents);

    if (data.secret !== SHARED_SECRET) {
      return jsonOut({ status: "ERROR", error: "unauthorized" });
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const startovkaSheet = ss.getSheetByName(STARTOVKA_SHEET);
    const resultsSheet   = getOrCreateResults(ss);

    // Idempotence: pokud už řádek s tímto ID existuje, podruhé nezapisujeme.
    // (Frontend posílá v no-cors a může poslat duplicitně při retry.)
    if (data.id && alreadyStored(resultsSheet, data.id)) {
      return jsonOut({ status: "OK", duplicate: true });
    }

    const startovkaData = startovkaSheet.getDataRange().getValues();
    const now = new Date();

    const riders = [
      { number: String(data.rider1).trim(), time: data.time1, ms: data.ms1 },
      { number: String(data.rider2).trim(), time: data.time2, ms: data.ms2 }
    ];

    const rows = riders.map(rider => {
      const found = findRider(startovkaData, rider.number);
      return [
        now,
        rider.number,
        found.name,
        found.category,
        rider.time,
        rider.ms,
        data.id || ""
      ];
    });

    // Oba řádky najednou – atomicky vůči zámku
    resultsSheet
      .getRange(resultsSheet.getLastRow() + 1, 1, rows.length, rows[0].length)
      .setValues(rows);

    return jsonOut({ status: "OK", written: rows.length });
  } catch (err) {
    return jsonOut({ status: "ERROR", error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getOrCreateResults(ss) {
  let sheet = ss.getSheetByName(RESULTS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(RESULTS_SHEET);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RESULTS_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function alreadyStored(resultsSheet, id) {
  const lastRow = resultsSheet.getLastRow();
  if (lastRow < 2) return false;
  // Sloupec G (7) = ID jízdy
  const ids = resultsSheet.getRange(2, 7, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return true;
  }
  return false;
}

function findRider(startovkaData, riderNumber) {
  for (let i = 1; i < startovkaData.length; i++) {
    const number = String(startovkaData[i][0]).trim();
    if (number === riderNumber) {
      return { name: startovkaData[i][1], category: startovkaData[i][2] };
    }
  }
  return { name: "NEZNÁMÝ ZÁVODNÍK", category: "NEZNÁMÁ KATEGORIE" };
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
