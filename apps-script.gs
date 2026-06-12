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
 *   6) Vytvoř list "Kategorie" (sloupec A: věkovky), pak v editoru spusť
 *      jednou funkci setupCategorySheets() – předpřipraví taby pro každou věkovku.
 *
 * Zápis: každá jízda jde do "Results" (kompletní log) a navíc se každý
 * závodník zrcadlí do tabu pojmenovaného podle jeho věkové kategorie.
 *
 * POZOR: po každé změně kódu udělej Deploy → Manage deployments → edit
 *        → Version: New version, jinak běží pod stejným URL stará verze.
 */

const STARTOVKA_SHEET = "Startovka";
const RESULTS_SHEET    = "Results";       // kompletní log (audit + idempotence)
const KATEGORIE_SHEET  = "Kategorie";     // seznam věkovek (zdroj pro taby i dropdown)
const SHARED_SECRET    = "rUOQe4tV52N0lZEwL7g986sb4U0escic"; // musí se shodovat s index.html

// Hlavička listu Results (vytvoří se automaticky, když chybí)
const RESULTS_HEADER = [
  "Čas zápisu", "Startovní číslo", "Barva", "Jméno", "Kategorie", "Čas", "Čas (ms)", "ID jízdy"
];

// Sloupec s ID jízdy (1-based) – musí sedět s pozicí v RESULTS_HEADER
const ID_COL = 8;

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
      { number: String(data.rider1).trim(), time: data.time1, ms: data.ms1, color: data.color1 || "" },
      { number: String(data.rider2).trim(), time: data.time2, ms: data.ms2, color: data.color2 || "" }
    ];

    const entries = riders.map(rider => {
      const found = findRider(startovkaData, rider.number, rider.color);
      return {
        category: found.category,
        row: [
          now,
          rider.number,
          rider.color,
          found.name,
          found.category,
          rider.time,
          rider.ms,
          data.id || ""
        ]
      };
    });

    // 1) Kompletní log do Results (oba řádky najednou, atomicky vůči zámku)
    const allRows = entries.map(e => e.row);
    resultsSheet
      .getRange(resultsSheet.getLastRow() + 1, 1, allRows.length, allRows[0].length)
      .setValues(allRows);

    // 2) Zrcadlo do tabu podle věkové kategorie (každý závodník do své kategorie)
    appendToCategorySheets(ss, entries);

    return jsonOut({ status: "OK", written: allRows.length });
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

// Seskupí řádky podle kategorie a zapíše je do příslušných tabů.
function appendToCategorySheets(ss, entries) {
  const byCat = {};
  entries.forEach(e => {
    const cat = e.category || "NEZNÁMÁ KATEGORIE";
    (byCat[cat] = byCat[cat] || []).push(e.row);
  });
  Object.keys(byCat).forEach(cat => {
    const sheet = getOrCreateCategorySheet(ss, cat);
    const rows = byCat[cat];
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);

    // Po každém zápisu tab přeřadíme podle Čas (ms) vzestupně (od nejrychlejšího),
    // ať zůstane seřazený a nerozsype se. Hlavičku (řádek 1) neřadíme.
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, RESULTS_HEADER.length)
        .sort({ column: 7, ascending: true });   // sloupec 7 = Čas (ms)
    }
  });
}

// Tab pro danou kategorii (název = kategorie). Vytvoří se i s hlavičkou, když chybí.
function getOrCreateCategorySheet(ss, category) {
  const name = sheetNameForCategory(category);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(RESULTS_HEADER);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Název listu nesmí obsahovat : \ / ? * [ ] a má limit 100 znaků.
function sheetNameForCategory(category) {
  let name = String(category || "NEZNÁMÁ KATEGORIE")
    .replace(/[:\\\/?*\[\]]/g, " ").trim().slice(0, 100);
  return name || "NEZNÁMÁ KATEGORIE";
}

/**
 * Jednorázově (nebo kdykoliv) spusť ručně z editoru Apps Scriptu.
 * Z listu `Kategorie` (sloupec A, od 2. řádku) předpřipraví prázdné taby
 * s hlavičkou pro každou věkovku, ať existují i před prvním zápisem.
 */
function setupCategorySheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const katSheet = ss.getSheetByName(KATEGORIE_SHEET);
  if (!katSheet) {
    throw new Error("Chybí list '" + KATEGORIE_SHEET + "' se seznamem věkovek.");
  }
  const last = katSheet.getLastRow();
  if (last < 2) return;
  const cats = katSheet.getRange(2, 1, last - 1, 1).getValues();
  cats.forEach(r => {
    const cat = String(r[0]).trim();
    if (cat) getOrCreateCategorySheet(ss, cat);
  });
}

function alreadyStored(resultsSheet, id) {
  const lastRow = resultsSheet.getLastRow();
  if (lastRow < 2) return false;
  const ids = resultsSheet.getRange(2, ID_COL, lastRow - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return true;
  }
  return false;
}

// Startovka: Startovní číslo | Barva | Jméno | Kategorie
// Číslo není unikátní napříč barvami (žlutá 50 ≠ bílá 50) – matchujeme obojí.
function findRider(startovkaData, riderNumber, riderColor) {
  const color = normColor(riderColor);
  for (let i = 1; i < startovkaData.length; i++) {
    const number = String(startovkaData[i][0]).trim();
    const rowColor = normColor(startovkaData[i][1]);
    if (number === riderNumber && rowColor === color) {
      return { name: startovkaData[i][2], category: startovkaData[i][3] };
    }
  }
  return { name: "NEZNÁMÝ ZÁVODNÍK", category: "NEZNÁMÁ KATEGORIE" };
}

// Sjednotí zápis barvy (diakritika/velikost/mezery), ať lookup nezáleží na formátu.
function normColor(c) {
  return String(c || "").trim().toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
