// These APIs run only when a button is clicked. Nothing polls or recalculates in the background.
const codes = ["ALL", "PL", "NPAT", "PBT", "EBIT", "EBITDA", "GP", "REV", "200", "REV.OTH", "260", "270", "EXP.COS", "300", "310"];
const statusElement = document.getElementById("status");
const output = document.getElementById("output");
const createButton = document.getElementById("create");
const readButton = document.getElementById("read");

async function createSheet() {
  return Excel.run(async context => {
    const sheet = context.workbook.worksheets.add();
    // Preserve the original case for code 300, identically in both copies.
    const formulas = codes.map(code => [`=REPRO.MEMBER("${code === "300" ? "account" : "Account"}","${code}")`]);
    sheet.getRange("B4:B18").formulas = formulas;
    sheet.getRange("B20:B34").formulas = formulas;
    sheet.getRange("B:B").format.columnWidth = 160;
    sheet.activate();
    sheet.getRange("B4").select();
    sheet.load("name");
    await context.sync();
    return `Created ${sheet.name}: B4:B18 and B20:B34. Recalculate with Ctrl+Alt+F9.`;
  });
}

async function readCells() {
  return Excel.run(async context => {
    const range = context.workbook.worksheets.getActiveWorksheet().getRange("B4:B34");
    range.load(["address", "formulas", "valuesAsJson"]);
    await context.sync();
    const rows = range.valuesAsJson.map((row, index) => ({
      cell: `B${index + 4}`,
      formula: range.formulas[index][0],
      value: row[0],
    })).filter(row => row.formula);
    const busy = rows.filter(row => row.value.errorType === "Busy").map(row => row.cell);
    const snapshot = {
      capturedAt: new Date().toISOString(),
      host: Office.context.diagnostics,
      address: range.address,
      busy,
      rows,
    };
    output.value = JSON.stringify(snapshot, null, 2);
    return busy.length ? `Busy cells: ${busy.join(", ")}. Copy the JSON below.` : "No Busy cells in this snapshot.";
  });
}

async function run(action) {
  createButton.disabled = readButton.disabled = true;
  statusElement.textContent = "Working...";
  try {
    statusElement.textContent = await action();
  } catch (error) {
    statusElement.textContent = String(error);
  } finally {
    createButton.disabled = readButton.disabled = false;
  }
}

Office.onReady(info => {
  if (info.host !== Office.HostType.Excel) {
    statusElement.textContent = "Open this add-in inside Excel.";
    return;
  }
  statusElement.textContent = `Ready. Excel ${Office.context.diagnostics.version}. Open repro.xlsx on Sheet2, or create a new reproduction sheet.`;
  createButton.disabled = readButton.disabled = false;
  createButton.onclick = () => run(createSheet);
  readButton.onclick = () => run(readCells);
});
