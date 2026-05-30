function processAdobeCSV() {

  const inputFolder = DriveApp.getFoldersByName("RAW_ADOBE").next();
  const outputFolder = DriveApp.getFoldersByName("PROCESSED_ADOBE").next();

  const files = inputFolder.getFiles();

  if (!files.hasNext()) {
    Logger.log("No hay archivos");
    return;
  }

  const file = files.next();
  const content = file.getBlob().getDataAsString();

  const parsed = parseAdobe(content);

  const rows = convertToRows(parsed);

  const sheet = SpreadsheetApp.create(file.getName() + "_processed");
  const ws = sheet.getActiveSheet();

  ws.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const driveFile = DriveApp.getFileById(sheet.getId());
  outputFolder.addFile(driveFile);
  DriveApp.getRootFolder().removeFile(driveFile);
}


/**
 * ===============================
 * PARSER
 * ===============================
 */


function parseAdobe(content) {

  const lines = content.split("\n");

  let output = [];

  let mode = "none";

  // 🔥 PANEL CONTEXT (NO se reinicia agresivamente)
  let panelName = "";
  let reportSuite = "";
  let date = "";
  let currentSegments = [];

  // 🔥 TABLE
  let currentTable = "";
  let headerRows = [];
  let dataStarted = false;

  for (let i = 0; i < lines.length; i++) {

    let line = lines[i]
      .replace(/"/g, "")
      .replace("\r", "")
      .trim();

    if (line === "") continue;

    // ===============================
    // PANEL BLOCK
    // ===============================
    if (line.startsWith("#=")) {
      mode = "panel";
      continue;
    }

    // ===============================
    // TABLE BLOCK
    // ===============================
    if (line.startsWith("###")) {
      mode = "table";
      continue;
    }

    // ===============================
    // CONTENT (#)
    // ===============================
    if (line.startsWith("#")) {

      let clean = line.replace(/^#+/, "").trim();

      if (clean === "" || clean.match(/^[-=]+$/)) continue;

      // -------- PANEL --------
      if (mode === "panel") {

        if (clean === "Panel" || clean === "Freeform") {
          panelName = clean;
          continue;
        }

        if (clean.startsWith("Report suite:")) {
          reportSuite = clean.replace("Report suite:", "").trim();
          continue;
        }

        if (clean.startsWith("Date:")) {
          date = clean.replace("Date:", "").trim();
          continue;
        }

        if (clean.startsWith("Segments:")) {

          let raw = clean.replace("Segments:", "").trim();

          currentSegments = raw
            .split(",")
            .map(s => s.trim())
            .filter(s => s !== "");

          continue;
        }
      }

      // -------- TABLE --------
      if (mode === "table") {

        currentTable = clean;

        headerRows = [];
        dataStarted = false;

        continue;
      }

      continue;
    }

    // ===============================
    // DATA
    // ===============================
    let row = line.split(",");
    let isData = isDataRow(row);

    if (!dataStarted && isData) {
      dataStarted = true;
    }

    if (!dataStarted) {
      headerRows.push(row);
      continue;
    }

    const columns = buildColumns(headerRows);
    const rowDims = extractRowDimensions(row);

    let valueIndex = 0;

    for (let j = 1; j < row.length; j++) {

      let value = row[j];

      if (!value || value === "" || isNaN(value)) continue;

      value = value.toString().trim();

      let columna = columns[valueIndex] || "";

      output.push({
        panel: panelName + " | " + reportSuite + " | " + date,
        filtro: currentSegments.join(" | "), // 👈 renombrado como querías
        tabla: currentTable,
        fila: rowDims.join(" | "),
        columna: columna,
        valor: value
      });

      valueIndex++;
    }
  }

  return output;
}


/**
 * ==================================
 * MÓDULO: TABLE NAME PROPAGATION (JOIN STYLE)
 * ==================================
 *
 * No divide bloques.
 * Mantiene memoria del último nombre de tabla (# Nombre)
 * y lo asigna a cada fila de data.
 */
function splitTables(lines) {

  let tables = [];
  let currentName = "";

  for (let i = 0; i < lines.length; i++) {

    let rawLine = lines[i];

    let line = rawLine
      .replace(/"/g, "")
      .replace("\r", "")
      .trim();

    // ✅ Detectar nombre de tabla (# Nombre o ## Nombre)
    if (line.startsWith("#")) {

      let clean = line.replace(/^##?/, "").trim();

      // Evitar agarrar separadores tipo ##### o ==== 
      if (clean !== "" && !clean.match(/^[-=]+$/) && !clean.match(/^#+$/)) {
        currentName = clean;
      }

      continue;
    }

    // ✅ Ignorar vacío
    if (line === "") continue;

    // ✅ Data real
    tables.push({
      name: currentName,
      row: line.split(",")
    });
  }

  return tables;
}

/**
 * ===============================
 * ROW DIMENSIONS
 * ===============================
 */
function extractRowDimensions(row) {

  let dims = [];

  for (let i = 0; i < row.length; i++) {

    let cell = row[i];

    if (!cell) continue;

    let clean = cell.toString().trim();

    if (clean !== "" && !isNaN(clean)) break;

    if (clean !== "") dims.push(clean);
  }

  return dims;
}


/**
 * ===============================
 * DATA ROW DETECTOR
 * ===============================
 */
function isDataRow(row) {

  if (!row || row.length === 0) return false;

  let first = row[0] ? row[0].trim() : "";

  // ✅ debe tener dimensión en primera columna
  if (first === "") return false;

  // ✅ debe tener número en el resto
  for (let i = 1; i < row.length; i++) {
    let cell = row[i];
    if (cell && !isNaN(cell)) {
      return true;
    }
  }

  return false;
}


/**
 * ===============================
 * HEADER BUILDER + CLEAN
 * ===============================
 */
function buildColumns(headerRows) {

  let maxCols = Math.max(...headerRows.map(r => r.length));
  let columns = [];

  for (let col = 0; col < maxCols; col++) {

    let levels = [];

    for (let r = 0; r < headerRows.length; r++) {
      let val = headerRows[r][col];

      if (val && val.trim() !== "") {
        levels.push(val.trim());
      }
    }

    // ✅ SOLO guardar columnas que realmente tienen contenido
    if (levels.length > 0) {
      columns.push(levels.join(" | "));
    }
  }

  return columns;
}

/**
 * ===============================
 * OUTPUT
 * ===============================
 */

function convertToRows(data) {

  let rows = [["panel", "filtro", "tabla", "fila", "columna", "valor"]];

  for (let i = 0; i < data.length; i++) {

    rows.push([
      data[i].panel || "",
      data[i].filtro || "",
      data[i].tabla || "",
      data[i].fila || "",
      data[i].columna || "",
      Number(data[i].valor)
    ]);
  }

  return rows;
}
