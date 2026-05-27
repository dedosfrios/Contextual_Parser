/**
 * ==========================================
 * MÓDULO: MAIN PROCESSOR
 * ==========================================
 */
function processAdobeCSV() {

  const inputFolder = DriveApp.getFoldersByName("RAW_ADOBE").next();
  const outputFolder = DriveApp.getFoldersByName("PROCESSED_ADOBE").next();

  const files = inputFolder.getFiles();

  if (!files.hasNext()) {
    Logger.log("No hay archivos en RAW_ADOBE");
    return;
  }

  const file = files.next();
  const content = file.getBlob().getDataAsString();

  const parsed = parseAdobe(content);

  if (!parsed || parsed.length === 0) {
    Logger.log("No hay datos parseados");
    return;
  }

  const rows = convertToRows(parsed);

  const sheet = SpreadsheetApp.create(file.getName() + "_processed");
  const ws = sheet.getActiveSheet();

  ws.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

  const fileDrive = DriveApp.getFileById(sheet.getId());
  outputFolder.addFile(fileDrive);
  DriveApp.getRootFolder().removeFile(fileDrive);
}


/**
 * ==========================================
 * MÓDULO: PARSER
 * ==========================================
 */
function parseAdobe(content) {

  const lines = content.split("\n");
  const tables = splitTables(lines);

  let output = [];

  for (let t = 0; t < tables.length; t++) {

    const table = tables[t];
    const rows = table.rows;

    if (!rows || rows.length === 0) continue;

    let headerRows = [];
    let dataStart = 0;

    for (let i = 0; i < rows.length; i++) {
      if (isDataRow(rows[i])) {
        dataStart = i;
        break;
      }
      headerRows.push(rows[i]);
    }

    const columns = buildColumns(headerRows);

    for (let i = dataStart; i < rows.length; i++) {

      const row = rows[i];
      const rowDims = extractRowDimensions(row);

      for (let j = 1; j < row.length; j++) {

        let value = row[j];

        if (!value || value === "" || isNaN(value)) continue;

        value = value.toString().replace("\r", "").trim();

        output.push({
          tabla: table.name,
          fila: rowDims.join(" | "),
          columna: columns[j - 1] || "",
          valor: value
        });
      }
    }
  }

  return output;
}


/**
 * ==========================================
 * MÓDULO: SPLIT TABLES
 * ==========================================
 */
function splitTables(lines) {

  let tables = [];
  let currentRows = [];
  let currentName = "";

  for (let i = 0; i < lines.length; i++) {

    let line = lines[i].replace(/"/g, "").replace("\r", "").trim();

    if (line.includes("################################")) {

      if (currentRows.length > 0) {
        tables.push({
          name: currentName,
          rows: currentRows
        });
        currentRows = [];
      }

    } else if (line.startsWith("##")) {

      currentName = line.replace("##", "").trim();

    } else {

      if (line !== "") {
        currentRows.push(line.split(","));
      }
    }
  }

  if (currentRows.length > 0) {
    tables.push({
      name: currentName,
      rows: currentRows
    });
  }

  return tables;
}


/**
 * ==========================================
 * MÓDULO: ROW DIMENSION DETECTOR
 * ==========================================
 */
function extractRowDimensions(row) {

  let dims = [];

  if (!row || row.length === 0) return dims;

  for (let i = 0; i < row.length; i++) {

    let cell = row[i];

    if (!cell) continue;

    let clean = cell.toString().replace("\r", "").trim();

    if (clean !== "" && !isNaN(clean)) break;

    if (clean !== "") dims.push(clean);
  }

  return dims;
}


/**
 * ==========================================
 * MÓDULO: DATA ROW DETECTOR
 * ==========================================
 */
function isDataRow(row) {

  if (!row || row.length === 0) return false;
  if (!row[0] || row[0].trim() === "") return false;

  return row.slice(1).some(cell => {
    return cell !== "" && !isNaN(cell);
  });
}


/**
 * ==========================================
 * MÓDULO: HEADER BUILDER
 * ==========================================
 */
function buildColumns(headerRows) {

  let columns = [];
  let maxCols = Math.max(...headerRows.map(r => r.length));

  for (let col = 0; col < maxCols; col++) {

    let levels = [];

    for (let r = 0; r < headerRows.length; r++) {

      let val = headerRows[r][col];

      if (val && val.trim() !== "") {
        levels.push(val.trim());
      }
    }

    columns.push(levels.join(" | "));
  }

  return columns;
}


/**
 * ==========================================
 * MÓDULO: OUTPUT FORMAT
 * ==========================================
 */
function convertToRows(data) {

  let rows = [["tabla", "fila", "columna", "valor"]];

  for (let i = 0; i < data.length; i++) {

    rows.push([
      data[i].tabla,
      data[i].fila,
      data[i].columna,
      Number(data[i].valor)
    ]);
  }

  return rows;
}