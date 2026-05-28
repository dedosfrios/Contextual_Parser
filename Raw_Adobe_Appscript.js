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
  const rawRows = splitTables(lines);

  let output = [];

  let block = [];

  // reconstruir bloques simples (solo para headers)
  for (let i = 0; i < rawRows.length; i++) {

    const entry = rawRows[i];

    block.push(entry);
  }

  // 👉 todo el archivo es un solo flujo
  // detectamos headers dinámicamente

  let headerRows = [];
  let dataStarted = false;

  for (let i = 0; i < block.length; i++) {

    let row = block[i].row;

    if (!dataStarted && isDataRow(row)) {
      dataStarted = true;
    }

    if (!dataStarted) {
      headerRows.push(row);
      continue;
    }

    const columns = buildColumns(headerRows);
    const rowDims = extractRowDimensions(row);

    for (let j = 1; j < row.length; j++) {

      let value = row[j];

      if (!value || value === "" || isNaN(value)) continue;

      value = value.toString().trim();

      output.push({
        tabla: block[i].name,
        fila: rowDims.join(" | "),
        columna: columns[j - 1] || "",
        valor: value
      });
    }
  }

  return output;
}

