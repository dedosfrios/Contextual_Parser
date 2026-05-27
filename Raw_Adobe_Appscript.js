/**
 * MÓDULO: MAIN PROCESSOR (FIXED OUTPUT FOLDER)
 * DESCRIPCIÓN:
 * Ahora el resultado se guarda dentro de PROCESSED_ADOBE
 */
function processAdobeCSV() {

  const inputFolderName = "RAW_ADOBE";
  const outputFolderName = "PROCESSED_ADOBE";

  const inputFolder = DriveApp.getFoldersByName(inputFolderName).next();
  const outputFolder = DriveApp.getFoldersByName(outputFolderName).next();

  const files = inputFolder.getFiles();

  while (files.hasNext()) {

    const file = files.next();
    const content = file.getBlob().getDataAsString();

    const parsed = parseAdobe(content);

    if (!parsed || parsed.length === 0) {
      Logger.log("Sin datos parseados en: " + file.getName());
      continue;
    }

    const rows = convertToRows(parsed);

    // --- Crear sheet ---
    const sheet = SpreadsheetApp.create(file.getName() + "_processed");
    const ws = sheet.getActiveSheet();

    ws.getRange(1, 1, rows.length, rows[0].length).setValues(rows);

    // --- MOVER A CARPETA ---
    const fileId = sheet.getId();
    const fileDrive = DriveApp.getFileById(fileId);

    outputFolder.addFile(fileDrive);
    DriveApp.getRootFolder().removeFile(fileDrive);

    Logger.log("✅ Sheet movido a carpeta: " + outputFolderName);
  }
}


/**
 * MÓDULO: UNIVERSAL ADOBE PARSER
 * DESCRIPCIÓN:
 * - Divide el CSV en bloques (tablas)
 * - Detecta headers multinivel
 * - Detecta breakdown en filas
 * - Reconstruye el contexto de cada valor
 * - Devuelve datos en formato LONG (tabla, fila, columna, valor)
 */
function parseAdobe(content) {

  const lines = content.split("\n");

  let tables = [];
  let currentTable = [];
  let currentName = "";

  // --- Paso 1: separar tablas ---
  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];


line = line.replace(/"/g, "").trim();

/**
 * FIX: TABLE NAME DETECTOR
 * DESCRIPCIÓN:
 * Ignora separadores (########)
 * Solo toma líneas con texto real como nombre de tabla
 */
line = line.replace(/"/g, "").trim();

if (line.startsWith("##")) {

  let name = line.replace("##", "").trim();

  // ignorar basura tipo ########
  if (name.replace(/#/g, "").trim().length > 0) {
    currentName = name;
  }
}


    if (line.includes("################################")) {
      if (currentTable.length > 0) {
        tables.push({
          name: currentName,
          rows: currentTable
        });
        currentTable = [];
      }
      continue;
    }

    if (line.trim() !== "") {
      currentTable.push(line.split(","));
    }
  }

  let output = [];

  // --- Paso 2: procesar cada tabla ---
  tables.forEach(table => {

    const rows = table.rows;
    if (!rows || rows.length === 0) return;

    let headerRows = [];
    let dataStart = 0;

    // detectar dónde empiezan los datos
    for (let i = 0; i < rows.length; i++) {
      if (isDataRow(rows[i])) {
        dataStart = i;
        break;
      }
      headerRows.push(rows[i]);
    }

    const columns = buildColumns(headerRows);

    // --- Paso 3: procesar filas de datos ---
    for (let i = dataStart; i < rows.length; i++) {

      let row = rows[i];

      
/**
 * FIX: ROW DIMENSION DETECTOR
 * DESCRIPCIÓN:
 * Detecta correctamente breakdown en filas.
 * Toma todas las columnas iniciales hasta que aparecen valores numéricos.
 */
function extractRowDimensions(row) {

  let dims = [];

  for (let i = 0; i < row.length; i++) {

    let cell = row[i];

    // cuando encontramos números → ya entramos a métricas
    if (cell !== "" && !isNaN(cell)) {
      break;
    }

    if (cell && cell.trim() !== "") {
      dims.push(cell.trim());
    }
  }

  return dims;
}


let rowDims = extractRowDimensions(row);
      // recorrer valores
      for (let j = 1; j < row.length; j++) {

        let value = row[j];

        if (value === "" || value === null || value === undefined) continue;
        if (value === "Infinity") continue;
        if (isNaN(value)) continue;

        output.push({
          tabla: table.name,
          fila: rowDims.join(" | "),
          columna: columns[j - 1] || "",
          valor: value
        });
      }
    }
  });

  return output;
}

/**
 * MÓDULO: DATA ROW DETECTOR
 * DESCRIPCIÓN:
 * Determina si una fila corresponde a datos reales
 * (no headers). Se basa en:
 * - Primera celda no vacía
 * - Al menos un valor numérico en la fila
 */
function isDataRow(row) {

  if (!row || row.length === 0) return false;
  if (!row[0] || row[0].trim() === "") return false;

  return row.slice(1).some(cell => {
    return cell !== "" && !isNaN(cell);
  });
}

/**
 * MÓDULO: MULTILEVEL HEADER BUILDER
 * DESCRIPCIÓN:
 * Reconstruye headers multinivel combinando todas las filas
 * de encabezado en una sola columna tipo:
 * "Segmento | Periodo | Métrica"
 */
function buildColumns(headerRows) {

  if (!headerRows || headerRows.length === 0) return [];

  let maxCols = Math.max(...headerRows.map(r => r.length));
  let columns = [];

  for (let col = 0; col < maxCols; col++) {

    let levels = [];

    headerRows.forEach(row => {
      let val = row[col];

      if (val && val.trim() !== "") {
        levels.push(val.trim());
      }
    });

    if (levels.length > 0) {
      columns.push(levels.join(" | "));
    }
  }

  return columns;
}

/**
 * MÓDULO: OUTPUT FORMATTER
 * DESCRIPCIÓN:
 * Convierte el resultado del parser en una matriz 2D
 * lista para ser escrita en Google Sheets
 */
function convertToRows(data) {

  const headers = ["tabla", "fila", "columna", "valor"];

  let rows = [headers];

  data.forEach(row => {
    rows.push([
      row.tabla,
      row.fila,
      row.columna,
      Number(row.valor)
    ]);
  });

  return rows;
}