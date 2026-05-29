/**
 * ==========================================
 * MÓDULO: UNIFIED PIPELINE
 * ==========================================
 *
 * Conecta:
 * - metadata (panel)
 * - parser (data)
 */
function processWithMetadata() {

  const inputFolder = DriveApp.getFoldersByName("RAW_ADOBE").next();
  const file = inputFolder.getFiles().next();

  const content = file.getBlob().getDataAsString();

  // ✅ 1. obtener metadata
  const panels = extractPanelMetadata(content);

  // ✅ 2. obtener data parseada
  const data = parseAdobe(content);

  // ✅ 3. DEBUG simple
  Logger.log("TOTAL PANELS: " + panels.length);
  Logger.log("TOTAL ROWS: " + data.length);

  // 🔥 ahora solo devolvemos ambos (aún no los unimos)
  return {
    panels: panels,
    data: data
  };
}