/**
 * ==========================================
 * MÓDULO: PANEL METADATA EXTRACTOR (DEBUG SAFE)
 * ==========================================
 *
 * Captura:
 * - panel (estructura general)
 * - report suite
 * - date
 * - segments (raw, sin limpiar aún)
 *
 * NO interactúa con el parser
 */
function extractPanelMetadata(content) {

  const lines = content.split("\n");

  let panels = [];

  let currentPanel = {
    panel: "",
    reportSuite: "",
    date: "",
    segments: ""
  };

  for (let i = 0; i < lines.length; i++) {

    let line = lines[i]
      .replace(/"/g, "")
      .replace("\r", "")
      .trim();

    if (!line.startsWith("#")) continue;

    let clean = line.replace(/^#+/, "").trim();

    if (clean === "" || clean.match(/^[-=]+$/)) continue;

    // ===============================
    // PANEL NAME
    // ===============================
    if (clean === "Panel" || clean === "Freeform") {

      // guardar anterior si existe
      if (currentPanel.panel !== "") {
        panels.push(currentPanel);
      }

      // iniciar nuevo panel
      currentPanel = {
        panel: clean,
        reportSuite: "",
        date: "",
        segments: ""
      };
    }

    // ===============================
    // REPORT SUITE
    // ===============================
    if (clean.startsWith("Report suite:")) {
      currentPanel.reportSuite =
        clean.replace("Report suite:", "").trim();
    }

    // ===============================
    // DATE
    // ===============================
    if (clean.startsWith("Date:")) {
      currentPanel.date =
        clean.replace("Date:", "").trim();
    }

    // ===============================
    // SEGMENTS (RAW)
    // ===============================
    if (clean.startsWith("Segments:")) {
      currentPanel.segments =
        clean.replace("Segments:", "").trim();
    }
  }

  // push último
  if (currentPanel.panel !== "") {
    panels.push(currentPanel);
  }

  Logger.log("===== PANEL METADATA =====");
  panels.forEach(p => Logger.log(JSON.stringify(p)));

  return panels;
}


function runMetadataDebug() {

  const inputFolder = DriveApp.getFoldersByName("RAW_ADOBE").next();
  const file = inputFolder.getFiles().next();

  const content = file.getBlob().getDataAsString();

  extractPanelMetadata(content);
}
