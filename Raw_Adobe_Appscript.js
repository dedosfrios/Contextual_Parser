function processAdobeCSV() {
  const inputFolderName = "RAW_ADOBE";
  const outputFolderName = "PROCESSED_ADOBE";

  const inputFolder = DriveApp.getFoldersByName(inputFolderName).next();
  const outputFolder = DriveApp.getFoldersByName(outputFolderName).next();

  const files = inputFolder.getFilesByType(MimeType.CSV);

  while (files.hasNext()) {
    const file = files.next();
    const content = file.getBlob().getDataAsString();
    
    const parsed = parseAdobe(content);

    const outputCsv = convertToCSV(parsed);

    outputFolder.createFile(
      file.getName().replace(".csv", "_processed.csv"),
      outputCsv,
      MimeType.CSV
    );
  }
}