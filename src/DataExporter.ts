/// <reference path="ConfigTypes.ts" />

interface ExportResult {
  fileId: string;
  fileName: string;
}

function exportAllData(): ExportResult {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportData: Record<string, unknown[][]> = {};

  const sheetNames = [
    SHEET_NAMES.SPIELER,
    SHEET_NAMES.ABWESENHEITEN,
    SHEET_NAMES.SAISON,
  ];

  for (const name of sheetNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 0) {
        const range = sheet.getRange(1, 1, lastRow, lastCol);
        exportData[name] = range.getValues();
      }
    }
  }

  const timestamp = Utilities.formatDate(new Date(), 'Europe/Berlin', 'yyyy-MM-dd_HH-mm-ss');
  const fileName = `einsatzplaner_export_${timestamp}.json`;
  const content = JSON.stringify(exportData, null, 2);

  const folder = getOrCreateExportFolder();
  const existing = folder.getFilesByName(fileName);

  let file: GoogleAppsScript.Drive.File;
  if (existing.hasNext()) {
    file = existing.next();
    file.setContent(content);
  } else {
    file = folder.createFile(fileName, content, 'application/json');
  }

  return {
    fileId: file.getId(),
    fileName: fileName,
  };
}

function getOrCreateExportFolder(): GoogleAppsScript.Drive.Folder {
  const folderName = 'Einsatzplaner Exports';
  const folders = DriveApp.getFoldersByName(folderName);

  if (folders.hasNext()) {
    return folders.next();
  }

  return DriveApp.createFolder(folderName);
}
