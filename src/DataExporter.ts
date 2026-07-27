/// <reference path="ConfigTypes.ts" />

interface ExportResult {
  fileId: string;
  fileName: string;
}

function exportAllData(): ExportResult {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const exportData: Record<string, unknown[][]> = {};

  const sheetNames = [SHEET_NAMES.SPIELER, SHEET_NAMES.ABWESENHEITEN, SHEET_NAMES.SAISON];

  for (const name of sheetNames) {
    const sheet = ss.getSheetByName(name);
    if (sheet) {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow > 0) {
        exportData[name] = sheet.getRange(1, 1, lastRow, lastCol).getValues();
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

  return { fileId: file.getId(), fileName };
}

function getOrCreateExportFolder(): GoogleAppsScript.Drive.Folder {
  const props = PropertiesService.getScriptProperties();
  const folderId = props.getProperty('EXPORT_FOLDER_ID');

  if (folderId) {
    try {
      return DriveApp.getFolderById(folderId);
    } catch (_e) {
      props.deleteProperty('EXPORT_FOLDER_ID');
    }
  }

  const folder = DriveApp.createFolder('Einsatzplaner Exports');
  props.setProperty('EXPORT_FOLDER_ID', folder.getId());
  return folder;
}
