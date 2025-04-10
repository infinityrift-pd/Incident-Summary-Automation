/**
 * Automatically copies the template spreadsheet and renames it with the current month and year.
 * The new spreadsheet is placed in a folder structure: _YYYY/_YYYY-MM.
 * If the folder structure doesn't exist, it is created.
 * If a file with the same name already exists in the target folder, the function exits without creating a duplicate.
 * This function is designed to run only on the template spreadsheet.
 */
function autoCopyAndRenameSpreadsheet() {
  // Check if this is the template spreadsheet
  if (!isTemplateSpreadsheet()) {
    Logger.log("This function can only run on the template spreadsheet. Exiting.");
    return;
  }

  // Get the active spreadsheet and its parent folder
  var originalSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var parentFolder = DriveApp.getFileById(originalSpreadsheet.getId()).getParents().next();

  // Get next month's date information (for creating the new spreadsheet)
  var now = new Date();
  var monthValue = new Date(now.getFullYear(), now.getMonth() + 1, 1); // +1 to get Next Month value
  var month = String(monthValue.getMonth() + 1).padStart(2, '0'); // Months are zero-based indexed (pos[Jan] = 0). +1 to get next month's date value
  var monthName = monthValue.toLocaleString('default', { month: 'long' });
  var year = monthValue.getFullYear();

  // Create the new file name
  var newFileName = `Incident Summary - ${monthName} ${year}`;

  // Create or get the year folder
  var yearFolderName = `_${year}`;
  var yearFolder = getOrCreateFolder(parentFolder, yearFolderName);

  // Create or get the month folder
  var monthFolderName = `_${year}-${month}`;
  var monthFolder = getOrCreateFolder(yearFolder, monthFolderName);

  // Check if the file already exists in the month folder
  if (monthFolder.getFilesByName(newFileName).hasNext()) {
    Logger.log(`File already exists: ${newFileName} in folder ${yearFolderName}/${monthFolderName}`);
    return; // Exit the function if the file exists
  }

  // Create a copy of the spreadsheet
  var newSpreadsheet = originalSpreadsheet.copy(newFileName);
  var newFile = DriveApp.getFileById(newSpreadsheet.getId());

  // Move the new file to the month folder
  newFile.moveTo(monthFolder);

  // Log the result
  var fileUrl = newFile.getUrl();
  Logger.log(`New spreadsheet created: ${newFileName} in folder ${yearFolderName}/${monthFolderName} at ${fileUrl}`);
}

/**
 * Gets an existing folder by name or creates a new one if it doesn't exist.
 * @param {Folder} parentFolder - The parent folder to search in or create the new folder.
 * @param {string} folderName - The name of the folder to find or create.
 * @return {Folder} The existing or newly created folder.
 */
function getOrCreateFolder(parentFolder, folderName) {
  var folder = parentFolder.getFoldersByName(folderName);
  if (folder.hasNext()) {
    return folder.next();
  } else {
    var newFolder = parentFolder.createFolder(folderName);
    Logger.log(`Created folder: ${folderName}`);
    return newFolder;
  }
}

/**
 * Checks if the current spreadsheet is the template spreadsheet.
 * @return {boolean} True if this is the template spreadsheet, false otherwise.
 */
function isTemplateSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  return spreadsheet.getName() === "TEMPLATE_Incident Summary";
}
