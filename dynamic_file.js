function onOpen() {
  SpreadsheetApp.getUi().createMenu('Custom Report Links')
      .addItem('Authenticate Perms & Triggers', 'createTriggers')
      .addItem('Refresh Folder Links', 'refreshFolderLinksAndAssignees')
      .addItem('Populate Mitigation Matrix', 'readDocToMitigationMatrix')
      .addToUi();
}

function refreshFolderLinksAndAssignees() {
  getFolderLinksOfCurrentSheet();
  populateAssignees();
}


/**
 * This function retrieves the links of subfolders within the Incident parent folder of the current spreadsheet
 * and populates them into the "True Positives" sheet as hyperlinks, ensuring no duplicates are added.
 * @function getFolderLinksOfCurrentSheet
 */
function getFolderLinksOfCurrentSheet() {
  // Get references to the active spreadsheet and the "True Positives" sheet
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName("True Positives");

  // Check if the "True Positives" sheet exists, alert and exit if not found
  if (!sheet) {
    SpreadsheetApp.getUi().alert('No sheet named "True Positives" found.');
    return;
  }

  // Get the parent folder of the current spreadsheet
  var folder = DriveApp.getFileById(spreadsheet.getId()).getParents().next();
  var row = 2; // Start populating from the second row

  // If the spreadsheet is within a folder
  if (folder) {
    // Get all subfolders within the parent folder
    var subfolders = folder.getFolders();

    // Iterate through each subfolder
    while (subfolders.hasNext()) {
      var subfolder = subfolders.next();
      var currentCell = sheet.getRange("A" + row);

      // Check if the current cell is empty or doesn't already contain a hyperlink to the subfolder
      if (currentCell.getValue() === "" || !currentCell.getFormula().includes(subfolder.getUrl())) {
        // Set the cell value to a hyperlink formula linking to the subfolder
        currentCell.setValue('=HYPERLINK("' + subfolder.getUrl() + '","' + subfolder.getName() + '")');
      }
      row++; // Move to the next row for the next subfolder link
    }
  } else {
    // If the spreadsheet is not in a folder and the sheet is empty, add a message
    if (sheet.getLastRow() === 0) {
      sheet.getRange("A2").setValue("This Sheet is not in a folder.");
    }
  }
}

/**
 * Populates the 'True Positives' sheet with assignee names, project statuses, and incident summaries.
 * 
 * This function iterates through Google Docs in subfolders of the parent folder containing the active spreadsheet.
 * For each subfolder, it locates the Google Doc with the same name as the subfolder and extracts:
 * 1. Assignee names from dropdown fields labeled as "ANALYST"
 * 2. Project statuses from dropdown fields labeled as "Project status"
 * 3. Incident summaries from the document content
 * 
 * The extracted data is then written to the 'True Positives' sheet, starting from row 2, columns B, C, and D.
 * @function populateAssignees
 */
/**
 * Populates the 'True Positives' sheet with assignee names, project statuses, and incident summaries.
 * Includes document change detection and intelligent caching.
 */
function populateAssignees() {
  Logger.log("Entered populateAssignees");
  const sheetName = 'True Positives';
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Cannot find sheet named "${sheetName}"`);
    return;
  }

  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  const subfolders = parentFolder.getFolders();
  
  // Create batch data array for writing
  const batchData = [];
  
  // Get cache service
  const cache = CacheService.getScriptCache();
  
  while (subfolders.hasNext()) {
    const subfolder = subfolders.next();
    const subfolderName = subfolder.getName();
    const files = subfolder.getFilesByType(MimeType.GOOGLE_DOCS);

    while (files.hasNext()) {
      const file = files.next();
      const docId = file.getId();
      const docName = file.getName();
      const lastModified = file.getLastUpdated().getTime();

      if (docName === subfolderName) {
        // Get cached data and metadata
        const cachedMetadata = cache.get(`metadata_${docId}`);
        const cachedData = cache.get(`data_${docId}`);
        
        let needsProcessing = true;
        
        if (cachedMetadata && cachedData) {
          const metadata = JSON.parse(cachedMetadata);
          // Check if the document has been modified since last processing
          if (metadata.lastModified === lastModified) {
            needsProcessing = false;
            const parsedData = JSON.parse(cachedData);
            batchData.push([
              parsedData.assignees,
              parsedData.reviewers,
              parsedData.status,
              parsedData.summary,
              parsedData.status === 'TP:EVIL' ? parsedData.ttd : '']);
            Logger.log(`Using cached data for document ${docName}`);
          } else {
            Logger.log(`Document ${docName} has been modified, reprocessing`);
          }
        }

        if (needsProcessing) {
          try {
            const { dropDownValues, summaryText, ttdDecimal } = processDocument(docId);
            
            // Extract values from dropdowns
            const assigneeNames = [];
            const reviewerNames = [];
            let projectStatus = '';
            
            for (const dropdown of dropDownValues) {
              if (dropdown.type === "ANALYST" && dropdown.currentValue !== "SELECT Analyst") {
                assigneeNames.push(dropdown.currentValue);
              } else if (dropdown.type === "REVIEWER" && dropdown.currentValue !== "SELECT Analyst") {
                reviewerNames.push(dropdown.currentValue);
              } else if (dropdown.type === "Project status") {
                projectStatus = dropdown.currentValue;
              }
            }

            // Prepare data for caching
            const processedData = {
              assignees: assigneeNames.join(", "),
              reviewers: reviewerNames.join(", "),
              status: projectStatus,
              summary: summaryText,
              ttdDecimal: ttdDecimal
            };

            // Cache both the data and metadata
            const metadata = {
              lastModified: lastModified,
              processedAt: new Date().getTime()
            };

            // Use batch operation for caching
            cache.putAll({
              [`data_${docId}`]: JSON.stringify(processedData),
              [`metadata_${docId}`]: JSON.stringify(metadata)
            }, 21600); // Cache for 6 hours
            
            batchData.push([
              processedData.assignees,
              processedData.reviewers,
              processedData.status,
              processedData.summary,
              processedData.status === 'TP:EVIL' ? processedData.ttdDecimal : ''
              ]);
            Logger.log(`Processed and cached document ${docName}`);
          } catch (error) {
            Logger.log(`Error processing document ${docId}: ${error}`);
            continue;
          }
        }
      }
    }
  }
  
  // Write all data at once if there's data to write
  if (batchData.length > 0) {
    sheet.getRange(2, 2, batchData.length, 5).setValues(batchData);
  }
}

/**
 * Process a single document to extract dropdown values and summary
 */
function processDocument(docId) {
  const docFile = DocumentApp.openById(docId);
  const docBody = docFile.getBody();
  
  // Get XML content using more efficient approach
  const url = `https://docs.google.com/feeds/download/documents/export/Export?exportFormat=docx&id=${docId}`;
  const blob = UrlFetchApp.fetch(url, { 
    headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  }).getBlob();
  
  blob.setContentType("application/zip");
  const xmlContent = Utilities.unzip(blob)
    .find(file => file.getName() === "word/document.xml")
    ?.getDataAsString() || "";
  
  // Use regex for faster parsing of dropdown values
  const dropDownRegex = /<w:alias w:val="([^"]+)"[^>]*>(?:[\s\S]*?)<w:dropDownList w:lastValue="([^"]+)"/g;
  const dropDownValues = [];
  let match;
  
  while ((match = dropDownRegex.exec(xmlContent)) !== null) {
      dropDownValues.push({
        type: match[1], // The alias value in the xml
        currentValue: match[2] // THe lastValue value
      });
  }

  // Extract TTD
  const ttdRegex = /&lt;(\d+:\d{2})&gt;/;
  const ttdMatch = xmlContent.match(ttdRegex);
  const ttdString = ttdMatch ? ttdMatch[1] : '';
  const ttdDecimal = convertTimeToDecimal(ttdString);

  
  // Get summary more efficiently
  let summaryText = "";
  let foundSummaryHeading = false;
  const paragraphs = docBody.getParagraphs();
  
  for (let i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].getText().trim();
    
    if (text === "A summary of the incident") {
      foundSummaryHeading = true;
      continue;
    }
    
    if (text === "Mitigation Matrix:") {
      break;
    }
    
    if (foundSummaryHeading && text) {
      summaryText += text + "\n";
    }
  }

  return {
    dropDownValues,
    summaryText: summaryText.trim(),
    ttd: ttdString,
    ttdDecimal: ttdDecimal
  };
}

/**
 * Categorizes an incident based on keywords in its name.
 * 
 * This function takes an incident name as input and compares it against
 * predefined categories and their associated keywords. It returns the
 * first matching category or 'UNCATEGORIZED' if no match is found.
 * 
 * The categorization is case-insensitive.
 * 
 * @param {string} incidentName - The name of the incident to categorize.
 * @returns {string} The matching category name or 'UNCATEGORIZED'.
 */
function categoriseIncident(incidentName) {
  // Define categories and their associated keywords
  // Each category is a key in the CATEGORIES object
  // The value for each key is an array of keywords associated with that category
  const CATEGORIES = {
    'Stores': ['STORES', 'RETAIL'],
    'Orca': ['ORCA'],
    'Network': ['EHOP', 'DISTB'],
    'Cloud': ['AWS', 'AZURE', 'GCP', 'GWS'],
    'Rewards': ['RWDS'],
    'MyDeal': ['MDEAL'],
    'BigW': ['BIGW'],
    'WowCorp': ['CORP', 'WOW', 'WOWGA'],
    'GFS': ['GFS'],
    'Hoax Mailbox': ['HOAX'],
  };

  // ADD MORE OR MODIFY AS NEEDED
  // To add a new category, add a new key-value pair to the CATEGORIES object
  // To modify keywords for a category, edit the array of keywords for that category

  // Iterate through each category and its keywords
  for (let [category, keywords] of Object.entries(CATEGORIES)) {
    // Check if any keyword for this category is included in the incident name
    // The .toUpperCase() method is used to make the comparison case-insensitive
    if (keywords.some(keyword => incidentName.toUpperCase().includes(keyword))) {
      return category; // Return the first matching category
    }
  }

  // If no category matches, return 'UNCATEGORIZED'
  return 'UNCATEGORIZED';
}

/**
 * This function reads tables from Google Docs within a specified folder structure and appends their data to a designated sheet in the active spreadsheet.
 *
 * @function readDocToMitigationMatrix
 */
function readDocToMitigationMatrix() {
  // Specify sheet name for data to be pasted and the required table headers in the GDocs the function will search for.
  var sheetName = 'rawMitigationScoreMatrix';
  var tableHeaders = ["inbound", "endpoint", "outbound", "other"]; 

  // Get dynamic references to the active spreadsheet and target sheet.
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);

  // Check if the target sheet exists; if not, raise alert and exit.
  if (!sheet) {
    SpreadsheetApp.getUi().alert(`Cannot find sheet named "${sheetName}"`);
    return; 
  }

  // Clear the entire sheet before adding the data
  sheet.clear()
  Logger.log("Sheet Cleared")

  // Get the parent folder of the spreadsheet and its subfolders in Shared SOC GDrive.
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();
  var subfolders = parentFolder.getFolders(); 
  var allData = []; 

  // Iterate through each subfolder within the parent folder.
  while (subfolders.hasNext()) {
    var subfolder = subfolders.next();
    var files = subfolder.getFilesByType(MimeType.GOOGLE_DOCS); 

    // Iterate through each Google Doc file in the current subfolder.
    //ASSUMPTION: Only 1 GDoc exists within each subfolder
    while (files.hasNext()) { 
      // Open the document, assign the document name, get its tables and paragraphs.
      var docFile = DocumentApp.openById(files.next().getId());
      var docName = docFile.getName();
      var docBody = docFile.getBody();
      var docTables = docBody.getTables();
      var docElements = docBody.getParagraphs();
      
      // Retrieve Incident Summary
      // 1. Find the headings
      let summaryText = "";
      let foundSummaryHeading = false;

      for (var element of docElements) {
        var elementText = element.getText();
        var headingType = element.getHeading();

        if (elementText.trim() === "A summary of the incident") {
          foundSummaryHeading = true;
        } else if (elementText.trim() === "Mitigation Matrix:") {
          summaryText = summaryText.trim() // Remove trailing enters included in the concatenation
          break; // Stop when we reach the matrix heading
        } else if (foundSummaryHeading) {
          summaryText += elementText + "\n"; //Append new line in between found paragraph elements
        }
      }

      let mitMatrixTable = null;

      // Iterate through the document's tables to find the one matching the headers.
      for (var table of docTables) {
        var headerRow = table.getRow(0);
        var docHeaderValues = [];

        // Extract header values from the table.
        for (let j = 0; j < headerRow.getNumCells(); j++) {
          docHeaderValues.push(headerRow.getCell(j).getText().toLowerCase());
        }

        // Check if the table headers match the required headers.
        if (tableHeaders.every(header => docHeaderValues.includes(header.toLowerCase()))) {
          mitMatrixTable = table; 
          break; 
        }
      }
      
      // Skip the document if no Mitigation Matrix table is found.
      if (!mitMatrixTable) continue; 

      var numRows = mitMatrixTable.getNumRows();
      const numCols = mitMatrixTable.getRow(0).getNumCells();

      // Iterate through the rows of the matched mitigation matrix table (skipping the first two header rows).
      for (let i = 2; i < numRows; i++) { 
        var row = mitMatrixTable.getRow(i);
        var category = categoriseIncident(docName);
        var rowData = [docName, summaryText, category]; //Set the docName and summaryText as the first values in the rowData Array
        // Extract data from each cell in the row.
        for (let j = 0; j < numCols; j++) {
          var cell = row.getChild(j);
          if (cell.getType() == DocumentApp.ElementType.TABLE_CELL) {
            rowData.push(cell.getChild(0).asText().getText());
          }
        }

        // Add the extracted row data to the allData array - entire row at a time.
        allData.push(rowData); 
      }
    }
  }

  if (allData.length > 0) {
    const targetRange = sheet.getRange(2, 1, allData.length, allData[0].length);
    targetRange.setValues(allData);
  } else {
    SpreadsheetApp.getUi().alert(`No tables matching the specified headers were found in any documents.`);
  }

  // // Commented-out code for auto-resizing columns and rows (is buggy, doesn't work well: https://stackoverflow.com/questions/54516371/google-apps-script-autoresizecolumn-works-incorrectly-not-works-as-expected).
  // SpreadsheetApp.flush(); 
  // sheet.autoResizeColumns(1, allData[0].length);
  // SpreadsheetApp.flush(); 
  // sheet.autoResizeRows(1, allData.length + lastRow);
}

/**
 * Converts time in "HH:MM" format to decimal hours
 * @param {string} timeString - Time in format "HH:MM"
 * @returns {number} Time in decimal hours
 */
function convertTimeToDecimal(timeString) {
  if (!timeString) return null;
  
  const [hours, minutes] = timeString.split(':').map(num => parseInt(num, 10));
  if (isNaN(hours) || isNaN(minutes)) return null;
  
  // Convert to decimal: hours + (minutes/60)
  const decimalTime = hours + (minutes / 60);
  Logger.log(decimalTime)
  return Number(decimalTime.toFixed(2)); // Round to 2 decimal places
}

function createTriggers() {
  // 1. onOpen Trigger
  ScriptApp.newTrigger("onOpen")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onOpen()
    .create();

  // 2. refreshFolderLinksAndAssignees Trigger (daily between 1 AM and 2 AM)
  ScriptApp.newTrigger("refreshFolderLinksAndAssignees")
    .timeBased() // Call timeBased() directly on the TriggerBuilder
    .everyDays(1)
    .atHour(1)
    .create();

  // 3. readDocToMitigationMatrix Trigger (daily between 2 AM and 3 AM)
  ScriptApp.newTrigger("readDocToMitigationMatrix")
    .timeBased() // Call timeBased() directly on the TriggerBuilder
    .everyDays(1)
    .atHour(2)
    .create();
}
