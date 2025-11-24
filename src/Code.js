/**
 * Google Drive Duplicator
 *
 * Logic to duplicate a folder structure from a Source ID to a Destination.
 * Handles execution time limits by saving state and allowing resumption.
 *
 * Update: Uses Drive Files to store state (Queue) to avoid PropertiesService limits.
 */

// Configuration
var CONFIG = {
  TIME_LIMIT_MS: 5.5 * 60 * 1000, // 5.5 minutes
  STATE_FILE_NAME: 'drive_duplicator_state.json',
  STATE_FILE_ID_KEY: 'DRIVE_DUPLICATOR_STATE_ID'
};

/**
 * Adds a custom menu to the active spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Drive Duplicator')
      .addItem('Start / Resume Copy', 'startCopy')
      .addItem('Verify Folder', 'verifyCopy')
      .addSeparator()
      .addItem('Reset / Clear Memory', 'resetMemory')
      .addToUi();
}

/**
 * Main function to initiate or resume the copy process.
 */
function startCopy() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // 1. Check if we have a saved job
  var state = loadState();

  if (state) {
    // We are in the middle of a job
    var rowIndex = state.rowIndex;
    var status = sheet.getRange(rowIndex, 2).getValue();

    if (status === 'Done') {
      clearState();
      startCopy(); // Restart fresh
      return;
    }

    ui.alert('Resuming copy for row ' + rowIndex);
    processQueue(sheet, rowIndex, state.queue);

  } else {
    // Find a new job
    var data = sheet.getDataRange().getValues();
    // Headers are row 1 (index 0), so start at row 2 (index 1)
    for (var i = 1; i < data.length; i++) {
      var sourceId = data[i][0];
      var status = data[i][1];

      if (sourceId && (status === '' || status === 'Pending')) {
        var rowIndex = i + 1; // 1-based index

        // Initialize the job
        if (!setupJob(sheet, rowIndex, sourceId)) {
          return; // Error occurred
        }

        // Load the newly created queue
        state = loadState();
        if (state) {
            processQueue(sheet, rowIndex, state.queue);
        }
        return;
      }
    }

    ui.alert('No pending jobs found. Please add a Source Folder ID and clear the Status column.');
  }
}

/**
 * Prepares the destination folder and initial queue.
 */
function setupJob(sheet, rowIndex, sourceId) {
  var ui = SpreadsheetApp.getUi();

  try {
    var sourceFolder = DriveApp.getFolderById(sourceId);
    var sourceName = sourceFolder.getName();
    var destName = sourceName + " duplicate";

    sheet.getRange(rowIndex, 2).setValue('Initializing...');

    // Create root destination folder
    var destFolder = DriveApp.getRootFolder().createFolder(destName);
    var destId = destFolder.getId();

    sheet.getRange(rowIndex, 3).setValue(destFolder.getUrl());

    // Initial Queue: [ { sourceId: ..., destId: ... } ]
    var queue = [{ sourceId: sourceId, destId: destId }];

    // Save State
    saveState(queue, rowIndex);

    return true;

  } catch (e) {
    sheet.getRange(rowIndex, 2).setValue('Error: ' + e.message);
    ui.alert('Error accessing source folder: ' + e.message);
    return false;
  }
}

/**
 * Processes the queue until finished or time runs out.
 */
function processQueue(sheet, rowIndex, queue) {
  var startTime = Date.now();

  sheet.getRange(rowIndex, 2).setValue('Processing...');
  SpreadsheetApp.flush(); // Update UI

  try {
    while (queue.length > 0) {
      // Check Time Limit
      if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
        saveState(queue, rowIndex);
        sheet.getRange(rowIndex, 2).setValue('Time Limit - Resume needed');
        return;
      }

      // Peek at the first item
      var currentItem = queue[0];
      var sourceId = currentItem.sourceId;
      var destId = currentItem.destId;

      var sourceFolder = DriveApp.getFolderById(sourceId);
      var destFolder = DriveApp.getFolderById(destId);

      // 1. Copy Files
      var files = sourceFolder.getFiles();
      while (files.hasNext()) {
        if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
          saveState(queue, rowIndex);
          sheet.getRange(rowIndex, 2).setValue('Time Limit - Resume needed');
          return;
        }

        var file = files.next();
        // Check if file already exists (Idempotency)
        var existing = destFolder.getFilesByName(file.getName());
        if (!existing.hasNext()) {
          file.makeCopy(file.getName(), destFolder);
        }
      }

      // 2. Prepare Subfolders (Add to queue)
      if (!currentItem.childrenQueued) {
        var subIter = sourceFolder.getFolders();
        while (subIter.hasNext()) {
           // We do check timer here loosely to avoid massive blocking
           if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
              saveState(queue, rowIndex);
              sheet.getRange(rowIndex, 2).setValue('Time Limit - Resume needed');
              return;
           }

           var sub = subIter.next();
           var subName = sub.getName();

           var dSub;
           var dSubIter = destFolder.getFoldersByName(subName);
           if (dSubIter.hasNext()) dSub = dSubIter.next();
           else dSub = destFolder.createFolder(subName);

           queue.push({ sourceId: sub.getId(), destId: dSub.getId() });
        }
        currentItem.childrenQueued = true;
      }

      // Remove item after processing
      queue.shift();
    }

    // Finished
    sheet.getRange(rowIndex, 2).setValue('Done');
    clearState();
    SpreadsheetApp.flush();

  } catch (e) {
    sheet.getRange(rowIndex, 2).setValue('Error: ' + e.toString());
    saveState(queue, rowIndex);
  }
}

// ----------------------------------------------------------------------------
// STATE MANAGEMENT (File Based)
// ----------------------------------------------------------------------------

function saveState(queue, rowIndex) {
  var props = PropertiesService.getDocumentProperties();
  var fileId = props.getProperty(CONFIG.STATE_FILE_ID_KEY);
  var content = JSON.stringify({ queue: queue, rowIndex: rowIndex });

  if (fileId) {
    try {
      var file = DriveApp.getFileById(fileId);
      file.setContent(content);
      return;
    } catch (e) {
      // File might be deleted, create new
    }
  }

  var file = DriveApp.getRootFolder().createFile(CONFIG.STATE_FILE_NAME, content, MimeType.PLAIN_TEXT);
  props.setProperty(CONFIG.STATE_FILE_ID_KEY, file.getId());
}

function loadState() {
  var props = PropertiesService.getDocumentProperties();
  var fileId = props.getProperty(CONFIG.STATE_FILE_ID_KEY);

  if (!fileId) return null;

  try {
    var file = DriveApp.getFileById(fileId);
    var content = file.getBlob().getDataAsString();
    return JSON.parse(content);
  } catch (e) {
    return null;
  }
}

function clearState() {
  var props = PropertiesService.getDocumentProperties();
  var fileId = props.getProperty(CONFIG.STATE_FILE_ID_KEY);

  if (fileId) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
    } catch (e) {
      // Ignore if already deleted
    }
    props.deleteProperty(CONFIG.STATE_FILE_ID_KEY);
  }
}

function resetMemory() {
  clearState();
  SpreadsheetApp.getUi().alert('Memory cleared. You can start a new copy.');
}

// ----------------------------------------------------------------------------
// VERIFICATION LOGIC
// ----------------------------------------------------------------------------

function verifyCopy() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var row = sheet.getActiveCell().getRow();

  var sourceId = sheet.getRange(row, 1).getValue();
  var destUrl = sheet.getRange(row, 3).getValue();

  if (!sourceId || !destUrl) {
    SpreadsheetApp.getUi().alert('Please select a row with a valid Source ID and Destination URL.');
    return;
  }

  try {
    var destId = destUrl.match(/[-\w]{25,}/);
    if (!destId) throw new Error("Invalid Dest URL");

    sheet.getRange(row, 4).setValue("Verifying...");
    SpreadsheetApp.flush();

    // Warning: Simple recursion might timeout on huge folders
    var sourceCount = countFiles(DriveApp.getFolderById(sourceId));
    var destCount = countFiles(DriveApp.getFolderById(destId[0]));

    var msg = "Source: " + sourceCount + " | Dest: " + destCount;
    if (sourceCount === destCount) msg = "OK (" + sourceCount + ")";
    else msg = "MISMATCH: " + msg;

    sheet.getRange(row, 4).setValue(msg);

  } catch (e) {
    var errorMsg = e.message;
    if (e.toString().indexOf("Exceeded maximum execution time") !== -1) {
       errorMsg = "Timeout (Folder too big)";
    }
    sheet.getRange(row, 4).setValue("Error: " + errorMsg);
  }
}

function countFiles(folder) {
  var count = 0;

  var files = folder.getFiles();
  while (files.hasNext()) {
    files.next();
    count++;
  }

  var subs = folder.getFolders();
  while (subs.hasNext()) {
    count += countFiles(subs.next());
  }

  return count;
}
