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
      .addItem('Start / Resume Verify', 'verifyCopy')
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
    var colStatus = (state.type === 'VERIFY') ? 4 : 2; // Column D or B
    var status = sheet.getRange(rowIndex, colStatus).getValue();
    
    if (status.indexOf('Done') === 0) { // Starts with Done
      clearState();
      deleteResumeTriggers();
      // Only restart if it was a copy job, otherwise stop (don't auto-restart verify)
      if (state.type !== 'VERIFY') {
        startCopy();
      }
      return;
    }
    
    if (state.type === 'VERIFY') {
      processVerifyQueue(sheet, rowIndex, state);
    } else {
      processQueue(sheet, rowIndex, state);
    }
    
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
            processQueue(sheet, rowIndex, state);
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
    
    // Initial Queue: [ { sourceId: ..., destId: ..., stage: 'FILES' } ]
    var queue = [{ sourceId: sourceId, destId: destId, stage: 'FILES' }];
    
    // Save State
    saveState({ queue: queue, rowIndex: rowIndex, totalCopied: 0 });
    
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
function processQueue(sheet, rowIndex, state) {
  var queue = state.queue;
  var totalCopied = state.totalCopied || 0;
  var startTime = Date.now();
  
  sheet.getRange(rowIndex, 2).setValue('Processing... (' + totalCopied + ' files)');
  SpreadsheetApp.flush(); // Update UI
  
  try {
    while (queue.length > 0) {
      var currentItem = queue[0];
      
      // Initialize stage if not set
      if (!currentItem.stage) currentItem.stage = 'FILES';

      var sourceFolder = DriveApp.getFolderById(currentItem.sourceId);
      var destFolder = DriveApp.getFolderById(currentItem.destId);
      
      // --- STAGE: FILES ---
      if (currentItem.stage === 'FILES') {
        var files;
        if (currentItem.fileToken) {
          try {
            files = DriveApp.continueFileIterator(currentItem.fileToken);
          } catch (e) {
             // Token expired or invalid
             logToSheet(rowIndex, 'WARN', 'File Iterator', 'Token expired. Restarting folder files.');
             files = sourceFolder.getFiles();
          }
        } else {
          files = sourceFolder.getFiles();
        }
        
        while (files.hasNext()) {
          // Check Time Limit
          if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
            currentItem.fileToken = files.getContinuationToken();
            saveState({ queue: queue, rowIndex: rowIndex, totalCopied: totalCopied });
            createResumeTrigger();
            sheet.getRange(rowIndex, 2).setValue('Pausing... (' + totalCopied + ' files)');
            return;
          }

          var file = files.next();
          try {
            // Check if file already exists (Idempotency)
            var existing = destFolder.getFilesByName(file.getName());
            if (!existing.hasNext()) {
              file.makeCopy(file.getName(), destFolder);
              totalCopied++;
            } else {
              // We count existing files as "processed" to give accurate progress feeling
              totalCopied++;
            }

            // Update UI periodically
            if (totalCopied % 20 === 0) {
              sheet.getRange(rowIndex, 2).setValue('Processing... (' + totalCopied + ' files)');
            }
          } catch (fileErr) {
            logToSheet(rowIndex, 'ERROR', file.getName(), fileErr.message);
          }
        }

        // Done with files, move to folders
        currentItem.stage = 'FOLDERS';
        delete currentItem.fileToken;
      }
      
      // --- STAGE: FOLDERS ---
      if (currentItem.stage === 'FOLDERS') {
         var folders;
         if (currentItem.folderToken) {
           try {
             folders = DriveApp.continueFolderIterator(currentItem.folderToken);
           } catch (e) {
             logToSheet(rowIndex, 'WARN', 'Folder Iterator', 'Token expired. Restarting folder subs.');
             folders = sourceFolder.getFolders();
           }
         } else {
           folders = sourceFolder.getFolders();
         }

         while (folders.hasNext()) {
            // Check Time Limit
            if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
              currentItem.folderToken = folders.getContinuationToken();
              saveState({ queue: queue, rowIndex: rowIndex, totalCopied: totalCopied });
              createResumeTrigger();
              sheet.getRange(rowIndex, 2).setValue('Pausing... (' + totalCopied + ' files)');
              return;
            }

            var sub = folders.next();
            try {
              var subName = sub.getName();
              var dSub;
              var dSubIter = destFolder.getFoldersByName(subName);
              if (dSubIter.hasNext()) {
                dSub = dSubIter.next();
              } else {
                dSub = destFolder.createFolder(subName);
              }

              // Add to queue
              queue.push({ sourceId: sub.getId(), destId: dSub.getId(), stage: 'FILES' });

            } catch (folderErr) {
              logToSheet(rowIndex, 'ERROR', sub.getName(), folderErr.message);
            }
         }

         // Done with folders
         queue.shift(); // Remove current item
      }
    }
    
    // Finished
    deleteResumeTriggers();
    var finalStatus = 'Done. (' + totalCopied + ' files processed)';
    sheet.getRange(rowIndex, 2).setValue(finalStatus);
    clearState();

    // Run verification if requested (but safe version?)
    // For now, let's just log completion.
    logToSheet(rowIndex, 'INFO', 'Job', 'Copy Completed. Files: ' + totalCopied);
    SpreadsheetApp.flush();
    
  } catch (e) {
    sheet.getRange(rowIndex, 2).setValue('Error: ' + e.toString());
    logToSheet(rowIndex, 'CRITICAL', 'ProcessQueue', e.toString());
    saveState({ queue: queue, rowIndex: rowIndex, totalCopied: totalCopied });
  }
}

// ----------------------------------------------------------------------------
// STATE MANAGEMENT (File Based)
// ----------------------------------------------------------------------------

function saveState(stateData) {
  var props = PropertiesService.getDocumentProperties();
  var fileId = props.getProperty(CONFIG.STATE_FILE_ID_KEY);
  var content = JSON.stringify(stateData);
  
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
  deleteResumeTriggers();
  SpreadsheetApp.getUi().alert('Memory cleared and triggers removed. You can start a new copy or verification.');
}

// ----------------------------------------------------------------------------
// VERIFICATION LOGIC
// ----------------------------------------------------------------------------

function verifyCopy() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();
  var row = sheet.getActiveCell().getRow();
  
  // Check if busy
  if (loadState()) {
    ui.alert("A job (Copy or Verify) is already in progress. Please wait for it to finish or reset memory.");
    return;
  }

  var sourceId = sheet.getRange(row, 1).getValue();
  var destUrl = sheet.getRange(row, 3).getValue();
  
  if (!sourceId || !destUrl) {
    ui.alert('Please select a row with a valid Source ID and Destination URL.');
    return;
  }
  
  try {
    var destIdMatch = destUrl.match(/[-\w]{25,}/);
    if (!destIdMatch) throw new Error("Invalid Dest URL");
    var destId = destIdMatch[0];
    
    sheet.getRange(row, 4).setValue("Initializing Verification...");
    SpreadsheetApp.flush();
    
    // Initialize Verify Job
    var queue = [{ sourceId: sourceId, destId: destId, stage: 'FILES' }];
    var state = {
      type: 'VERIFY',
      queue: queue,
      rowIndex: row,
      checked: 0,
      mismatches: 0
    };
    
    saveState(state);
    processVerifyQueue(sheet, row, state);
    
  } catch (e) {
    sheet.getRange(row, 4).setValue("Error: " + e.message);
  }
}

function processVerifyQueue(sheet, rowIndex, state) {
  var queue = state.queue;
  var checked = state.checked || 0;
  var mismatches = state.mismatches || 0;
  var startTime = Date.now();

  sheet.getRange(rowIndex, 4).setValue('Verifying... (' + checked + ' checked)');
  SpreadsheetApp.flush();

  try {
    while (queue.length > 0) {
      var currentItem = queue[0];
      if (!currentItem.stage) currentItem.stage = 'FILES';

      var sourceFolder = DriveApp.getFolderById(currentItem.sourceId);
      var destFolder = DriveApp.getFolderById(currentItem.destId);

      // --- STAGE: FILES ---
      if (currentItem.stage === 'FILES') {
        var files;
        if (currentItem.fileToken) {
           try {
             files = DriveApp.continueFileIterator(currentItem.fileToken);
           } catch (e) {
             logToSheet(rowIndex, 'WARN', 'Verify-Files', 'Token expired, restarting folder.');
             files = sourceFolder.getFiles();
           }
        } else {
           files = sourceFolder.getFiles();
        }

        while (files.hasNext()) {
          // Time Limit
          if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
            currentItem.fileToken = files.getContinuationToken();
            saveState(state);
            createResumeTrigger();
            sheet.getRange(rowIndex, 4).setValue('Verifying... (Resume needed)');
            return;
          }

          var file = files.next();
          checked++;

          // Verify existence
          var existing = destFolder.getFilesByName(file.getName());
          if (!existing.hasNext()) {
             mismatches++;
             logToSheet(rowIndex, 'MISSING', file.getName(), 'File missing in destination');
          }
        }

        currentItem.stage = 'FOLDERS';
        delete currentItem.fileToken;
      }

      // --- STAGE: FOLDERS ---
      if (currentItem.stage === 'FOLDERS') {
        var folders;
        if (currentItem.folderToken) {
           try {
             folders = DriveApp.continueFolderIterator(currentItem.folderToken);
           } catch (e) {
             logToSheet(rowIndex, 'WARN', 'Verify-Folders', 'Token expired, restarting folder.');
             folders = sourceFolder.getFolders();
           }
        } else {
           folders = sourceFolder.getFolders();
        }

        while (folders.hasNext()) {
           // Time Limit
           if (Date.now() - startTime > CONFIG.TIME_LIMIT_MS) {
            currentItem.folderToken = folders.getContinuationToken();
            saveState(state);
            createResumeTrigger();
            sheet.getRange(rowIndex, 4).setValue('Verifying... (Resume needed)');
            return;
          }

          var sub = folders.next();
          var subName = sub.getName();

          var dSubIter = destFolder.getFoldersByName(subName);
          if (dSubIter.hasNext()) {
            var dSub = dSubIter.next();
            queue.push({ sourceId: sub.getId(), destId: dSub.getId(), stage: 'FILES' });
          } else {
            mismatches++;
            logToSheet(rowIndex, 'MISSING_DIR', subName, 'Folder missing in destination');
          }
        }

        queue.shift();
      }
    }

    // Done
    deleteResumeTriggers();
    clearState();

    var resultMsg = "Done. " + checked + " files checked.";
    if (mismatches > 0) {
      resultMsg += " " + mismatches + " ISSUES (See Logs)";
    } else {
      resultMsg += " Perfect Match.";
    }
    sheet.getRange(rowIndex, 4).setValue(resultMsg);

  } catch (e) {
    sheet.getRange(rowIndex, 4).setValue('Error: ' + e.message);
    logToSheet(rowIndex, 'CRITICAL', 'Verify', e.toString());
    saveState(state);
  }
}

// ----------------------------------------------------------------------------
// LOGGING & TRIGGERS
// ----------------------------------------------------------------------------

/**
 * Ensures the "Logs" sheet exists.
 */
function ensureLogsSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Logs");
  if (!sheet) {
    sheet = ss.insertSheet("Logs");
    sheet.appendRow(["Timestamp", "Job Row", "Type", "File/Folder", "Message"]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Logs an entry to the "Logs" sheet.
 */
function logToSheet(jobRowIndex, type, name, message) {
  try {
    var sheet = ensureLogsSheet();
    sheet.appendRow([new Date(), jobRowIndex, type, name, message]);
  } catch (e) {
    console.error("Failed to log to sheet: " + e.message);
  }
}

/**
 * Creates a one-time trigger to resume execution after 1 minute.
 */
function createResumeTrigger() {
  // Avoid duplicate triggers
  deleteResumeTriggers();
  ScriptApp.newTrigger('startCopy')
      .timeBased()
      .after(60 * 1000) // 1 minute
      .create();
}

/**
 * Deletes any existing triggers for 'startCopy'.
 */
function deleteResumeTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'startCopy') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}
