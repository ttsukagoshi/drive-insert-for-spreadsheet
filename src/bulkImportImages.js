function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Test')//////////////////////////////////////
    .addItem('test', 'test')
    .addToUi();
  ui.createMenu('Insert Image from Drive')
    .addItem('Insert Image', 'insertImage')
    .addSeparator()
    .addItem('Setup', 'setParameters')
    .addItem('Check Settings', 'checkParameters')
    .addToUi();
}

function insertImage() {
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var ui = SpreadsheetApp.getUi();
  try {
    var isSettingComplete = (
      !scriptProperties.folderId ||
      !scriptProperties.fileExt ||
      !scriptProperties.selectionVertical ||
      !scriptProperties.insertPosNext
    );
    if (isSettingComplete) {
      throw new Error('Initial settings is not complete. Try running menu "Insert Image from Drive" > "Setup"')
    }
    let folderId = scriptProperties.folderId;
    let selectedRange = SpreadsheetApp.getActiveRange();
    let options = {
      fileExt: scriptProperties.fileExt,
      selectionVertical: toBoolean_(scriptProperties.selectionVertical),
      insertPosNext: toBoolean_(scriptProperties.insertPosNext)
    };
    insertImageFromDrive(folderId, selectedRange, options);
  } catch (error) {
    let errorMessage = errorMessage_(error);
    ui.alert(errorMessage);
  }
}

/**
 * Convert string booleans into boolean data
 * @param {string} stringBoolean 
 * @return {boolean}
 */
function toBoolean_(stringBoolean) {
  return stringBoolean.toLowerCase() === 'true';
}

/**
 * Insert image blobs obtained from a designated Google Drive folder.
 * @param {string} folderId ID of Google Drive folder
 * @param {Object} selectedRange Range class object of Google Spreadsheet.
 * @param {Object} options Advanced parameters.
 * @param {string} options.fileExt File extension to search for in the Google Drive folder. Defaults to 'jpg'.
 * Note that the period before the extension is NOT required.
 * @param {Boolean} options.selectionVertical Direction of cell selection. Assumes it is vertical when true, as by default.
 * @param {Boolean} options.insertPosNext Position to insert the image. 
 * When true, as by default, the image will be inserted in the next column (or row, if selectionVertical is false)
 * of the selected cells.
 * @returns {Object} Object with file name as the key and the number of files in the Drive folder with the same name as its value.
 */
function insertImageFromDrive(folderId, selectedRange, options = {}) {
  // Set default values of options
  options.fileExt = options.fileExt || 'jpg';
  options.selectionVertical = options.selectionVertical || true;
  options.insertPosNext = options.insertPosNext || true;
  // Define the object to return.
  var result = {};
  try {
    // Check the selected range and get file names
    let rangeNumRows = selectedRange.getNumRows();
    let rangeNumColumns = selectedRange.getNumColumns();
    let fileNames = [];
    if (options.selectionVertical && rangeNumColumns > 1) {
      throw new Error('More than one column is selected. Check the selected range.');
    } else if (options.selectionVertical && rangeNumColumns == 1) {
      selectedRange.getValues().forEach(value => fileNames.push(value[0]));
    } else if (!options.selectionVertical && rangeNumRows > 1) {
      throw new Error('More than one row is selected. Check the selected range.');
    } else if (!options.selectionVertical && rangeNumRows == 1) {
      fileNames = fileNames.concat(selectedRange.getValues()[0]);
    } else if (selectedRange.isBlank()) {
      throw new Error('Empty cells. Check the selected range.')
    } else {
      let errorMessage = `Unknown Error:
      Selected Range: ${selectedRange.getA1Notation()}
      options.selectionVertical = ${options.selectionVertical}
      rangeNumRows = ${rangeNumRows}
      rangeNumColumns = ${rangeNumColumns}`;
      throw new Error(errorMessage);
    }
    console.log(fileNames);//////////////////////////////////
    // Get images
    let targetFolder = DriveApp.getFolderById(folderId);
    let imageBlobs = fileNames.map((value) => {
      let fileNameExt = `${value}.${options.fileExt}`;
      let targetFile = targetFolder.getFilesByName(fileNameExt);
      let fileCounter = 0;
      let fileBlob = null;
      while (targetFile.hasNext()) {
        let file = targetFile.next();
        fileCounter += 1;
        if (fileCounter <= 1) {
          fileBlob = file.getBlob();
        }
      }
      result[value] = fileCounter;
      return fileBlob;
    });
    // Set the offset row and column to insert image blobs
    let offsetPos = (options.insertPosNext ? 1 : -1);
    let offsetPosRow = (options.selectionVertical ? 0 : offsetPos);
    let offsetPosCol = (options.selectionVertical ? offsetPos : 0);
    // Define the range to insert image
    let insertRange = selectedRange.offset(offsetPosRow, offsetPosCol);
    // Verify the contents of the insertRange, i.e., make sure the cells in the range are empty
    if (!insertRange.isBlank()) {
      throw new Error('Existing Content in Insert Cell Range: Check the selected range.');
    }
    // Insert the image blobs
    console.log(imageBlobs);/////////////////////////////////////////
    return result;
  } catch (error) {
    throw error;
  }
}

/**
 * Prompt to set intial settings
 */
function setParameters() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  if (!scriptProperties.setupComplete || scriptProperties.setupComplete == 'false') {
    setup_(ui);
  } else {
    let alreadySetupMessage = 'Initial settings are already complete. Do you want to overwrite the settings?\n\n';
    for (let k in scriptProperties) {
      alreadySetupMessage += `${k}: ${scriptProperties[k]}\n`;
    }
    let response = ui.alert(alreadySetupMessage, ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      setup_(ui, scriptProperties);
    }
  }
}

/**
 * Sets the required script properties
 * @param {Object} ui Apps Script Ui class object, as retrieved by SpreadsheetApp.getUi()
 * @param {Object} currentSettings [Optional] Current script properties
 */
function setup_(ui, currentSettings = {}) {
  try {
    // folderId
    let promptFolderId = 'Google Drive folder ID to get the images from.';
    promptFolderId += (currentSettings.folderId ? `\n\nCurrent Value: ${currentSettings.folderId}` : '');
    let responseFolderId = ui.prompt(promptFolderId, ui.ButtonSet.OK_CANCEL);
    if (responseFolderId.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let folderId = responseFolderId.getResponseText();

    // fileExt
    let promptFileExt = 'File extension of the image file(s) without the period. e.g., NOT ".jpg" but "jpg"';
    promptFileExt += (currentSettings.fileExt ? `\n\nCurrent Value: ${currentSettings.fileExt}` : '');
    let responseFileExt = ui.prompt(promptFileExt, ui.ButtonSet.OK_CANCEL);
    if (responseFileExt.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let fileExt = responseFileExt.getResponseText();

    // selectionVertical
    let promptSelectionVertical = 'selectionVertical: Enter "true" or "false". When true, the script will assume that the cells are selected vertically, i.e., in a single column.';
    promptSelectionVertical += (currentSettings.selectionVertical ? `\n\nCurrent Value: ${currentSettings.selectionVertical}` : '');
    let responseSelectionVertical = ui.prompt(promptSelectionVertical, ui.ButtonSet.OK_CANCEL);
    if (responseSelectionVertical.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let selectionVertical = responseSelectionVertical.getResponseText();

    // insertPosNext
    let promptInsertPosNext = 'insertPosNext: Enter "true" or "false". When true, the images will be inserted in the next row or column, depending on the value of selectionVertical.';
    promptInsertPosNext += (currentSettings.insertPosNext ? `\n\nCurrent Value: ${currentSettings.insertPosNext}` : '');
    let responseInsertPosNext = ui.prompt(promptInsertPosNext, ui.ButtonSet.OK_CANCEL);
    if (responseInsertPosNext.getSelectedButton() !== ui.Button.OK) {
      throw new Error('Canceled.');
    }
    let insertPosNext = responseInsertPosNext.getResponseText();

    // Set script properties
    let properties = {
      'folderId': folderId,
      'fileExt': fileExt,
      'selectionVertical': selectionVertical,
      'insertPosNext': insertPosNext,
      'setupComplete': true
    };
    PropertiesService.getScriptProperties().setProperties(properties, false);
    ui.alert('Complete: setup of script properties');
  } catch (error) {
    let message = errorMessage_(error);
    ui.alert(message);
  }
}

/**
 * Shows the list of current script properties.
 */
function checkParameters() {
  var ui = SpreadsheetApp.getUi();
  var scriptProperties = PropertiesService.getScriptProperties().getProperties();
  var currentSettings = '';
  for (let k in scriptProperties) {
    currentSettings += `${k}: ${scriptProperties[k]}\n`;
  }
  ui.alert('Current Settings', currentSettings, ui.ButtonSet.OK);
}

/**
 * Standarized error message
 * @param {Object} error Error object
 * @return {string} Standarized error message
 */
function errorMessage_(error) {
  let message = error.stack;
  return message;
}