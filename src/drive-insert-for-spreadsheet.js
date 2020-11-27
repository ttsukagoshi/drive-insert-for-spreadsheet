// Copyright 2020 Taro TSUKAGOSHI
// 
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
// 
//     http://www.apache.org/licenses/LICENSE-2.0
// 
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// 
// For latest information, see https://github.com/ttsukagoshi/drive-insert-for-spreadsheet

// Boolean indicating the mode of this script; `true` for Editor Add-on, `false` when used as a spreadsheet-bound script.
const IS_EDITOR_ADDON = true;

function onOpen() {
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var localizedMessage = new LocalizedMessage(locale);
  var ui = SpreadsheetApp.getUi();
  if (IS_EDITOR_ADDON) {
    ui.createAddonMenu()
      .addItem(localizedMessage.messageList.menuInsertImage, 'insertImage')
      .addSeparator()
      .addItem(localizedMessage.messageList.menuSetup, 'setParameters')
      .addItem(localizedMessage.messageList.menuCheckSettings, 'checkParameters')
      .addToUi();
  } else {
    ui.createMenu(localizedMessage.messageList.menuTitle)
      .addItem(localizedMessage.messageList.menuInsertImage, 'insertImage')
      .addSeparator()
      .addItem(localizedMessage.messageList.menuSetup, 'setParameters')
      .addItem(localizedMessage.messageList.menuCheckSettings, 'checkParameters')
      .addToUi();
  }
}

function onInstall() {
  onOpen();
}

function insertImage() {
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var documentProperties = PropertiesService.getDocumentProperties().getProperties();
  var localizedMessage = new LocalizedMessage(locale);
  var ui = SpreadsheetApp.getUi();
  try {
    let isSettingComplete = (
      documentProperties.folderId
      && documentProperties.fileExt
      && documentProperties.selectionVertical
      && documentProperties.insertPosNext
    );
    if (!isSettingComplete) {
      throw new Error(localizedMessage.messageList.errorInitialSettingNotComplete);
    }
    let folderId = documentProperties.folderId;
    let activeSheet = SpreadsheetApp.getActiveSheet();
    let selectedRange = SpreadsheetApp.getActiveRange();
    let options = {
      fileExt: documentProperties.fileExt,
      selectionVertical: toBoolean_(documentProperties.selectionVertical),
      insertPosNext: toBoolean_(documentProperties.insertPosNext)
    };
    let result = insertImageFromDrive(folderId, activeSheet, selectedRange, options);
    let message = localizedMessage.replaceAlertMessageOnComplete(result.getBlobsCompleteSec, result.insertImageCompleteSec);
    for (let k in result) {
      if (k == 'getBlobsCompleteSec' || k == 'insertImageCompleteSec') {
        continue;
      } else if (result[k] <= 1) {
        continue;
      } else {
        message += localizedMessage.replaceAlertMessageAdd(k, result[k]);
      }
    }
    ui.alert(localizedMessage.messageList.alertMessageOnCompleteTitle, message, ui.ButtonSet.OK);
  } catch (error) {
    let errorMessage = errorMessage_(error);
    ui.alert(errorMessage);
  }
}

/**
 * Convert string booleans into boolean
 * @param {string} stringBoolean 
 * @return {boolean}
 */
function toBoolean_(stringBoolean) {
  return stringBoolean.toLowerCase() === 'true';
}

/**
 * Insert image blobs obtained from a designated Google Drive folder.
 * @param {string} folderId ID of Google Drive folder where the images are stored. If this value is "root", the root Drive folder is selected.
 * @param {Object} activeSheet Sheet class object of the Google Spreadsheet to insert the image.
 * @param {Object} selectedRange Range class object of Google Spreadsheet that contains the image file names.
 * @param {Object} options Advanced parameters.
 * @param {string} options.fileExt File extension to search for in the Google Drive folder. Defaults to 'jpg'.
 * Note that the period before the extension is NOT required.
 * @param {Boolean} options.selectionVertical Direction of cell selection. Assumes it is vertical when true, as by default.
 * @param {Boolean} options.insertPosNext Position to insert the image. 
 * When true, as by default, the image will be inserted in the next column (or row, if selectionVertical is false)
 * of the selected cells.
 * @returns {Object} Object with file name as the key and the number of files in the Drive folder with the same name as its value.
 */
function insertImageFromDrive(folderId, activeSheet, selectedRange, options = {}) {
  var start = new Date();
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var localizedMessage = new LocalizedMessage(locale);
  var result = {};
  try {
    // Check the selected range and get file names
    let rangeNumRows = selectedRange.getNumRows();
    let rangeNumColumns = selectedRange.getNumColumns();
    let fileNames = [];
    if ((options.selectionVertical && rangeNumColumns == 1) || (!options.selectionVertical && rangeNumRows == 1)) {
      fileNames = fileNames.concat(selectedRange.getValues().flat());
    } else if (options.selectionVertical && rangeNumColumns > 1) {
      throw new Error(localizedMessage.messageList.errorMoreThanOneColumnSelected);
    } else if (!options.selectionVertical && rangeNumRows > 1) {
      throw new Error(localizedMessage.messageList.errorMoreThanOneRowSelected);
    } else if (selectedRange.isBlank()) {
      throw new Error(localizedMessage.messageList.errorEmptyCellsSelected);
    } else {
      let errorMessage = localizedMessage.replaceErrorUnknownError(selectedRange.getA1Notation(), options.selectionVertical, rangeNumRows, rangeNumColumns);
      throw new Error(errorMessage);
    }
    // Get images as blobs
    let targetFolder = (folderId == 'root' ? DriveApp.getRootFolder() : DriveApp.getFolderById(folderId));
    let imageBlobs = fileNames.map((value) => {
      let fileNameExt = `${value}.${options.fileExt}`;
      let targetFile = targetFolder.getFilesByName(fileNameExt);
      let fileCounter = 0;
      let fileBlob = null;
      while (targetFile.hasNext()) {
        let file = targetFile.next();
        fileCounter += 1;
        if (fileCounter <= 1) { // Get the first image file with the designated file name, and ignore all others.
          fileBlob = file.getBlob().setName(value);
          let imageInfo = ImgAppR.getSize(fileBlob);
          if (['JPG', 'PNG', 'GIF'].indexOf(imageInfo.identification) < 0) {
            throw new Error(localizedMessage.replaceErrorImageFileFormat(fileNameExt));
          } else if ((imageInfo.height * imageInfo.width) > 1048576) {
            throw new Error(localizedMessage.replaceErrorImageFileSizeExceedsLimit(fileNameExt));
          }
        }
      }
      result[value] = fileCounter;
      return fileBlob;
    });
    let getBlobsComplete = new Date();
    result['getBlobsCompleteSec'] = (getBlobsComplete.getTime() - start.getTime()) / 1000;
    // Set the offset row and column to insert image blobs
    let offsetPos = (options.insertPosNext ? 1 : -1);
    let offsetPosRow = (options.selectionVertical ? 0 : offsetPos);
    let offsetPosCol = (options.selectionVertical ? offsetPos : 0);
    // Define the range to insert image
    let insertRange = selectedRange.offset(offsetPosRow, offsetPosCol);
    // Verify the contents of the insertRange, i.e., make sure the cells in the range are empty
    if (!insertRange.isBlank()) {
      throw new Error(localizedMessage.messageList.errorExistingContentInInsertCellRange);
    }
    // Insert the image blobs
    let startCell = { 'row': insertRange.getRow(), 'column': insertRange.getColumn() };
    let cellPxSizes = cellPixSizes_(activeSheet, insertRange).flat();
    imageBlobs.forEach(function (blob, index) {
      let img = (
        options.selectionVertical
          ? activeSheet.insertImage(blob, startCell.column, startCell.row + index)
          : activeSheet.insertImage(blob, startCell.column + index, startCell.row)
      );
      let [imgHeight, imgWidth] = [img.getHeight(), img.getWidth()];
      let { height, width } = cellPxSizes[index];
      let fraction = Math.min(height / imgHeight, width / imgWidth);
      let [imgHeightResized, imgWidthResized] = [imgHeight * fraction, imgWidth * fraction];
      let offsetX = Math.trunc((width - imgWidthResized) / 2);
      img.setHeight(imgHeightResized).setWidth(imgWidthResized).setAnchorCellXOffset(offsetX);
    });
    let insertImageComplete = new Date();
    result['insertImageCompleteSec'] = (insertImageComplete.getTime() - start.getTime()) / 1000;
    return result;
  } catch (error) {
    throw error;
  }
}

/**
 * Gets the cells' height and width in pixels for the selected range in Google Spreadsheet in form of a 2-d JavaScript array;
 * the array values are ordered in the same way as executing Range.getValues()
 * @param {Object} activeSheet The active Sheet class object in Google Spreadsheet, e.g., SpreadsheetApp.getActiveSheet()
 * @param {Object} activeRange The selected Range class object in Google Spreadsheet, e.g., SpreadsheetApp.getActiveRange()
 * @returns {array} 2-d array of objects with 'height' and 'width' as keys and pixels as values.
 * Each object represents a cell and is aligned in the same order as Range.getValues()
 */
function cellPixSizes_(activeSheet, activeRange) {
  var rangeStartCell = { 'row': activeRange.getRow(), 'column': activeRange.getColumn() };
  var cellValues = activeRange.getValues();
  var cellPixSizes = cellValues.map(function (row, rowIndex) {
    let rowPix = activeSheet.getRowHeight(rangeStartCell.row + rowIndex);
    let rowPixSizes = row.map(function (cell, colIndex) {
      let cellSize = { 'height': rowPix, 'width': activeSheet.getColumnWidth(rangeStartCell.column + colIndex) };
      return cellSize;
    });
    return rowPixSizes;
  });
  return cellPixSizes;
}

/**
 * Prompt to set intial settings
 */
function setParameters() {
  var ui = SpreadsheetApp.getUi();
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var localizedMessage = new LocalizedMessage(locale);
  var documentProperties = PropertiesService.getDocumentProperties().getProperties();
  if (!documentProperties.setupComplete || documentProperties.setupComplete == 'false') {
    setup_(ui);
  } else {
    let alreadySetupMessage = localizedMessage.messageList.alertAlreadySetupMessage;
    for (let k in documentProperties) {
      alreadySetupMessage += `${k}: ${documentProperties[k]}\n`;
    }
    let response = ui.alert(alreadySetupMessage, ui.ButtonSet.YES_NO);
    if (response == ui.Button.YES) {
      setup_(ui, documentProperties);
    }
  }
}

/**
 * Sets the required script properties
 * @param {Object} ui Apps Script Ui class object, as retrieved by SpreadsheetApp.getUi()
 * @param {Object} currentSettings [Optional] Current script properties
 */
function setup_(ui, currentSettings = {}) {
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var localizedMessage = new LocalizedMessage(locale);
  try {
    // folderId
    let promptFolderId = localizedMessage.messageList.promptFolderId;
    promptFolderId += (currentSettings.folderId ? localizedMessage.replacePromptCurrentValue(currentSettings.folderId) : '');
    let responseFolderId = ui.prompt(promptFolderId, ui.ButtonSet.OK_CANCEL);
    if (responseFolderId.getSelectedButton() !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
    }
    let folderId = responseFolderId.getResponseText();

    // fileExt
    let promptFileExt = localizedMessage.messageList.promptFileExt;
    promptFileExt += (currentSettings.fileExt ? localizedMessage.replacePromptCurrentValue(currentSettings.fileExt) : '');
    let responseFileExt = ui.prompt(promptFileExt, ui.ButtonSet.OK_CANCEL);
    if (responseFileExt.getSelectedButton() !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
    }
    let fileExt = responseFileExt.getResponseText();

    // selectionVertical
    let promptSelectionVertical = localizedMessage.messageList.promptSelectionVertical;
    promptSelectionVertical += (currentSettings.selectionVertical ? localizedMessage.replacePromptCurrentValue(currentSettings.selectionVertical) : '');
    let responseSelectionVertical = ui.prompt(promptSelectionVertical, ui.ButtonSet.OK_CANCEL);
    if (responseSelectionVertical.getSelectedButton() !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
    }
    let selectionVertical = responseSelectionVertical.getResponseText();

    // insertPosNext
    let promptInsertPosNext = localizedMessage.messageList.promptInsertPosNext;
    promptInsertPosNext += (currentSettings.insertPosNext ? localizedMessage.replacePromptCurrentValue(currentSettings.insertPosNext) : '');
    let responseInsertPosNext = ui.prompt(promptInsertPosNext, ui.ButtonSet.OK_CANCEL);
    if (responseInsertPosNext.getSelectedButton() !== ui.Button.OK) {
      throw new Error(localizedMessage.messageList.errorCanceled);
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
    PropertiesService.getDocumentProperties().setProperties(properties, false);
    ui.alert(localizedMessage.messageList.alertSetupComplete);
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
  var locale = (IS_EDITOR_ADDON ? Session.getActiveUserLocale() : SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetLocale());
  var localizedMessage = new LocalizedMessage(locale);
  var documentProperties = PropertiesService.getDocumentProperties().getProperties();
  var currentSettings = '';
  for (let k in documentProperties) {
    currentSettings += `${k}: ${documentProperties[k]}\n`;
  }
  ui.alert(localizedMessage.messageList.alertCurrentSettingsTitle, currentSettings, ui.ButtonSet.OK);
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