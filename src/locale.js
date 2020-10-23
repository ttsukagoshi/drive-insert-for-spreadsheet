// MIT License
// 
// Copyright (c) 2020 Taro TSUKAGOSHI
// https://github.com/ttsukagoshi/spreadsheet-bulk-import-images
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
// 
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.
// 
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

const MESSAGES = {
  'en_US': {
    'menuTitle': 'Insert Image from Drive',
    'menuInsertImage': 'Insert Image',
    'menuSetup': 'Setup',
    'menuCheckSettings': 'Check Settings',
    'errorInitialSettingNotComplete': 'Initial settings is not complete. Try running menu "Insert Image from Drive" > "Setup"',
    'alertMessageOnCompleteTitle': 'Image Insert Complete',
    'alertMessageOnComplete': 'Reload the spreadsheet if the images are not visible.\n\nGetting image from Drive folder: {{getBlobsCompleteSec}} secs\nWhole process completed in {{insertImageCompleteSec}} secs\n\n',
    'alertMessageAdd': '{{key}}: {{value}} files with the same name\n',
    'errorMoreThanOneColumnSelected': 'More than one column is selected. Check the selected range.',
    'errorMoreThanOneRowSelected': 'More than one row is selected. Check the selected range.',
    'errorEmptyCellsSelected': 'Empty cells. Check the selected range.',
    'errorImageFileFormat': 'Check Image File Type: {{fileName}}\nImage file must be either JPG, PNG, or GIF.',
    'errorImageFileSizeExceedsLimit': 'Image Size Exceeds Pixel Size Limit: {{fileName}}\nThe pixel area, or the product of horizontal and vertical pixel length, of the image must be less than or equal to 2^20 = 1,048,576 pixel^2.\nFor details, see\n https://github.com/ttsukagoshi/spreadsheet-bulk-import-images#note',
    'errorUnknownError': 'Unknown Error:\n\nSelected Range: {{selectedRangeA1Notation}}\noptions.selectionVertical = {{optionsSelectionVertical}}\nrangeNumRows = {{rangeNumRows}}\nrangeNumColumns = {{rangeNumColumns}}',
    'errorExistingContentInInsertCellRange': 'Existing Content in Insert Cell Range: Check the selected range.',
    'alertAlreadySetupMessage': 'Initial settings are already complete. Do you want to overwrite the settings?\n\n',
    'promptCurrentValue': '\n\nCurrent Value: {{value}}',
    'errorCanceled': 'Canceled.',
    'promptFolderId': 'Google Drive folder ID to get the images from. Enter "root" if you want to designated your Google Drive\'s root folder.',
    'promptFileExt': 'File extension of the image file(s) without the period. e.g., NOT ".jpg" but "jpg"',
    'promptSelectionVertical': 'selectionVertical: Enter "true" or "false". When true, the script will assume that the cells are selected vertically, i.e., in a single column.',
    'promptInsertPosNext': 'insertPosNext: Enter "true" or "false". When true, the images will be inserted in the next row or column, depending on the value of selectionVertical.',
    'alertSetupComplete': 'Setup of this add-on properties complete.',
    'alertCurrentSettingsTitle': 'Current Settings'
  },
  'ja': {
    'menuTitle': 'Googleドライブから画像挿入',
    'menuInsertImage': '画像挿入',
    'menuSetup': '初期設定',
    'menuCheckSettings': '設定確認',
    'errorInitialSettingNotComplete': '初期設定が完了していません。メニュー「Googleドライブから画像挿入」から「初期設定」を実行してください。',
    'alertMessageOnCompleteTitle': '画像挿入 完了',
    'alertMessageOnComplete': '挿入された画像が見えない場合は、スプレッドシートのウェブページを更新してください。\n\nGoogleドライブからの画像取得：{{getBlobsCompleteSec}}秒\n処理全体は{{insertImageCompleteSec}}秒で完了\n\n',
    'alertMessageAdd': '{{key}}: 同じ名前の画像ファイルが{{value}}個、見つかりました。',
    'errorMoreThanOneColumnSelected': '複数の列が選択されています。選択範囲をご確認ください。',
    'errorMoreThanOneRowSelected': '複数の行が選択されています。選択範囲をご確認ください。',
    'errorEmptyCellsSelected': '選択範囲が空白です。ご確認ください。',
    'errorImageFileFormat': '画像ファイルの形式を確認してください。 {{fileName}}\n画像ファイルはJPG、PNG、またはGIFのいずれかの形式である必要があります。',
    'errorImageFileSizeExceedsLimit': '画像が大きすぎます：{{fileName}}\n縦横のピクセル数の積＝ピクセル面積が2^20 = 1,048,576ピクセル^2以下となるように調整してください。\n詳細は\n https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/blob/main/README.ja.md#%E4%BB%98%E8%A8%98 \nにて確認できます。',
    'errorUnknownError': '不明なエラー：\n\n選択範囲：{{selectedRangeA1Notation}}\noptions.selectionVertical = {{optionsSelectionVertical}}\nrangeNumRows = {{rangeNumRows}}\nrangeNumColumns = {{rangeNumColumns}}',
    'errorExistingContentInInsertCellRange': '処理が中断されました。\n画像を挿入する予定の範囲にデータが存在します。範囲をご確認ください。',
    'alertAlreadySetupMessage': '初期設定はすでに完了しています。既存の設定を上書きしますか？',
    'promptCurrentValue': '\n\n現在の値: {{value}}',
    'errorCanceled': 'キャンセルされました。.',
    'promptFolderId': '画像が保存されているGoogleドライブのフォルダID。「マイドライブ」直下を指定したい場合は「root」と入力してください。',
    'promptFileExt': '画像ファイルの拡張子。ピリオド（.）は除いてください。\n例）「jpg」→○、「.jpg」→×',
    'promptSelectionVertical': 'selectionVertical：選択範囲の縦横方向を指定。「true」（縦方向）または「false」（横方向）を入力。',
    'promptInsertPosNext': 'insertPosNext：画像を挿入する位置を指定。「true」（選択範囲の方向によって右列または下行）または「false」（選択範囲の方向によって左列または上行）を入力。',
    'alertSetupComplete': '設定完了。',
    'alertCurrentSettingsTitle': '現在の設定'
  }
};

class LocalizedMessage {
  constructor(userLocale = 'en_US') {
    this.locale = userLocale;
    this.messageList = (MESSAGES[this.locale] ? MESSAGES[this.locale] : MESSAGES.en_US);
  }

  /**
   * Replace placeholder values in the designated text. String.prototype.replace() is executed using regular expressions with the 'global' flag on.
   * @param {string} text 
   * @param {array} placeholderValues Array of objects containing a placeholder string expressed in regular expression and its corresponding value.
   * @returns {string} The replaced text.
   */
  replacePlaceholders_(text, placeholderValues = []) {
    let replacedText = placeholderValues.reduce((acc, cur) => acc.replace(new RegExp(cur.regexp, 'g'), cur.value), text);
    return replacedText;
  }

  /**
   * Replace placeholder string in this.messageList.alertMessageOnComplete
   * @param {number} getBlobsCompleteSec
   * @param {number} insertImageCompleteSec
   * @returns {string} The replaced text.
   */
  replaceAlertMessageOnComplete(getBlobsCompleteSec, insertImageCompleteSec) {
    let text = this.messageList.alertMessageOnComplete;
    let placeholderValues = [
      {
        'regexp': '\{\{getBlobsCompleteSec\}\}',
        'value': getBlobsCompleteSec
      },
      {
        'regexp': '\{\{insertImageCompleteSec\}\}',
        'value': insertImageCompleteSec
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }

  /**
   * Replace placeholder string in this.messageList.alertMessageAdd
   * @param {string} key
   * @param {number} value
   * @returns {string} The replaced text.
   */
  replaceAlertMessageAdd(key, value) {
    let text = this.messageList.alertMessageAdd;
    let placeholderValues = [
      {
        'regexp': '\{\{key\}\}',
        'value': key
      },
      {
        'regexp': '\{\{value\}\}',
        'value': value
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }

  /**
   * Replace placeholder string in this.messageList.errorUnknownError
   * @param {string} selectedRangeA1Notation
   * @param {boolean} optionsSelectionVertical
   * @param {number} rangeNumRows
   * @param {number} rangeNumColumns
   * @returns {string} The replaced text.
   */
  replaceErrorUnknownError(selectedRangeA1Notation, optionsSelectionVertical, rangeNumRows, rangeNumColumns) {
    let text = this.messageList.errorUnknownError;
    let placeholderValues = [
      {
        'regexp': '\{\{selectedRangeA1Notation\}\}',
        'value': selectedRangeA1Notation
      },
      {
        'regexp': '\{\{optionsSelectionVertical\}\}',
        'value': optionsSelectionVertical
      },
      {
        'regexp': '\{\{rangeNumRows\}\}',
        'value': rangeNumRows
      },
      {
        'regexp': '\{\{rangeNumColumns\}\}',
        'value': rangeNumColumns
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }

  /**
   * Replace placeholder string in this.messageList.promptCurrentValue
   * @param {string} value
   * @returns {string} The replaced text.
   */
  replacePromptCurrentValue(value) {
    let text = this.messageList.promptCurrentValue;
    let placeholderValues = [
      {
        'regexp': '\{\{value\}\}',
        'value': value
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }

  /**
   * Replace placeholder string in this.messageList.errorImageFileFormat
   * @param {string} fileName
   * @returns {string} The replaced text.
   */
  replaceErrorImageFileFormat(fileName) {
    let text = this.messageList.errorImageFileFormat;
    let placeholderValues = [
      {
        'regexp': '\{\{fileName\}\}',
        'value': fileName
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }

  /**
   * Replace placeholder string in this.messageList.errorImageFileSizeExceedsLimit
   * @param {string} fileName
   * @returns {string} The replaced text.
   */
  replaceErrorImageFileSizeExceedsLimit(fileName) {
    let text = this.messageList.errorImageFileSizeExceedsLimit;
    let placeholderValues = [
      {
        'regexp': '\{\{fileName\}\}',
        'value': fileName
      }
    ];
    text = this.replacePlaceholders_(text, placeholderValues);
    return text;
  }
}
