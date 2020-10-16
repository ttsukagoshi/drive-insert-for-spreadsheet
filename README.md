# Insert Image into Google Sheets from a private Google Drive Folder (English / [日本語](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/blob/main/README.ja.md))
## Background & About this App/Script
The [built-in `=IMAGE(<url>)` function of Google Sheets](https://support.google.com/docs/answer/3093333) is an easy-to-use and highly available function for inserting images into Google Sheets. There is a limitation, however, that the image it can refer to must be publically accessible. For example, images stored in the user's private Google Drive folder cannot be displayed on the same user's Google Sheets using the function. This can be quite inconvenient when preparing an internal document, containing some sensitive or non-disclosed images and figures, for your team at work.

This app (or the script) is a work-around for this limitation, enabling the user to directly refer to image files in a Drive folder regardless of the folders shared status, public or private, from their Google Sheets file. The details on how to use the script is described on [this sample Google Sheets](https://docs.google.com/spreadsheets/d/1Ck2GgMwbTUZeag5HWeG05ZS_j7IN935nqXfcunPZgC4/edit#gid=0).  
Create your file by copying the sample file from `File` > `Copy`.

## Note
- The language of the sample Google Sheets file is set to `en_US`, i.e., English (USA). Change the language setting of your copied file from `File` > `Spreadsheet settings`.  
For more information on how to change the locale, see [Set a spreadsheet’s location & calculation settings - Google Docs Editors Help](https://support.google.com/docs/answer/58515)
- This Google Apps Script utilizes [`Sheet.insertImage()`](https://developers.google.com/apps-script/reference/spreadsheet/sheet#insertimageblobsource,-column,-row). Although not officially described, there seems to be some limitations to the image files that can be inserted using this method (@tanaikech [reports this in detail](https://gist.github.com/tanaikech/9414d22de2ff30216269ca7be4bce462)):
  - Pixel size: The pixel area, or the product of horizontal and vertical pixel length, of the image must be less than or equal to 2<sup>20</sup> = 1,048,576 pixel<sup>2</sup>
  - File type:`image/jpeg`, `image/png`, and `image/gif` are accepted