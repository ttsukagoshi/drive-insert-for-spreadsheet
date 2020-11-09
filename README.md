# Insert Image into Google Sheets from a private Google Drive Folder (English / [日本語](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/blob/main/README.ja.md))
[![Language grade: JavaScript](https://img.shields.io/lgtm/grade/javascript/g/ttsukagoshi/spreadsheet-bulk-import-images.svg?logo=lgtm&logoWidth=18)](https://lgtm.com/projects/g/ttsukagoshi/spreadsheet-bulk-import-images/context:javascript)
## Background & About this App/Script
The [built-in `IMAGE(<url>)` function of Google Sheets](https://support.google.com/docs/answer/3093333) is an easy-to-use and highly available function for inserting images into Google Sheets. There is a limitation, however, that the image it can refer to must be publicly accessible. For example, images stored in the user's private Google Drive folder cannot be displayed on the same user's Google Sheets using the function. This can be quite inconvenient when preparing an internal document, containing some sensitive or non-disclosed images and figures, for your team at work.

This app (or the script) is a work-around for this limitation, enabling the user to directly refer to image files in a Drive folder regardless of the folders shared status, public or private, from their Google Sheets file.  
1. The app version is currently available as an Editor Add-on in the Google Workspace (G Suite) Marketplace, limited to members of a specific organization.  
The [`src` folder of the `main` branch](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/tree/main/src) describes the scripts for this app version.
2. Alternatively, you can use this feature as a spreadsheet-bound script.  
The details on how to use the script is described on [this sample Google Sheets](https://docs.google.com/spreadsheets/d/1Ck2GgMwbTUZeag5HWeG05ZS_j7IN935nqXfcunPZgC4/edit#gid=0). Create your file by copying the sample file from `File` > `Copy`, and you can start using the functions immediately. The source code is identical to that of the [branch `main-script-bound`](https://github.com/ttsukagoshi/spreadsheet-bulk-import-images/tree/main-script-bound)

### Note
- The language of the sample Google Sheets file is set to `en_US`, i.e., English (USA). Change the language setting of your copied file from `File` > `Spreadsheet settings`.  
For more information on how to change the locale, see [Set a spreadsheet’s location & calculation settings - Google Docs Editors Help](https://support.google.com/docs/answer/58515)
- This Google Apps Script utilizes [`Sheet.insertImage()`](https://developers.google.com/apps-script/reference/spreadsheet/sheet#insertimageblobsource,-column,-row). Although not officially described, there seems to be some limitations to the image files that can be inserted using this method ([@tanaikech](https://github.com/tanaikech) reports [this in detail](https://gist.github.com/tanaikech/9414d22de2ff30216269ca7be4bce462)):
  - Pixel size: The pixel area, or the product of horizontal and vertical pixel length, of the image must be less than or equal to 2<sup>20</sup> = 1,048,576 pixel<sup>2</sup>
  - File type:`image/jpeg`, `image/png`, and `image/gif`, i.e., file extensions `jpg`, `png`, and `gif`, are accepted

## How to Use (Add-on ver.)
Let's assume that you have a Google Drive folder containing three image files, `Cat.jpg`, `Dog.jpg`, and `Fish.jpg`.  
You want to insert these images into a Google Sheets file that you are currently working on:  
![Empty Cells](/src/images/readme/01_empty-cells.png)

### 1. Initial Settings
After installing the add-on, you will have to setup necessary parameters before going on to the actual image inserting. This setup needs to be done whenever there is a change in the source of the image, i.e., the Google Drive folder, or if you are working on a new file. The required parameters are  
  - `folderId`: Google Driver folder ID that the images are save in. It's the `*****` part of the folder URL `https://drive.google.com/drive/folders/*****`. To designate the root folder under `My Drive`, enter `root`.
  - `fileExt`: File extension of the image files. This defaults to `jpg`. Note that the period before the file extension is NOT required.
  - `selectionVertical`: Designate the direction of selecting the cells in the spreadsheet. Enter `true` if your selection is vertical, `false` if horizontal.
  - `insertPosNext`: Relative position in which to insert the images based on the position of the selected cells. Enter `true` if you want the images to come after, i.e., to the column right of or the row under, the selected cells.  
The setup can be done from the add-on menu `Insert Image from Drive` > `Setup`  
![Initial setup](/src/images/readme/02_setup.png)

### 2. Inserting the Images
Select the cells that represent the image file name, and from the menu, select `Insert Image from Drive` > `Insert Image` to insert the cooresponding images to the cells next to the selected cells.  
![Insert images](/src/images/readme/03_insert-image.png)

### About Parameters `selectionVertical` and `insertPosNext` and Insert Positions of the Images
| selectionVertical | insertPosNext | Insert Positions |
| --- | --- | --- |
| `true` | `true` | File names are aligned **vertically** on the spreadsheet,<br>and the images should be inserted in the column to the **right** of the file names.<br>![Inserted images: selectionVertical = true, insertPosNext = true](/src/images/readme/04_images-inserted-tt.png) |
| `true` | `false` | File names are aligned **vertically** on the spreadsheet,<br>and the images should be inserted in the column to the **left** of the file names.<br>![Inserted images: selectionVertical = true, insertPosNext = false](/src/images/readme/05_images-inserted-tf.png) |
| `false` | `true` | File names are aligned **horizontally** on the spreadsheet,<br>and the images should be inserted in the row **below** the file names.<br>![Inserted images: selectionVertical = false, insertPosNext = true](/src/images/readme/06_images-inserted-ft.png) |
| `false` | `false` | File names are aligned **horizontally** on the spreadsheet,<br>and the images should be inserted in the row **above** the file names.<br>![Inserted images: selectionVertical = false, insertPosNext = false](/src/images/readme/07_images-inserted-ff.png) |

## Your Privacy
This app/script neither collects nor retains any personal information except for the sole purpose of inserting the images designated by the user into the open spreadsheet. The settings, i.e., the Google Drive folder ID and other parameters that can be modified by `Insert Image from Drive` > `Setup`, are save in the [document properties and script properties](https://developers.google.com/apps-script/guides/properties#comparison_of_property_stores) for the add-on and spreadsheet-bound script versions, respectively, and are shared between the users who have either ownership or editor authorization of the file. No log is recorded for individual script executions, and all data pertaining to the user is not visible to the developer.

### Purpose of Authorization Scopes
| Scope/Meaning | Usage in this app/script |
| --- | --- |
| `https://www.googleapis.com/auth/spreadsheets`<br>Allows read/write access to the user's sheets and their properties.<br>See more on https://developers.google.com/sheets/api/guides/authorizing | <ul><li>Read the contents of the selected cells as strings and use them as file names to search for relevant image files in Google Drive.</li><li>Insert images into the spreadsheet.</li></ul> |
| `https://www.googleapis.com/auth/drive.readonly`<br>Allows read-only access to file metadata and file content.<br>See more on https://developers.google.com/drive/api/v2/about-auth | Search for image files by its file name and get the content as [blob](https://developers.google.com/apps-script/reference/base/blob) to insert in spreadsheet|

## Attribution
- The icon of this Google Sheets add-on is made by [Freepik](https://www.flaticon.com/authors/freepik) from [www.flaticon.com](https://www.flaticon.com/)
- Photos used in the above screenshot and the sample spreadsheet is from Pexels:
  - Cat: Photo by [EVG Culture](https://www.pexels.com/@evgphotos?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/selective-focus-photography-of-orange-tabby-cat-1170986/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)
  - Dog: Photo by [Chevanon Photography](https://www.pexels.com/@chevanon?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/two-yellow-labrador-retriever-puppies-1108099/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)
  - Fish: Photo by [Skitterphoto](https://www.pexels.com/@skitterphoto?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels) from [Pexels](https://www.pexels.com/photo/orange-and-white-fish-886210/?utm_content=attributionCopyText&utm_medium=referral&utm_source=pexels)