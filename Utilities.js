function searchFilesByName(pdfName) {
  //Please enter your search term in the place of Letter
  var searchFor = "title = '" + pdfName + "'"
  var names = [];
  var fileIds = [];
  var files = ReDriveApp.searchFiles(searchFor);
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();// To get FileId of the file
    fileIds.push(fileId);
    var name = file.getName();
    names.push(name);

  }
  return fileIds
}

function removeFilesByName(pdfName) {
  filesToBeRemoved = searchFilesByName(pdfName);
  filesToBeRemoved.forEach(function (fileId) {
    var file = ReDriveApp.getFileById(fileId);
    file.setTrashed(true);
  })
}
function removeFilesById(pdfId) {
  pdfId.forEach(function (id) {
    var file = ReDriveApp.getFileById(id);
    file.setTrashed(true);
  })
}


/**
 * Returns a Google Drive folder in the same location 
 * in Drive where the spreadsheet is located. First, it checks if the folder
 * already exists and returns that folder. If the folder doesn't already
 * exist, the script creates a new one. The folder's name is set by the
 * "OUTPUT_FOLDER_NAME" variable from the Code.gs file.
 *
 * @param {string} folderName - Name of the Drive folder. 
 * @return {object} Google Drive Folder
 */
function getFolderByName_(folderName) {

  // Gets the Drive Folder of where the current spreadsheet is located.
  const ssId = SpreadsheetApp.getActiveSpreadsheet().getId();
  const parentFolder = ReDriveApp.getFileById(ssId).getParents().next();

  // Iterates the subfolders to check if the PDF folder already exists.
  const subFolders = parentFolder.getFolders();
  while (subFolders.hasNext()) {
    let folder = subFolders.next();

    // Returns the existing folder if found.
    if (folder.getName() === folderName) {
      return folder;
    }
  }
  // Creates a new folder if one does not already exist.
  return parentFolder.createFolder(folderName)
    .setDescription(`Created by ${APP_TITLE} application to store PDF output files`);
}


/**
 * Test function to run getFolderByName_.
 * @prints a Google Drive FolderId.
 */
function test_getFolderByName() {

  // Gets the PDF folder in Drive.
  const folder = getFolderByName_(OUTPUT_FOLDER_NAME);

  console.log(`Name: ${folder.getName()}\rID: ${folder.getId()}\rDescription: ${folder.getDescription()}`)
  // To automatically delete test folder, uncomment the following code:
  // folder.setTrashed(true);
}

function columnToLetter(column = 696) {
  var temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

function letterToColumn(letter = "E") {
  var column = 0, length = letter.length;
  for (var i = 0; i < length; i++) {
    column += (letter.charCodeAt(i) - 64) * Math.pow(26, length - i - 1);
  }
  return column;
}

function getRandomInt(max) {
  return Math.floor(Math.random() * max);
}

const validateEmail = (email) => {
  return String(email)
    .toLowerCase()
    .match(
      /^(([^<>()[\]\\.,;:\s@"]+(\.[^<>()[\]\\.,;:\s@"]+)*)|.(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/
    );
};

function urlFetchWihtoutError(url, requestData) {
  const NB_RETRY = 3;
  var nbSecPause = 1.5;
  var nbErr = 0;
  var error;
  while (nbErr < NB_RETRY) {
    try {
      if (nbErr > 0) SpreadsheetApp.getActiveSpreadsheet().toast("Retrying PDF generation: " + nbErr);
      var res = UrlFetchApp.fetch(url, requestData);
      return res;
    }
    catch (error) {
      nbErr++;
      Utilities.sleep(nbSecPause * 1000)
      nbSecPause += 0.5;
    }
  }
  throw "Too many retries:" + error;
}

function checkValidColumn(column) {
  const columnRegex = new RegExp("[a-zA-Z]+");
  if (!columnRegex.test(column)) {
    return false;
  }
  return true;
}
function checkValidRow(row) {
  if (isNaN(parseInt(row))) {
    return false;
  }
  return true;
}

function checkValidCell(cell){
  const cellRegex = new RegExp("[a-zA-Z]+[0-9]+");
  if (!cellRegex.test(cell)) {
    return false;
  }
  return true;
}

