function setBinding(column = "TestColumn", cell = "TestCell") {
  if (!checkValidColumn(column)) {
    throw ("Invalid column name");
  }
  if (!checkValidCell(cell)) {
    throw ("Invalid cell name");
  }
  var documentProperties = PropertiesService.getDocumentProperties();
  // id = getLatestPropertyID(userProperties) + 1
  id = getRandomInt(10000);
  data = {
    "column": column,
    "cell": cell
  }
  documentProperties.setProperties({ [id]: JSON.stringify(data) })
}
function setFolderProperty(id, url, name) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperties({
    "pdfFolder": JSON.stringify({
      "id": id,
      "url": url,
      "name": name
    })
  })
}
function setProperty(id, value) {
  var documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(id, value);
  d = {};
  d[id]=value;
  return d
}
function setPropertyIfNeeded(currentProperties, id, value) {
  if (!(id in currentProperties)) {setProperty(id, value); return}
  if ((id in currentProperties)&&(currentProperties[id] != value)) {setProperty(id, value); return}
}
function getFolderProperty() {
  var documentProperties = PropertiesService.getDocumentProperties();
  prop = JSON.parse(documentProperties.getProperty("pdfFolder"));
  Logger.log(prop);
  return prop
}

function printAllProperties() {
  var documentProperties = PropertiesService.getDocumentProperties();
  prop = documentProperties.getProperties();
  var ui = SpreadsheetApp.getUi();
  ui.alert(JSON.stringify(prop,null, "\t"));
  return prop
}

function importAllProperties(){
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('Importing properties from Json', 'Insert the text with settings as you find it when using "Print all Settings" in Genera menu. Then, to continue, press OK.\n All your current settings will be deleted.', ui.ButtonSet.OK_CANCEL);

  // Process the user's response.
  if (response.getSelectedButton() == ui.Button.OK) {
    try{
      dictToImport = JSON.parse(response.getResponseText());
      const documentProperties = PropertiesService.getDocumentProperties();
      documentProperties.deleteAllProperties();
      for (id in dictToImport){
        setProperty(id, dictToImport[id]);
      }
    }catch(e){
      ui.alert("Something is wrong with the text you are trying to import \n "+e)
    }
    ui.alert("Import ok.")

  }

}


function readProperties() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperties();
}

function deleteAllProperties() {
  try {
    // Get user properties in the current user.
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to delete all settings?\n Please, restart Genera afterwards.', ui.ButtonSet.YES_NO);
    // Process the user's response.
    if (response == ui.Button.YES) {
      const documentProperties = PropertiesService.getDocumentProperties();
    // Delete all user properties in the current user.
      documentProperties.deleteAllProperties();
    }
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
function deleteSingleProperty(key = 1) {

  try {
    // Get user properties in the current user.
    const documentProperties = PropertiesService.getDocumentProperties();
    // Delete all user properties in the current user.
    documentProperties.deleteProperty(key);
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
function getLatestPropertyID(documentProperties) {
  properties = documentProperties.getProperties();
  var keys = documentProperties.getKeys()
  keys = keys.filter(function (ele) {
    return !isNaN(parseInt(ele));
  })

  var latestID = keys.reduce((a, b) => properties[a] > properties[b] ? a : b, 0);
  return Number(latestID)
}

