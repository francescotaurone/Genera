function setProperties(column = "TestColumn", cell = "TestCell") {
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
  ui.alert(JSON.stringify(prop));
  return prop
}
function readProperties() {
  var documentProperties = PropertiesService.getDocumentProperties();
  return documentProperties.getProperties();
}

function deleteAllProperties() {
  try {
    // Get user properties in the current user.
    var ui = SpreadsheetApp.getUi();
    var response = ui.alert('Are you sure you want to delete all settings?', ui.ButtonSet.YES_NO);
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

