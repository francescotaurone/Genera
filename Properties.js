function setProperties(column = "TestColumn", cell = "TestCell"){
  var userProperties = PropertiesService.getUserProperties();
  // id = getLatestPropertyID(userProperties) + 1
  id = getRandomInt(10000);
  data = {
    "column" : column,
    "cell" : cell
  }
  userProperties.setProperties({[id]: JSON.stringify(data)})  
}
function setFolderProperty(id, url, name){
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperties({"pdfFolder": JSON.stringify({
    "id": id,
    "url": url,
    "name":name
  })})  
}
function setProperty(id, value){
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(id, value);
}

function getFolderProperty(){
  var userProperties = PropertiesService.getUserProperties();
  prop = JSON.parse(userProperties.getProperty("pdfFolder"));
  Logger.log(prop);
  return prop
}

function readProperties(){
  var userProperties = PropertiesService.getUserProperties();
  Logger.log(userProperties.getProperties());
  return userProperties.getProperties()
}

function deleteAllProperties(){
  try {
    // Get user properties in the current user.
    const userProperties = PropertiesService.getUserProperties();
    // Delete all user properties in the current user.
    userProperties.deleteAllProperties();
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
function deleteSingleProperty(key=1){
  
  try {
    // Get user properties in the current user.
    const userProperties = PropertiesService.getUserProperties();
    // Delete all user properties in the current user.
    userProperties.deleteProperty(key);
  } catch (err) {
    // TODO (developer) - Handle exception
    console.log('Failed with error %s', err.message);
  }
}
function getLatestPropertyID(userProperties){
  properties = userProperties.getProperties();
  var keys = userProperties.getKeys()
  keys = keys.filter(function(ele){
   return !isNaN(parseInt(ele));
  })

  var latestID = keys.reduce((a, b) => properties[a] > properties[b] ? a : b, 0);
  return Number(latestID)
}

