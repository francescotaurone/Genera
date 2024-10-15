/** * @OnlyCurrentDoc */ 
//"https://www.googleapis.com/auth/drive.file",
//"https://www.googleapis.com/auth/script.locale",
// Application constants
const APP_TITLE = 'GENERA';


function onFormSubmitFunction(e) {
  Logger.log(JSON.stringify(e))
  console.log("Trigger in execution");
  Utilities.sleep(4000); //Per essere sicuri che la risposta appena entrata vada nel tab
  generateResultingPDF(rowToProcess = 2) //the last arriving row should be row 2. It is vilnerable from many forms coming in at the same time.
}

function generateResultingPDF(rowToProcess = 1) {
    //{pdfname=A, pdfsheet=pdf, 1={"column":"B","cell":"B3"}, 2={"column":"C","cell":"B4"}, pdflastrow=6, pdfFolder={"id":"1a_jQNoce82PR3vNZaPJYsHlBt3SFQWuI","url":"https://drive.google.com/drive/folders/1a_jQNoce82PR3vNZaPJYsHlBt3SFQWuI","name":"outputFolderTest"}, datasheet=Foglio2, pdflastcol=4}

    
    properties = readProperties();
    Logger.log("Reading properties "+ JSON.stringify(properties))

    // Bindings
    var bindings = [];
    for (key in properties) {
        if (!isNaN(parseInt(key))) {
            bindings.push(JSON.parse(properties[key]));
        }
    }

    // Dati
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var dataSheet = ss.getSheetByName(properties["datasheet"]);
    var data = dataSheet.getDataRange().getValues();

    // PDF Sheet
    const pdfSheet = ss.getSheetByName(properties["pdfsheet"]);
    //const pdfFolder = JSON.parse(properties["pdfFolder"]);
    //const pdfFolderId = pdfFolder["id"];
    const pdfColumnIndex = letterToColumn(properties["pdfname"]) - 1;
    const mailColumnIndex = letterToColumn(properties["emailcolumn"]) - 1;
    var pdfs = [];
    var rowData = data[rowToProcess - 1];
    var pdfName = rowData[pdfColumnIndex];
    

    /*
    // Verifica se il file è già presente
    const filesForSearch = ReDriveApp.getFolderById(pdfFolderId).getFiles();
    var searchFor ="title = '" + pdfName + "'"
    var fileIds=[];
    while (filesForSearch.hasNext()) {
      var file = filesForSearch.next();
      var fileId = file.getId();// To get FileId of the file
      fileIds.push(fileId);
    }
    if (fileIds.length > 0 ) {
       removeFilesById(fileIds);
    } 
    */
    Logger.log("Clean fields in template");
    // pulisciCampi
    cells = [];
    for (binding of bindings) {
        cells.push(binding["cell"]);
    }

    const rngClear = pdfSheet.getRangeList(cells).getRanges()
    rngClear.forEach(function (cell) {
        cell.clearContent();
    });

    // Set values in the template.
    Logger.log("Set values in template");
    for (binding of bindings) {
        pdfSheet.getRange(binding["cell"]).setValue(rowData[letterToColumn(binding["column"]) - 1]);
    }

    SpreadsheetApp.flush();
    //Utilities.sleep(500); // Using to offset any potential latency in creating .pdf

    const pdf = createPDF(ss.getId(), pdfSheet, pdfName, properties["pdflastrow"], properties["pdflastcol"]);

    /*
    rngClear.forEach(function (cell) {
        cell.clearContent();
    });
    
    if (pdf === false){
      return false
    }
    else{
      return JSON.stringify({
        "pdfName": pdfName, 
        "pdfUrl": pdf//pdf.getUrl()
      })
    }
      */

    // Mail
    if (properties["emailchecked"] === "true"){
        
        //body = convertSheetToHtml(properties["pdflastrow"], properties["pdflastcol"]) ;
        body = properties["emailbody"];
        subject = properties["emailsubject"]
        senderName = properties["emailsendername"]
        attachment = pdf["pdf"];
        emailAddress = rowData[mailColumnIndex];
        Logger.log("Sending email to "+ emailAddress);
        sendEmail(emailAddress, subject, body, senderName, attachment);
    }
    
    return JSON.stringify({
        "pdfName": pdfName,
        "pdfUrl": pdf["url"]//pdf.getUrl()
    })
}
function processDate(dateString = "28/09/2012") {
    year = +dateString.substring(6)
    month = +dateString.substring(3, 5)
    day = +dateString.substring(0, 2)

    pubdate = new Date(year, month - 1, day)
    newDate = Utilities.formatDate(pubdate, 'Europe/Rome', 'dd/MM/yyyy')
    return newDate
}

function ottieniDataDaInfoCronologiche(infoCronologiche, perNomeFile = false) {
    var year
    var month
    var day
    var date
    if (typeof (infoCronologiche) == "string") {
        stringDate = infoCronologiche;
        year = +stringDate.substring(6, 10)
        month = +stringDate.substring(3, 5)
        day = +stringDate.substring(0, 2)
        dateForProcessing = new Date(year, month - 1, day)
    }
    else {
        dateForProcessing = new Date(infoCronologiche)

    }
    if (perNomeFile === true) {
        date = Utilities.formatDate(dateForProcessing, 'Europe/Rome', 'yyyy_MM_dd')
    }
    else {
        date = Utilities.formatDate(dateForProcessing, 'Europe/Rome', 'dd/MM/yyyy')
    }
    return date
}

function generaNomePDFRicevutaFinale(form, nome_e_cognome_figlio) {
    const date = ottieniDataDaInfoCronologiche(form[0], perNomeFile = true);
    //return `IscrizioneER2024_RICEVUTA_${date}_${nome_e_cognome_figlio.replace(/\W+/g, '_').toLowerCase()}`
    return `ER2024_RICEVUTA_${nome_e_cognome_figlio.replace(/\W+/g, '_').toLowerCase()}`
}

function createPDF(ssId, sheet, pdfName, lastRow, lastCol) {
    const fr = 0, fc = 0, lc = lastCol, lr = lastRow;

    const url = "https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=" + ssId +"&exportFormat=pdf&" +
    //const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +"?format=pdf&" +
        "size=7&" +
        "fzr=true&" +
        "portrait=true&" +
        "horizontal_alignment=CENTER&" +
        //"fitw=true&" +
        "scale=4&" +
        "gridlines=false&" +
        "printtitle=false&" +
        "top_margin=0.5&" +
        "bottom_margin=0.25&" +
        "left_margin=0.5&" +
        "right_margin=0.5&" +
        "sheetnames=false&" +
        "pagenum=UNDEFINED&" +
        //"attachment=true&" +
        "gid=" + sheet.getSheetId() + '&' +
        "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;



    var requestData = {
      "oAuthServiceName": "spreadsheets",
      "oAuthUseToken": "always",
    };
    var response = UrlFetchApp.fetch(url, requestData)
    var pdfBlob = response.getBlob().setName(pdfName + '.pdf');
    pdf = pdfBlob.getAs("application/pdf");
    
    return {"pdf":pdf, "url":url}
    
    /*
    const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } , 'muteHttpExceptions' : false};
    var response = UrlFetchApp.fetch(url, params);
    if(response.getResponseCode() != 200)
    {
      Logger.log("URL: "+url );
      Logger.log("Params:" + JSON.stringify(params) );
      Logger.log(pdfName + "\nResponse Code: " + response.getResponseCode() + " \nContent Text:\n" + response.getContentText());
      return false
    }
    
    const blob = response.getBlob().setName(pdfName + '.pdf');
    // Gets the folder in Drive where the PDFs are stored.
    //console.log(pdfFolderId);
    //const folder = ReDriveApp.getFolderById(pdfFolderId);

    //const pdfFile = ReDriveApp.createFile(blob,undefined,undefined,pdfFolderId);
    ////const pdfFile = folder.createFile(blob);
    //return pdfFile;
    return {"pdf":pdf, "url":url}
    */
}

function getNamesOfSheets() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheets = ss.getSheets();
    var names = [];
    for (sheet of sheets) {
        names.push(sheet.getName());
    }
    return names
}

function testPDF() {
    SpreadsheetApp.flush();
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=zip&id=1dWmWrkvyzb5LO_f3Zucmrtf80xdkzb10ALvEHfuq-R0`;
    const blob = UrlFetchApp.fetch(url, {
        headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
    }).getBlob();
    const blobs = Utilities.unzip(blob);
    blobs.forEach( (blob) =>{
        Logger.log(blob.getString());
    });
    return blobs
}
function testPDF2(printSheet,docName) {
  //
  // returns a PDF of the given sheet
  //
  //

  var requestData = {
    "oAuthServiceName": "spreadsheets",
    "oAuthUseToken": "always",
  };
 
  //
  printSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("pdf");
  var docKey = printSheet.getParent().getId();
  var printSheetIndex = printSheet.getSheetId();
  //
  // Trim the sheet to length
  //
  var lastRow = printSheet.getLastRow();
  var maxRow = printSheet.getMaxRows();
  var maxCol = printSheet.getMaxColumns();
  var lastCol = printSheet.getLastColumn();
 
  if (maxCol > lastCol)
    printSheet.deleteColumns(lastCol + 1, maxCol - lastCol);
  if (maxRow > lastRow)
    printSheet.deleteRows(lastRow + 1, maxRow - lastRow);
 
  url = "https://spreadsheets.google.com/feeds/download/spreadsheets/Export?key="
               +docKey
               +"&exportFormat=pdf&gid="
               +printSheetIndex
               +"&gridlines=true&printtitle=false&size=A4&sheetnames=false&fzr=true"
               +"&fitw=true";
  var pdfBlob = UrlFetchApp.fetch(url, requestData).getBlob().setName(docName);
  bytes = pdfBlob.getAs("application/pdf").getBytes();
  Logger.log(bytes[0]);
  Logger.log(url);
  SpreadsheetApp.getUi().alert("URL: "+ url + " Byte0: "+bytes[0]);
  return bytes
}

function convertSheetToHtml(lastRow = 5, lastCol = 6) {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var pdfSheet = ss.getSheetByName("pdf");
    var pdfRange = pdfSheet.getRange(1, 1, lastRow, lastCol)
    var htmlTable = convertRange2html(pdfRange); 
    images = pdfSheet.getImages();
    images.forEach((image) => {
        Logger.log("getUrl: "+image.getUrl());
        Logger.log("getAnchorCell: "+JSON.stringify(image.getAnchorCell()));
        Logger.log("getAnchorCellXOffset: "+JSON.stringify(image.getAnchorCellXOffset()));
        Logger.log("getAnchorCellYOffset: "+JSON.stringify(image.getAnchorCellYOffset()));
        Logger.log("getAltTextDescription(): "+JSON.stringify(image.getAltTextDescription()));
    })
    // There is no blob support as of now, see https://issuetracker.google.com/119800855
    
    return htmlTable
}

function sendEmail(recipient, subject, body, name, attachments) {

    GmailApp.sendEmail(recipient, subject, body, {
        attachments: attachments,
        name:name
    }); 
    
  }

/**
 * Dummy function for API authorization only.
 */
function forAuth_() {
  DriveApp.getFileById("Just for authorization"); // https://code.google.com/p/google-apps-script-issues/issues/detail?id=3579#c36
}