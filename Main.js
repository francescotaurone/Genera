/** * @OnlyCurrentDoc */ 

// Application constants
const APP_TITLE = 'GENERA';

function generateResultingPDF(rowToProcess = 1) {
    //{pdfname=A, pdfsheet=pdf, 1={"column":"B","cell":"B3"}, 2={"column":"C","cell":"B4"}, pdflastrow=6, pdfFolder={"id":"1a_jQNoce82PR3vNZaPJYsHlBt3SFQWuI","url":"https://drive.google.com/drive/folders/1a_jQNoce82PR3vNZaPJYsHlBt3SFQWuI","name":"outputFolderTest"}, datasheet=Foglio2, pdflastcol=4}

    properties = readProperties();

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
    var emailAddress = rowData[mailColumnIndex];

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
    for (binding of bindings) {
        pdfSheet.getRange(binding["cell"]).setValue(rowData[letterToColumn(binding["column"]) - 1]);
    }

    SpreadsheetApp.flush();
    Utilities.sleep(500); // Using to offset any potential latency in creating .pdf

    const pdfUrl = createPDF(ss.getId(), pdfSheet, pdfName, properties["pdflastrow"], properties["pdflastcol"]);

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
        body = convertSheetToHtml(properties["pdflastrow"], properties["pdflastcol"]) 
        subject = "Ricevuta"
        senderName = "Genera"
        createDraft(emailAddress, subject, body, senderName)
    }
    
    return JSON.stringify({
        "pdfName": pdfName,
        "pdfUrl": pdfUrl//pdf.getUrl()
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
    const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
        "?format=pdf&" +
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
        "attachment=true&" +
        "gid=" + sheet.getSheetId() + '&' +
        "r1=" + fr + "&c1=" + fc + "&r2=" + lr + "&c2=" + lc;

    /*const params = { method: "GET", headers: { "authorization": "Bearer " + ScriptApp.getOAuthToken() } , 'muteHttpExceptions' : true};
    var response = UrlFetchApp.fetch(url, params);
    if(response.getResponseCode() != 200)
    {
      Logger.log("URL: "+url );
      Logger.log("Params:" + JSON.stringify(params) );
      Logger.log(pdfName + "\nResponse Code: " + response.getResponseCode() + " \nContent Text:\n" + response.getContentText());
      return false
    }
    else{
      const blob = response.getBlob().setName(pdfName + '.pdf');
      // Gets the folder in Drive where the PDFs are stored.
      //console.log(pdfFolderId);
      //const folder = ReDriveApp.getFolderById(pdfFolderId);
  
      //const pdfFile = ReDriveApp.createFile(blob,undefined,undefined,pdfFolderId);
      ////const pdfFile = folder.createFile(blob);
      //return pdfFile;
      return "data:application/pdf;base64," + Base64.encode(blob.getBytes());
      
    }
    */
    return url
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
    const url = `https://docs.google.com/spreadsheets/export?exportFormat=zip&id=1M9s8noMuleg2Q4qTul5-hLQ_rdcCb3l2mewsOtns7-E`;
    const blob = UrlFetchApp.fetch(url, {
        headers: { authorization: "Bearer " + ScriptApp.getOAuthToken() },
    }).getBlob();
    const blobs = Utilities.unzip(blob);
    return blobs
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

function createDraft(mailAddress, subject, body, name) {

    let recipient = mailAddress;
    var drafts = GmailApp.getDrafts();
    if (drafts.length === 0){
        
    }
    for (var i = 0; i < drafts.length; i++) {
        Logger.log(drafts[i].getId());
    }
    /*
    GmailApp.createDraft(recipient, subject, "Ricevuta. Cordialmente", {
        name: name,
        htmlBody: body
    });
    */ 
    
  }