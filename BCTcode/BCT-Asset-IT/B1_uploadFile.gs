function openDialog() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var colActive = sheet.getActiveCell().getColumn();  
  
  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');  
  var fields = BCT.getFields(sheet, rowFields, 1, 0);  
  var linkPRfileCol = BCT.numberColumnByFliedName(fields, 'linkPRfile');
  
  if(rowActive > rowStartValue && colActive == linkPRfileCol)
  {
    var html = HtmlService.createHtmlOutputFromFile('B1_formUpload');
    html.setSandboxMode(HtmlService.SandboxMode.IFRAME);
    SpreadsheetApp.getUi().showModalDialog(html, 'Upload File');    
  }else { SpreadsheetApp.getUi().alert( "กรุณาคลิกช่องให้ถูกต้อง");}
 
}

/* This function will process the submitted form */
function uploadFiles(form) {  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');  
  var fields = BCT.getFields(sheet, rowFields, 1, 0);  
  var values = BCT.getValue(sheet, rowActive, 1);
  var pr_no = BCT.valueByFliedName(fields, values, 'pr_no');
  var datePR = BCT.valueByFliedName(fields, values, 'po_create_date'); 
  var linkPRfileCol_A = BCT.nameColumnByFliedName(fields, 'linkPRfile');
  
  var d1 = new Date();    
  var now = new Date(d1.getTime());
  var email = Session.getActiveUser().getEmail();

  SpreadsheetApp.getActiveSpreadsheet().toast('กำลังอัพโหลดไฟล์','Script Running',30);
  Logger.log(form);
  try {    
    /* Name of the Drive folder where the files should be saved */
    var folder = DriveApp.getFolderById('0BwxmTPgUy0_XLW9WZHB3cElyOGs'); //folder AGS PurchaseFiles
    
    /* Get the file uploaded though the form as a blob */
    var blob = form.myFile;    
    var file = folder.createFile(blob);    
    
    /* Set the file description as the name of the uploader */
//    var no_name = checkStamp.getValue().split("_")[3];
//    var car_number = sheet.getRange(eventRange.getRow(), AGS.NumberColumn('N')).getValue();
//    
    file.setName("PR-"+pr_no+"-"+datePR);
    file.setDescription("Uploaded by " + email);
    
    // Set Url and Timestampe to Spreadsheet 
    var values = [];
//    var data = [file.getUrl(), now, email];
    var data = [file.getUrl()];
    values.push(data);
    sheet.getRange(linkPRfileCol_A+rowActive).setValue(values);
//    eventRange.offset(0,-1,1,2).setValues(values); 
    
    /* Return the download URL of the file once its on Google Drive */
    SpreadsheetApp.getActiveSpreadsheet().toast('File uploaded successfully','Script Running',3);
    SpreadsheetApp.getUi().alert( "File uploaded successfully " + file.getUrl())
    return "File uploaded successfully " + file.getUrl();
  } catch (error) {    
    SpreadsheetApp.getActiveSpreadsheet().toast('Error : '+error.toString(),'Script Running',3);
    /* If there's an error, show the error message */
    return error.toString();
  }
}