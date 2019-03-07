function queryAsset_repair() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var lastrow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var rowMaster =  22;
  var rowRangeMaster1 = sheet.getRange(rowMaster,4,1,9);
  var rowRangeMaster2 = sheet.getRange('O22');
  
  BCT.loadDataSpreadsheetByTemplate(false);
  rowRangeMaster1.copyFormatToRange(sheet, 4, 9, lastrow+1, lastrow+1)
  rowRangeMaster2.copyTo(sheet.getRange(lastrow+1,15));
}

function saveRepair(){
  BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange  
}

function sendrepair2pay(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var filelink = ss.getUrl();
  var sheetID = sheet.getSheetId();

  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, rowActive, 1);

  var repair_no = BCT.valueByFliedName(fields, values, 'running');
  var assetID = BCT.valueByFliedName(fields, values, 'id_assset_full');
  var assetname = BCT.valueByFliedName(fields, values, 'name_asset');
  var assetType = BCT.valueByFliedName(fields, values, 'Type');
  var repair_detail = BCT.valueByFliedName(fields, values, 'repair_detail');
  var repair_amount = BCT.valueByFliedName(fields, values, 'repair_amount');

  
  var subject = 'แจ้งFAM : ทำจ่ายค่าซ่อม ';
  var message = '\n เลขที่ใบซ่อม : '+repair_no
                +'\n รหัสทรัพย์สิน : '+assetID+' ,'+assetType+' '+assetname
                +'\n รายการซ่อม : '+repair_detail
                +'\n ค่าซ่อม : '+repair_amount;
                
//  BCTCENTER.LineNotify("8IGwhopbpj22QR733Sf54OaC5cFw7pN5VvJS6Llc4ND", subject, message);//line kowit
  BCTCENTER.LineNotify("xujgaQsWgJxH61NBQyBAOsKjapjh6jrMHGD0JjK41h0", subject, message);//line FAM+AGS
  
  BCT.saveDataSpreadsheetByTemplate(true);   
}