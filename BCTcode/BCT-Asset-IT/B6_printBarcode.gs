function printBarcode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  sheet.getRange('A11:F').clearContent();
  var cond = sheet.getRange("B5").getValue();

  BCTCENTER.queryDB("RDS","BCT_Asset_Pkg","Asset", cond,sheet.getRange("B7:F7").getValues(),2,11);  
  ss.toast("เรียกข้อมูลแล้วเสร็จ");
}
