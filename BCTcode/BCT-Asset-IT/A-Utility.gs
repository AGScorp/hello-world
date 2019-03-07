function insertNewRowPR() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var lastrow = sheet.getLastRow();
  BCT.insertRows(15);
 var newlastrow = sheet.getLastRow();
  sheet.getRange('M'+(lastrow+1)+':'+'P'+newlastrow).clearContent();
  
}
