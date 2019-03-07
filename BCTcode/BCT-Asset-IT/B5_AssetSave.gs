//coding by kowit

function loadPR2Asset() {
//  BCT.loadDataSpreadsheetByTemplate(false);  //เรรียกข้อมูลแบบไม่ลบของเก่า
  BCT.loadDataSpreadsheetByTemplate(); 

  
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  var sheet = ss.getActiveSheet();
//  var sheetName = sheet.getSheetName();  
//    var DBfields_condition = BCT.form_Fields(sheet, "condition"); 
//    var DBvalues_condition = BCT.form_Values(sheet, "condition"); 
//    var DBserver = BCT.valueByFliedName(DBfields_condition, DBvalues_condition, 'DBserver'); 
//    var DBname = BCT.valueByFliedName(DBfields_condition, DBvalues_condition, 'DBName'); 
//    var tableName = BCT.valueByFliedName(DBfields_condition, DBvalues_condition, 'tableName'); 
//    var queryValue = BCT.valueByFliedName(DBfields_condition, DBvalues_condition, 'send2Asset'); //ชื่อfiled ในการค้นหา  
//    var query = "select * from "+tableName;
//        query += " WHERE send2Asset ='"+queryValue+"' ";
//
//  var datas = BCT.loadXMLDatas(DBserver, DBname, query);
//  for(var i=0;i<datas.length;i++){
//     var qty = datas[i].getChild("qty").getValue();
//Browser.msgBox(qty);
//  }  
//
//
//  
//  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');
//  var rowstart = rowStartValue+1;  
//  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
//  var fieldReport =  BCT.getFields(sheet, rowFields, 1, 0);  
//  BCT.autoInsert(sheet,fieldReport, datas, rowstart, 1);  
//  
  
  
}

function saveAssetAGS(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var eventRange = SpreadsheetApp.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows= eventRange.getNumRows();  
  var rowActive = sheet.getActiveCell().getRow();
  
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');
  var rowstart = rowStartValue;
  
  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);
  var values = BCT.getValue(sheet, rowActive, 1);
  var productID = BCT.nameColumnByFliedName(fields, 'product_ID');
  var fullAsset_ID = BCT.numberColumnByFliedName(fields, 'id_assset_full');
  var fullID_value = BCT.valueByFliedName(fields, values, 'id_assset_full');

//  Browser.msgBox(fullID_value);
//  return;
  if(eventRow >= rowstart)
  {
    var DBsever = "RDS";
    var DBName = "BCT_Asset_Pkg";
    var tableName = "Asset";  
    var idName = "id_assset_full";
    var lastColumnA = BCTCENTER.replaceAll(sheet.getRange(1, sheet.getLastColumn()).getA1Notation(), "1");  
    if(fullID_value =="")
    {      
    ss.toast("กรุณารอ....","กำลังบันทึก",30);      
    var idAsset = sheet.getRange(productID+eventRow).getValues();
    var prefigKey = idAsset + "-"+Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "yy");
    var ID = BCTCENTER.syncDB(DBsever,DBName,tableName,sheet,"A",lastColumnA,rowStartValue-4,rowStartValue-3,eventRow,numRows,fullAsset_ID,fullAsset_ID,idName,prefigKey,4);

    ss.toast("บันทึกทรัพย์สินแล้ว");
    }
    else { 
      ss.toast("updateข้อมูลทรัพย์สิน");  
      var ID = BCTCENTER.syncDB(DBsever,DBName,tableName,sheet,"A",lastColumnA,rowStartValue-4,rowStartValue-3,eventRow,numRows,fullAsset_ID,fullAsset_ID,idName);  
//               BCTCENTER.syncDB(DBsever, DBName, tableName, sheet,startColumn, lastColumn, startDBRow, lastDBRow, eventRow, numRows, key_column, key_column_in_range, idName)
//      BCT.saveDataSpreadsheetByTemplate(true); 
    }
  } else{ Browser.msgBox("กรุณา เลือกบรรทัดที่จะบันทึกให้ถูกต้อง"); }
  
  //update table ข้อมูลจัดซื้อ AGS ว่ารายการไหน บันทึกทรัยพ์สินแล้ว
  var DBsever = "RDS";
  var DBName = "BCT_Asset_Pkg";  
  var fullAsset_ID_Colname = BCT.nameColumnByFliedName(fields, 'id_assset_full');
  var pr_no_Colname = BCT.nameColumnByFliedName(fields, 'pr_no');
  for(var i=0;i<numRows;i++)
  {
    var fullID_value = sheet.getRange(fullAsset_ID_Colname+(eventRow+i)).getValue();
    var prno_value = sheet.getRange(pr_no_Colname+(eventRow+i)).getValue();
    if(fullID_value != "")
    {
      var query = "UPDATE purchase_IT set send2Asset='"+fullID_value+"' where pr_no ="+prno_value;
      BCTCENTER.loadXMLQueryInsertUpdate(DBsever, DBName, query);
    }
  }
  
  
}