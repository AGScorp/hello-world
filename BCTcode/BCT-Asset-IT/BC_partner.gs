function loadPartner() {
  BCT.loadDataSpreadsheetByTemplate();
}

function savePartner(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var eventRane = ss.getActiveRange();
  var eventRow = eventRane.getRow();  
  var eventColumn = eventRane.getColumn();  
  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, eventRow, 1);
  var DBsever = BCT.getDBServer();
  var DBName = "BCT_partner_datacenter";
  
  
      var tableName = 'partnercenter';
      var fieldKey = 'partnerID';
      var partnerID = BCT.valueByFliedName(fields, values, 'partnerID');
      var partnerIDCol = BCT.nameColumnByFliedName(fields, 'partnerID');
      if(partnerID==''){
        var new_partnerID = '';
        var prefigKey = 'PN-';
        var query = "SELECT RIGHT("+fieldKey+",6) as newId FROM "+tableName+" WHERE "+fieldKey+" LIKE '"+prefigKey+"%' ORDER BY newId DESC LIMIT 1";   
        var datas = BCT.loadXMLDatas(DBsever, DBName, query);
        if(datas.length>0){
          var id = datas[0].getChild('newId').getValue();
          new_partnerID = prefigKey+Utilities.formatString("%06d", Number(id)+1);
        }else{
          new_partnerID = prefigKey+Utilities.formatString("%06d", 1);
        }
       sheet.getRange(partnerIDCol+eventRow).setValue(new_partnerID);
      }
        
 BCT.saveDataSpreadsheetByTemplate(true);        
}