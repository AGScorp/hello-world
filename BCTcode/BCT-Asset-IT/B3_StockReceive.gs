function AGSreceive() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');
  
  if(eventRow>=rowStartValue){  
    var rowFields = BCT.form_getRowFieldsByKey(sheet, 'update');
    var fields = BCT.getFields(sheet, rowFields, 1, 0);     
    
    var tableName = 'purchase_IT';
    var queryInsert ='';
    var pr_no = BCT.numberColumnByFliedName(fields, 'pr_no');
    var statusCol = BCT.nameColumnByFliedName(fields, 'status');
    for(var i = 0;i<numRows;i++)
    {
    sheet.getRange(statusCol+(eventRow+i)).setValue('IT-BUรับของแล้ว');
    }
    var values = BCT.getValuesAll(sheet,rowStartValue,1);
    for(var v=0;v<values.length;v++){
      var newValues = [values[v]];
      if(newValues[0][1] == 'X')
      {
        if(newValues[0][pr_no-1]!=''){          
          queryInsert += BCT.createQueryUpdateStr(tableName, fields, [newValues[0]], 'pr_no', '', '', '');    
        }
      }
    }
//    SpreadsheetApp.getUi().alert(queryInsert);
//    return;
    
    var xml = BCT.loadXMLQueryInsertUpdateMulti(BCT.getDBServer(), 'BCT_Asset_Pkg', queryInsert);  //update table purchase
    Logger.log(xml)
 BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange      
//    var xmlDoc = XmlService.parse(xml);
//    if(Number(xmlDoc.getRootElement().getChildText('status'))==1){
  //     indexSaveFunctionOption2();  //บันทึกเข้า stockAGS

      // SpreadsheetApp.getUi().alert('บันทึกเรียบร้อย');
//    }else{
//      SpreadsheetApp.getUi().alert('แจ้งเตือน', 'Error : \n'+xmlDoc.getRootElement().getChildText('error')+'\nกรุณาบันทึกใหม่', SpreadsheetApp.getUi().ButtonSet.OK);
//    }

  }

}
