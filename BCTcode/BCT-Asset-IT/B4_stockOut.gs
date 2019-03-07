function stock2OUTsheet() { //sheet4.1
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();
  var numCols = sheet.getLastRow();
  var dataRange = sheet.getRange(eventRow, 3, numRows,numCols);
//  Browser.msgBox(dataRange);return;
  var sheet2 = ss.setActiveSheet(ss.getSheetByName('B4.2_เบิกสต็อก'));  
  var sh2lastrow = sheet2.getLastRow();
  sheet2.insertRowsAfter(sh2lastrow, numRows)
  dataRange.copyTo(sheet2.getRange(sh2lastrow+1,5));
  
}

///function สำหรับชีท 4.2 stockOUT
function save_update_stockOUT(){   
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');
  
  var rowFields_out = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields_out = BCT.getFields(sheet, rowFields_out, 1, 0);  
  var qtyOLD = BCT.nameColumnByFliedName(fields_out, 'qty');
  var qtyOUT = BCT.nameColumnByFliedName(fields_out, 'qty_out');
  
  if(eventRow>=rowStartValue){  
    var qtyCheck = 0;
    for(var j=0;j<numRows;j++)
    {
      var qtyOld_value = sheet.getRange(qtyOLD+(eventRow+j)).getValue();
      var qtyOut_value = sheet.getRange(qtyOUT+(eventRow+j)).getValue();
      if(qtyOld_value < qtyOut_value)
      {
        SpreadsheetApp.getUi().alert('มีเบิกเกินจำนวนสต็อกที่มีอยู่ กรุณตรวจเช็คก่อนบันทึกใหม่');
        qtyCheck = 0;
        return;
      } 
      else { qtyCheck =1;}
    }
    
    if(qtyCheck ==1)
    {
      for(var j=0;j<numRows;j++)
      {
        var qtyOld_value = sheet.getRange(qtyOLD+(eventRow+j)).getValue();
        var qtyOut_value = sheet.getRange(qtyOUT+(eventRow+j)).getValue();       
        var qtyNew_value = qtyOld_value - qtyOut_value;
        sheet.getRange(qtyOLD+(eventRow+j)).setValue(qtyNew_value);
      }
      var rowFields = BCT.form_getRowFieldsByKey(sheet, 'update');
      var fields = BCT.getFields(sheet, rowFields, 1, 0);       
      var tableName = 'stock_IT';
      var queryInsert ='';
      var runningCol = BCT.numberColumnByFliedName(fields, 'running');
      var statusCol = BCT.nameColumnByFliedName(fields_out, 'status');
      for(var i = 0;i<numRows;i++)
      {
        sheet.getRange(statusCol+(eventRow+i)).setValue('บันทึกการเบิกแล้ว');
      }
      var values = BCT.getValuesAll(sheet,rowStartValue,1);
      for(var v=0;v<values.length;v++){
        var newValues = [values[v]];
        if(newValues[0][1] == 'X')
        {
          if(newValues[0][runningCol-1]!=''){          
            queryInsert += BCT.createQueryUpdateStr(tableName, fields, [newValues[0]], 'running', '', '', '');    
          }
        }
      }
      //    SpreadsheetApp.getUi().alert(queryInsert);
      //    return;
      

      var xml = BCT.loadXMLQueryInsertUpdateMulti(BCT.getDBServer(), 'BCT_Asset_Pkg', queryInsert);  
      Logger.log(xml)
      var xmlDoc = XmlService.parse(xml);
      if(Number(xmlDoc.getRootElement().getChildText('status'))==1){
         BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange  
//        indexSaveFunctionOption2(); //บันทึกการเบิก table stock_out_IT
        // SpreadsheetApp.getUi().alert('บันทึกเรียบร้อย');
      }else{
        SpreadsheetApp.getUi().alert('แจ้งเตือน', 'Error : \n'+xmlDoc.getRootElement().getChildText('error')+'\nกรุณาบันทึกใหม่', SpreadsheetApp.getUi().ButtonSet.OK);
      }
    } //end if qtyCheck 
    
  } //end if eventrow
}

function clearRowStockOut(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var rowStartValue = BCT.form_getRowStartValueByKey(sheet, 'process');
  var lastrow = sheet.getLastRow();
  if (lastrow > rowStartValue)
  {
    sheet.deleteRows(rowStartValue+1,(lastrow-rowStartValue));
  }
  
}