function createPO() {
 var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();

  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, eventRow, 1);
  
  var BU = BCT.valueByFliedName(fields, values, 'BU_pay');
  var date_PO = BCT.valueByFliedName(fields, values, 'po_create_date');
  if(date_PO =='')
  { Browser.msgBox('กรุณากรอกวันที่ใบPOก่อน'); }
  else
  {
    var year_po = Utilities.formatDate(date_PO,"GMT+07:00", "yyyy");  
    var month_po = Utilities.formatDate(date_PO, "GMT+07:00", "MM"); 
 
  var monthCol = BCT.nameColumnByFliedName(fields, 'po_create_user');
    var month_number = sheet.getRange(monthCol+9).getValue(); //ช่องที่เก็บเลขเดือนใบ PO
  var po_run_Col = BCT.nameColumnByFliedName(fields, 'po_no');
    var po_run = sheet.getRange(po_run_Col+9).getValue(); //เก็บค่าเลขrun ใบ PO

    if(month_po != month_number)
    {
      sheet.getRange(monthCol+9).setValue(month_po);
      sheet.getRange(po_run_Col+9).setValue('1');
      po_run = 1;
    } else {
       po_run = po_run+1;
       sheet.getRange(po_run_Col+9).setValue(po_run);
    }
   var po_no_promt = BU+'-'+year_po+month_po+'-'+po_run;
   var po_noCol = BCT.numberColumnByFliedName(fields, 'po_no');
   var statusCol = BCT.numberColumnByFliedName(fields, 'status');    
//  Browser.msgBox(po_no_promt);
   sheet.getRange(eventRow, po_noCol, numRows).setValue(po_no_promt);
   sheet.getRange(eventRow, statusCol, numRows).setValue('ออกใบ PO แล้ว');    


   BCT.saveDataSpreadsheetByTemplate(true);
  }    
}

function savePurchased(){  //ปุ่มบันทึกสั่งซื้อแล้ว
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var sheetName = sheet.getSheetName();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();  
  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, eventRow, 1);  
  
   var statusCol = BCT.numberColumnByFliedName(fields, 'status');    
   sheet.getRange(eventRow, statusCol, numRows).setValue('สั่งซื้อแล้ว');
  
  BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange
  
}

function sendLine_Receive_Product() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var filelink = ss.getUrl();
  var sheetID = sheet.getSheetId();
  var eventRange = ss.getActiveRange();
  var eventRow = eventRange.getRow();
  var numRows = eventRange.getNumRows();    
  var lineToken = sheet.getRange('D1').getValue();  

  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, rowActive, 1);

  var date_re = BCT.valueByFliedName(fields, values, 'FAM_receive_date');
//  Browser.msgBox(date_pr0);    
  var date_receive = Utilities.formatDate(date_re, "GMT+07:00", "yyyy-MM-dd")
  var product = BCT.valueByFliedName(fields, values, 'product_name');
  var product_detail = BCT.valueByFliedName(fields, values, 'product_detail');
  var qty = BCT.valueByFliedName(fields, values, 'qty');
  var unit = BCT.valueByFliedName(fields, values, 'unit');
  var price_unit = BCT.valueByFliedName(fields, values, 'price_unit');
  var price_total = BCT.valueByFliedName(fields, values, 'price_total');
  var groupBU = BCT.valueByFliedName(fields, values, 'groupBU');
  var BUuse = BCT.valueByFliedName(fields, values, 'BU');
  var BUpay = BCT.valueByFliedName(fields, values, 'BU_pay'); 
//  var user_request = BCT.valueByFliedName(fields, values, 'user_request');
//  var team = BCT.valueByFliedName(fields, values, 'team');
//  var cause = BCT.valueByFliedName(fields, values, 'cause_request');
//  var approver = BCT.valueByFliedName(fields, values, 'approve_name');
  
  var subject = '\nแจ้งIT-'+groupBU+' : สินค้าที่ซื้อได้รับของแล้ว ';
  var message = 'วันที่'+date_receive+'\nสินค้า '+product+' '+product_detail+'\nจำนวน'+qty+' '+unit+' ราคา'+unit+'ละ '+price_unit+' รวม '+price_total;
                
//  BCTCENTER.LineNotify("8IGwhopbpj22QR733Sf54OaC5cFw7pN5VvJS6Llc4ND", subject, message);//line kowit
  BCTCENTER.LineNotify(lineToken, subject, message);//line FAM+AGS

   var statusCol = BCT.numberColumnByFliedName(fields, 'status');    
   sheet.getRange(eventRow, statusCol, numRows).setValue('FAMรับของแล้ว');

 // indexSaveFunctionOption2(); 
 BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange    
  //BCT.saveDataSpreadsheetByTemplate(true); 
  
}
