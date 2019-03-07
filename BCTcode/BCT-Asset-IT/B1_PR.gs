function sendLine_PR() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var filelink = ss.getUrl();
  var sheetID = sheet.getSheetId();
  var lineToken = sheet.getRange('D1').getValue();

  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, rowActive, 1);

  var date_pr0 = BCT.valueByFliedName(fields, values, 'date_add');
//  Browser.msgBox(date_pr0);    
    var date_pr = Utilities.formatDate(date_pr0, "GMT+07:00", "yyyy-MM-dd")
  var product = BCT.valueByFliedName(fields, values, 'product_name');
  var product_detail = BCT.valueByFliedName(fields, values, 'product_detail');
  var qty = BCT.valueByFliedName(fields, values, 'qty');
  var unit = BCT.valueByFliedName(fields, values, 'unit');
  var price_unit = BCT.valueByFliedName(fields, values, 'price_unit');
  var price_total = BCT.valueByFliedName(fields, values, 'price_total');
  var user_request = BCT.valueByFliedName(fields, values, 'user_request');
  var team = BCT.valueByFliedName(fields, values, 'team');
  var cause = BCT.valueByFliedName(fields, values, 'cause_request');
  var approver = BCT.valueByFliedName(fields, values, 'approve_name');
  var note = BCT.valueByFliedName(fields, values, 'note');  
  var pr_user = BCT.valueByFliedName(fields, values, 'pr_create_user'); 
  var cash_plan = BCT.valueByFliedName(fields, values, 'cash_plan'); 
  var groupBU = BCT.valueByFliedName(fields, values, 'groupBU');
  var BUuse = BCT.valueByFliedName(fields, values, 'BU');
  var BUpay = BCT.valueByFliedName(fields, values, 'BU_pay');  
  
//  var planBalance = checkCashPlanBalance(cash_plan);
////  Browser.msgBox(planBalance);
//  if(planBalance < price_total || planBalance == '')
//  {
//    Browser.msgBox('!!!ยอดเงินคงเหลือแผนงานไม่เพียงพอ');
//    return;
//  }
  
  var subject = '\n'+groupBU+' : แจ้งการเสนอซื้อ ';
  var message = 'วันที่'+date_pr+'\nเสนอซื้อ '+product+' '+product_detail+'\nจำนวน'+qty+' '+unit+' ราคา'+unit+'ละ '+price_unit+' รวม '+price_total
                +' บาท\nลูกค้าผู้ขอ '+user_request+' ทีม '+team+'BU '+BUuse
                +' \nBUจ่ายเงิน '+BUpay
                +' \nเหตุผลการซื้อ '+cause
                +'\nชื่อผู้อนุมัติ : '+approver+'\nคลิ๊กอนุมัติได้ที่ '+filelink+'#gid='+sheetID
                +'\nหมายเหตุ : '+note
                +'\nผู้บันทึกPR : '+pr_user;
                
//  BCTCENTER.LineNotify("8IGwhopbpj22QR733Sf54OaC5cFw7pN5VvJS6Llc4ND", subject, message);//line kowit
  BCTCENTER.LineNotify(lineToken, subject, message);//line FAM+AGS
    
  var send2PR_col = BCT.nameColumnByFliedName(fields, 'send_PR_date');
  var date_send_PR = sheet.getRange(send2PR_col+rowActive).setValue(new Date());  
  
  BCT.saveDataSpreadsheetByTemplate(true); 
  
}

function sendLine_to_PO(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();
  var lineToken = sheet.getRange('D1').getValue();

  var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
  var fields = BCT.getFields(sheet, rowFields, 1, 0);   
  var values = BCT.getValue(sheet, rowActive, 1);

  var date_pr0 = BCT.valueByFliedName(fields, values, 'date_add');
    var date_pr = Utilities.formatDate(date_pr0, "GMT+07:00", "yyyy-MM-dd");
  var product = BCT.valueByFliedName(fields, values, 'product_name');
  var product_detail = BCT.valueByFliedName(fields, values, 'product_detail');
  var qty = BCT.valueByFliedName(fields, values, 'qty');
  var unit = BCT.valueByFliedName(fields, values, 'unit');
  var price_unit = BCT.valueByFliedName(fields, values, 'price_unit');
  var price_total = BCT.valueByFliedName(fields, values, 'price_total');
  var user_request = BCT.valueByFliedName(fields, values, 'user_request');
  var team = BCT.valueByFliedName(fields, values, 'team');
  var BUuse = BCT.valueByFliedName(fields, values, 'BU');  
  var cause = BCT.valueByFliedName(fields, values, 'cause_request');
  var BU_pay = BCT.valueByFliedName(fields, values, 'BU_pay');  
  var note = BCT.valueByFliedName(fields, values, 'note');    
  var BUbill = BCT.valueByFliedName(fields, values, 'BU_pay_billname');    
  
  
  var subject = '\nแจ้งFAM : ทำการจัดซื้อ';
  var message = 'วันที่'+date_pr
                +'\nแจ้งจัดซื้อ '+product+' '+product_detail+'\nจำนวน'+qty+' '+unit+' ราคา'+unit+'ละ '+price_unit+' รวม '+price_total+' บาท'
                +'\nเป็นค่าใช้จ่ายของBU : '+BU_pay
                +'\nลงบิล : '+BUbill
                +'\nผู้ลูกค้าที่ขอ '+user_request+' ทีม '+team+'BU '+BUuse
                +' \nเหตุผลการซื้อ '+cause
                +'\nหมายเหตุ : '+note;
                
//  BCTCENTER.LineNotify("8IGwhopbpj22QR733Sf54OaC5cFw7pN5VvJS6Llc4ND", subject, message);//line kowit
//  BCTCENTER.LineNotify("xujgaQsWgJxH61NBQyBAOsKjapjh6jrMHGD0JjK41h0", subject, message);//line FAM+AGS
  BCTCENTER.LineNotify(lineToken, subject, message);//line FAM+AGS
  
//  var status_col = BCT.nameColumnByFliedName(fields, 'status');
 // var status = sheet.getRange(status_col+rowActive).setValue('ส่งเข้าระบบจัดซื้อแล้ว');
  var send2PO_col = BCT.nameColumnByFliedName(fields, 'send_toPO_date');
  var date_send_PO = sheet.getRange(send2PO_col+rowActive).setValue(new Date());  
  
 BCT.saveDataSpreadsheetByTemplate(true); //save เฉพาะตาม eventRange  
//  BCT.saveDataSpreadsheetByTemplate(true); 
  
}


function checkCashPlanBalance(cash_plan)
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName(); 
  
  var DBName = "BCT_ACC";
  var query = "select * from budjetPlan_17 where planCode ='"+cash_plan+"'";
  var datas = BCT.loadJSONDatas(BCT.getDBServer(), DBName, query);
  var plan_balance = datas[0]['balance'];
  if (plan_balance != '')
  {
    return plan_balance;
  }
  else { Browser.msgBox('ไม่พบยอดคงเหลือแผนการเงิน'); }
   Logger.log('Balance : ', plan_balance);
   Logger.log('datas : ', datas);    
}

function cancelPR()
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet()
  var sheetname = sheet.getSheetName();
  var rowActive = sheet.getActiveCell().getRow();

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('ยืนยันการยกเลิกPR', ui.ButtonSet.YES_NO);

 // Process the user's response.
   if (response == ui.Button.YES) {
      Logger.log('The user clicked "Yes."');
      var rowFields = BCT.form_getRowFieldsByKey(sheet, 'process');
      var fields = BCT.getFields(sheet, rowFields, 1, 0);   
      var values = BCT.getValue(sheet, rowActive, 1); 
  
      var pr_no = BCT.valueByFliedName(fields, values, 'pr_no');  
      var status_col = BCT.nameColumnByFliedName(fields, 'status');
      var status = sheet.getRange(status_col+rowActive).setValue('ยกเลิกPR'); 
      sheet.getRange('B'+rowActive).setValue('X');
  
     // BCT.saveDataSpreadsheetByTemplate(true);      
   } else {
   Logger.log('The user clicked "No" or the dialog\'s close button.');
   }  
}