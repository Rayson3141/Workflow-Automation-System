function dropdown1(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(0);
  var spreadsheet0 = getSheetById(1033467041);
  var ui = SpreadsheetApp.getUi();
  
  var row_span = spreadsheet0.getRange(3,1).getValue();
  var col_span = spreadsheet.getRange(5,2).getValue();
  var col_lim = 50
  var row_lim = 50
  
  if((row_span-(row_lim+1))*(row_span)>=0){
    ui.alert("活动No. Error","只能有1到"+row_lim+"个活动", ui.ButtonSet.OK);
  }
  else if((col_span-(col_lim+1))*(col_span)>=0){
    ui.alert("DropdownNo. Error","请输入1到"+col_lim, ui.ButtonSet.OK);
  }else{
    dropdown_1(row_span,col_span);
  }  
}

function save1(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(0);
  var spreadsheet0 = getSheetById(1033467041);
  var ss = getSheetById(1035613208);
  var ui = SpreadsheetApp.getUi();
  
  var lastrow2 = spreadsheet0.getMaxRows();
  var name = spreadsheet0.getRange(4,1).getValue();
  var divdate = spreadsheet.getRange(2,2).getValue();
  var divdate_lim = spreadsheet0.getRange(3,1).getValue();
  
  if(name==""){
    ui.alert("第几年??","请选一年", ui.ButtonSet.OK);
  }
  else if((divdate-(divdate_lim+1))*(divdate)>=0){
    ui.alert("下半年分割线 Error","请输入1到"+(divdate_lim)+"的数字", ui.ButtonSet.OK);
  }
  else{
    //copy
    var x = spreadsheet0.getRange(1,8).getValue();
  
    spreadsheet0.getRange(2,20,lastrow2-1,3)
    .copyTo(spreadsheet0.getRange(2,17,lastrow2-1,3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet0.getRange(2,30,lastrow2-1,2)
    .copyTo(spreadsheet0.getRange(2,28,lastrow2-1,2), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //clear
    clear1();
    spreadsheet.getRange(5,1).clearContent();
    spreadsheet0.getRange(8,1).clearContent();
    
    //save
    spreadsheet0.getRange(11,1).clearContent();
    
    //reset dropdown 
    spreadsheet.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(ss.getRange('$M$2:$M'), true)
    .build());
    
    spreadsheet.getRange('A5').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(ss.getRange('$Q$2:$Q'), true)
    .build());
    formula1(x);
  }
}

function error1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s2 = getSheetById(1033467041);
  var ui = SpreadsheetApp.getUi();
  var name1 = s2.getRange(7,1).getValue();
  var savecheck = s2.getRange(11,1).getValue();

  if(name1==""){
    ui.alert("第几年??（Error）","请选一年", ui.ButtonSet.OK);
  }
  else if(savecheck==""){
    error_1();
  }
  else{
    var button = ui.alert("Did u save", "Are u sure u don't want to save",ui.ButtonSet.YES_NO);
    
    if(button == ui.Button.YES){
      error_1();
    }
    else{
    }
  }

}





function dropdown2(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(430900926);
  var spreadsheet0 = getSheetById(104186564);
  var ui = SpreadsheetApp.getUi();
  
  var row_span = spreadsheet.getRange(4,3).getValue();
  var col_span = spreadsheet.getRange(5,3).getValue();
  var col_lim = 50
  var row_lim = 50
  
  if((row_span-(row_lim+1))*(row_span)>=0){
    ui.alert("执委No. Error","请输入组1到"+row_lim+"的数字", ui.ButtonSet.OK);
  }
  else if((col_span-(col_lim+1))*(col_span)>=0){
    ui.alert("DropdownNo. Error","请输入1到"+col_lim+"的数字", ui.ButtonSet.OK);
  }else{
    dropdown_2(row_span,col_span);
  }  

}

function save2(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(430900926);
  var spreadsheet0 = getSheetById(104186564);
  var ss = getSheetById(1035613208);
  var ui = SpreadsheetApp.getUi();
  
  var lastrow2 = spreadsheet0.getMaxRows();
  var name = spreadsheet0.getRange(4,1).getValue();
  
  if(name==""){
    ui.alert("第几年??","请选一年", ui.ButtonSet.OK);
  }
  else{
    var x = spreadsheet0.getRange(1,8).getValue();

    //copy
    spreadsheet0.getRange(2,20,lastrow2-1,3)
    .copyTo(spreadsheet0.getRange(2,17,lastrow2-1,3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //clear
    clear2();
    spreadsheet.getRange(5,1).clearContent();
    spreadsheet0.getRange(8,1).clearContent();
    
    //save
    spreadsheet0.getRange(11,1).clearContent();
    
    //reset dropdown 
    spreadsheet.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(ss.getRange('$O$2:$O'), true)
    .build());
    
    spreadsheet.getRange('A5').setDataValidation(SpreadsheetApp.newDataValidation()
    .setAllowInvalid(false)
    .requireValueInRange(ss.getRange('$S$2:$S'), true)
    .build());
    formula2(x);
  }
}

function error2(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s2 = getSheetById(104186564);
  var ui = SpreadsheetApp.getUi();
  var name1 = s2.getRange(7,1).getValue();
  var savecheck = s2.getRange(11,1).getValue();

  if(name1==""){
    ui.alert("第几年??（Error）","请选一年", ui.ButtonSet.OK);
  }
  else if(savecheck==""){
    error_2();
  }
  else{
    var button = ui.alert("Did u save", "Are u sure u don't want to save",ui.ButtonSet.YES_NO);
    
    if(button == ui.Button.YES){
      error_2();
    }
    else{
    }
  }

}  




