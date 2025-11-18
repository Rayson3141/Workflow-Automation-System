function dropdown1(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(0); //活动planning ;
  var spreadsheet0 = getSheetById(1033467041); //data0 ;
  var ui = SpreadsheetApp.getUi();
  
  var row_span = spreadsheet0.getRange(3,1).getValue();
  var col_span = spreadsheet.getRange(5,2).getValue();
  var col_lim = 5
  var row_lim = 10
  
  if((row_span-(row_lim+1))*(row_span)>=0){
    ui.alert("组No. Error","请输入1到"+row_lim+"的数字", ui.ButtonSet.OK);
  }
  else if((col_span-(col_lim+1))*(col_span)>=0){
    ui.alert("DropdownNo. Error","请输入1到"+col_lim+"的数字", ui.ButtonSet.OK);
  }else{
    dropdown_1(row_span,col_span);
  }  
}

function save1(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(0); //活动planning ;
  var spreadsheet0 = getSheetById(1033467041); //data0 ;
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1035613208); //data;
  
  var lastrow2 = spreadsheet0.getMaxRows();
  var name = spreadsheet0.getRange(4,1).getValue();
  
  if(name==""){
    ui.alert("第几次学会活动??","请选一个学会活动", ui.ButtonSet.OK);
  }
  else{
    //copy
    spreadsheet0.getRange(2,24,lastrow2-1,4)
    .copyTo(spreadsheet0.getRange(2,20,lastrow2-1,4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //clear
    clear1();
    spreadsheet.getRange(5,1).clearContent();
    
    //save
    spreadsheet0.getRange(9,1).clearContent();
    spreadsheet0.getRange(11,1).clearContent();
    
    //reset dropdown
    spreadsheet.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .setAllowInvalid(false)
                                                 .requireValueInRange(s1.getRange('$J$2:$J'), true)
                                                 .build());
    spreadsheet.getRange('A5').setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .setAllowInvalid(false)
                                                 .requireValueInRange(s1.getRange('$N$2:$N'), true)
                                                 .build());
    
  }
}

function error1(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s1 = getSheetById(0); //活动planning ;
  var s2 = getSheetById(1033467041); //data0 ;
  var ui = SpreadsheetApp.getUi();
  var name1 = s2.getRange(7,1).getValue();
  var savecheck = s2.getRange(11,1).getValue();

  if(name1==""){
    ui.alert("第几次学会活动??（Error）","请选一个学会活动", ui.ButtonSet.OK);
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
  var ss1 = getSheetById(430900926); //筹主，主讲人
  var ss2 = getSheetById(1033467041); //data0 ;
  var ui = SpreadsheetApp.getUi();
  
  var name = ss2.getRange(5,1).getValue();
  var num_incharge = ss1.getRange(4,3).getValue();
  var speaker_check = ss1.getRange(5,3).getValue();
  
  if(name==""){
    ui.alert("第几次学会活动??","请选一个学会活动", ui.ButtonSet.OK);
  }else if(num_incharge<0){
    ui.alert("筹组No. Error","筹组No.>= 0", ui.ButtonSet.OK);
  }
  else if(speaker_check<0){
    ui.alert("主讲人No. Error","主讲人No.>= 0 ", ui.ButtonSet.OK);
  }
  else{
    var num_speaker = ss2.getRange(1, 17).getValue();
    
    if(num_speaker==""){
      num_speaker=0
    }
    if(num_incharge==""){
      num_incharge=0
    }
    
    dropdown_2(num_incharge,num_speaker);
  }   

}

function save2(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var spreadsheet = getSheetById(430900926); //筹主，主讲人
  var spreadsheet0 = getSheetById(1033467041); //data0 ;
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1035613208); //data;
  var name = spreadsheet0.getRange(5,1).getValue();
  var lastrow = spreadsheet.getMaxRows()
  var lastrow2 = spreadsheet0.getMaxRows();
  
  if(name==""){
    ui.alert("第几次学会活动??","请选一个学会活动", ui.ButtonSet.OK);
  }
  else{
    //copy
    spreadsheet0.getRange(2,34,lastrow2-1,2)
    .copyTo(spreadsheet0.getRange(2,32,lastrow2-1,2), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    spreadsheet0.getRange(2,41,lastrow2-1,3)
    .copyTo(spreadsheet0.getRange(2,38,lastrow2-1,3), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //clear
    clear2();
    
    //save
    spreadsheet0.getRange(10,1).clearContent();
    spreadsheet0.getRange(12,1).clearContent();
    if(lastrow>8){
      spreadsheet.deleteRows(9,lastrow-8);
    }
    
    //reset dropdown
    spreadsheet.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .setAllowInvalid(false)
                                                 .requireValueInRange(s1.getRange('$L$2:$L'), true)
                                                 .build());
    spreadsheet.getRange('A5').setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .setAllowInvalid(false)
                                                 .requireValueInRange(s1.getRange('$P$2:$P'), true)
                                                 .build());
  }
}

function error2(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var ss1 = getSheetById(430900926); //筹主，主讲人;
  var ss2 = getSheetById(1033467041); //data0 ;
  
  var name1 = ss2.getRange(8,1).getValue();
  var savecheck = ss2.getRange(12,1).getValue();
  
  if(name1==""){
    ui.alert("第几次学会活动??（Error）","请选一个学会活动", ui.ButtonSet.OK);
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




