function save(){
  var s =SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(26441098); //活动定时;
  var s2 = getSheetById(884156533); //data0;
  var lastrow = s2.getMaxRows();
  
  var check1 = s2.getRange(1,1,3,1).getValues();
  var check2 = s2.getRange(12,1).getValue();
  var actId =  s2.getRange(15,1).getValue(); 
  var check_null_allow = s2.getRange(5,1,2,1).getValues();
  var msg = [["活动种类Error","请输入一个有效的“活动”"], ["叙述Error", "请输入一个有效的叙述"], ["日期Error", "请输入一个有效的日期"]]
  var c = 0

  if(check_null_allow[0][0]==""||check_null_allow[1][0]!=""){
    //  check1
    for(var i=0;i<3;i++){
      if(check1[i][0]==""){
        ui.alert(msg[i][0],msg[i][1],ui.ButtonSet.OK);
        c=c+1
      }
    }
  }  
  //save
  if(c==0){
    //check2
    if(check2!=""){
      ui.alert("重复Error","此活动已被输入进系统",ui.ButtonSet.OK);
    }
    else{
      //save
      s2.getRange(2,16,lastrow-1,4).copyTo(s2.getRange(2,12,lastrow-1,4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      
      //act code
      s2.getRange(15, 1).setValue(actId+1)
      
      //clear
      s1.getRange(1,2,3,1).clearContent();
      
      //beautify
      s1.insertRowsBefore(9,1);
      
      //clear error
      s2.getRange(5,1).clearContent();
      
      //reset dropdown
      s1.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation()
                                          .setAllowInvalid(false)
                                          .requireValueInRange(s1.getRange('data0!$F$2:$F'), true)
                                          .build());
      s1.getRange('C2').setDataValidation(SpreadsheetApp.newDataValidation()
                                          .setAllowInvalid(false)
                                          .requireValueInRange(s1.getRange('data0!$K$2:$K'), true)
                                          .build());
    }
    
  }
  
}

function error(){
  var s =SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(26441098); //活动定时;
  var s2 = getSheetById(884156533); //data0;
  var check = s2.getRange(3,20).getValue();
  
  if(check==""){
    ui.alert("Error","请选一个活动",ui.ButtonSet.OK);
  }
  else{
    
    //error
    s2.getRange(2,20,3,1).copyTo(s1.getRange(1,2,3,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //save
    s1.getRange(2,3,1,1).copyTo(s2.getRange(5,1,1,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //clear
    s1.getRange(2,3).clearContent();
  }
}






function update1(){
  var s =SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s2 = getSheetById(884156533); //data0;
  var lastrow = s2.getMaxRows();
  var num1 = s2.getRange(1,31).getValue();
  var num2 = s2.getRange(1,32).getValue();
  
  if(num1==0){
    ui.alert("Empty Error","请输入至少一个活动类型",ui.ButtonSet.OK);
  }
  else{
    
    var form = FormApp.openByUrl("https://docs.google.com/forms/d/1-mjIblBK9DcHxuSEUHk92YlK9pW-03jfGkYPpJa7aqQ/edit");
    var item = form.getItems();
    var array1 = s2.getRange(2,31,num1,1).getValues().map(function(d){ return d[0];});
    var array2 = s2.getRange(2,32,num2,1).getValues().map(function(d){ return d[0];});
    
    //save
    s2.getRange(2,31,lastrow-1,2).copyTo(s2.getRange(2,29,lastrow-1,2), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
    //form
    item[3].asMultipleChoiceItem().setChoiceValues(array1);
    item[4].asMultipleChoiceItem().setChoiceValues(array2);
  }
}


function update2(){
  var s =SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1211536927); //活动种类(定);
  var s2 = getSheetById(884156533); //data0;
  var lastrow = s1.getMaxRows();
  
  var row_span = s2.getRange(1,27).getValue();
  
  if(row_span<1){
    ui.alert("Empty Error","请输入至少一个组",ui.ButtonSet.OK);
  }
  else{
    s2.getRange(2,25,row_span*2,1).copyTo(s2.getRange(2,24,row_span*2,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    s1.getRange(2,6,lastrow-1,2).clearContent();
    s2.getRange(2,26,row_span,2).copyTo(s1.getRange(2,6,row_span,2), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
  }
  
}

function update3(){
  var s =SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s2 = getSheetById(884156533); //data0;
  var lastrow = s2.getMaxRows();
  
  s2.getRange(2,9,lastrow-1,1).copyTo(s2.getRange(2,6,lastrow-1,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  s2.getRange(2,10,lastrow-1,1).copyTo(s2.getRange(2,8,lastrow-1,1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
}



function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}
