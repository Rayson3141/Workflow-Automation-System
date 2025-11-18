function id(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(206660123); //data
  var s2 = getSheetById(0); //create
  var name = s2.getRange(2,1).getValue();
  var AI_Use = s2.getRange(2,2).getValue(); // check for AI Use
  if(AI_Use==true){
    var prompt = name
    name = GemPrompt("请通过此‘"+prompt+"',想出一个简短吸引人的实验或活动标题,只能有6个字。请使用中文打出来。")
    }
  
  if(name==""){
    ui.alert("ERROR", "DON'T LEAVE IT BLANK", ui.ButtonSet.OK);
  }
  else{
    var n1 = s1.getRange(1,8).getValue();
    var n2 = s1.getRange(1,9).getValue();
    var k = n2-n1;
    
    if(k==0){
      ui.alert("ERROR", "计划书量不足，请通知课活组长。", ui.ButtonSet.OK);
    }
    else{
      var colorcode = s1.getRange(1,7).getValue();
      var date =new Date(); 
      var val = s1.getRange(k+1,2,1,3).getValues();
      var f = getIdFromUrl(val[0][0])[0];
      var document = getIdFromUrl(val[0][1])[0];
      var ppt = getIdFromUrl(val[0][2])[0];
      
      //set name
      DriveApp.getFolderById(f).setName(name)
      DriveApp.getFileById(document).setName(name);
      DriveApp.getFileById(ppt).setName(name);
      
      var s2val = [[val[0][0]],[val[0][1]],[val[0][2]]];
      
      //save
      s1.getRange(k+1,1).setValue(date);
      s1.getRange(k+1,5).setValue(name);
      // s2.getRange('A2').copyTo(s1.getRange(k+1,5), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      
      //present
      s2.getRange(8,1,3,1).setValues(s2val);
      
      //color
      if(colorcode==1){
        s1.getRange(1,7).setValue(2);
      }
      else{
        s1.getRange(1,7).setValue(1);
      }

      //AI create draft doc
      if(AI_Use == true){
        AI_Create(prompt,document)
      }
      //clear
      s2.getRange(2,1).clearContent(); 
    }
  }
}


function showUrl(f,ppt,document){
  var doc = DocumentApp.openByUrl(document.getUrl());
  var body = doc.getBody();
  var tables = body.getTables();
  var url = [f.getUrl(),document.getUrl(),ppt.getUrl()];
  // var name = ["Folder","活动","活动ppt"]  
  for(var i=0;i<3;i++){
    var table = tables[5];
    //table.getCell(i+1,1).setText(name[i]+" Url: ");
    table.getCell(i+1,0).setText(url[i]);
  }
  
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

function getIdFromUrl(url) { 
  return url.match(/[-\w]{25,}/); 
}
