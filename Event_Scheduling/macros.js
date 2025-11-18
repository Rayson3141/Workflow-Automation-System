function CREATE() {
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(206660123); //data
  var s2 = getSheetById(0); //create
  var name = s2.getRange(2,1).getValue();
  
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
      
      //    var f = DriveApp.getFolderById('1W-_Fp_mbHPp_QCzsU7bPscqK9yXv1hNa').createFolder(name);
      //    f.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
      //    var document = DriveApp.getFileById('1wrvrZzlqYNSODdWPlO3d5bFNn_o8MwIO-V19Tt6dcrU').makeCopy();
      //    var ppt = DriveApp.getFileById('10JID0C7bFPKlodhWOuw5XTZxWUjVAYs7d6cZ1EMYBFA').makeCopy();
      
      //set name
      DriveApp.getFolderById(f).setName(name)
      DriveApp.getFileById(document).setName(name);
      DriveApp.getFileById(ppt).setName(name);
      
      //      var s1val = [[date,f.getUrl(),document.getUrl(),ppt.getUrl()]];
      var s2val = [[val[0][0]],[val[0][1]],[val[0][2]]];
      
      //save
      //      s1.insertRowBefore(2);
      s1.getRange(k+1,1).setValue(date);
      s2.getRange('A2').copyTo(s1.getRange(k+1,5), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
      
      //present
      s2.getRange(8,1,3,1).setValues(s2val);
      
      //color
      if(colorcode==1){
        s1.getRange(1,7).setValue(2);
      }
      else{
        s1.getRange(1,7).setValue(1);
      }
      
      //clear
      s2.getRange(2,1).clearContent();
      
      
    }
    
    
    
    
  }
}

//function CREATE() {
//  var s = SpreadsheetApp.getActiveSpreadsheet();
//  var ui = SpreadsheetApp.getUi();
//  var s1 = getSheetById(206660123); //data
//  var s2 = getSheetById(0); //create
//  var name = s2.getRange(2,1).getValue();
//  if(name==""){
//    ui.alert("ERROR", "DON'T LEAVE IT BLANK", ui.ButtonSet.OK);
//  }
//  else{
//    var colorcode = s1.getRange(1,7).getValue();
//    var date =new Date();  
//    var f = DriveApp.getFolderById('1W-_Fp_mbHPp_QCzsU7bPscqK9yXv1hNa').createFolder(name);
//    f.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
//    var document = DriveApp.getFileById('1wrvrZzlqYNSODdWPlO3d5bFNn_o8MwIO-V19Tt6dcrU').makeCopy();
//    var ppt = DriveApp.getFileById('10JID0C7bFPKlodhWOuw5XTZxWUjVAYs7d6cZ1EMYBFA').makeCopy();
//    
//    DriveApp.getFileById(document.getId()).setName(name).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
//    DriveApp.getFileById(ppt.getId()).setName(name).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
//    
//    var s1val = [[date,f.getUrl(),document.getUrl(),ppt.getUrl()]];
//    var s2val = [[f.getUrl()],[document.getUrl()],[ppt.getUrl()]];
//    //save
//    s1.insertRowBefore(2);
//    s1.getRange(2,1,1,4).setValues(s1val);
//    s2.getRange('A2').copyTo(s1.getRange('E2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
//    
//    //present
//    s2.getRange(8,1,3,1).setValues(s2val);
//    
//    //color
//    if(colorcode==1){
//      s1.getRange(1,7).setValue(2);
//    }
//    else{
//      s1.getRange(1,7).setValue(1);
//    }
//    
//    //clear
//    s2.getRange(2,1).clearContent();
//    
//    //showUrl
//    showUrl(f,ppt,document);
//  }
//};

function showUrl(f,ppt,document){
  var doc = DocumentApp.openByUrl(document.getUrl());
  var body = doc.getBody();
  var tables = body.getTables();
  var url = [f.getUrl(),document.getUrl(),ppt.getUrl()];
  var name = ["Folder","活动","活动ppt"]
  
  
  for(var i=0;i<3;i++){
    var table = tables[5];
    table.getCell(i+1,1).setText(name[i]+" Url: ");
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


