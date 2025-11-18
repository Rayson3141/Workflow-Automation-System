//function id(){
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
//    //move
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
//    
//    //setowner
////    document.setOwner("ps.act.design@gmail.com");
////    ppt.setOwner("ps.act.design@gmail.com");
////    f.setOwner("ps.act.design@gmail.com");
//  }
//}
//
//
//function showUrl(f,ppt,document){
//  var doc = DocumentApp.openByUrl(document.getUrl());
//  var body = doc.getBody();
//  var tables = body.getTables();
//  var url = [f.getUrl(),document.getUrl(),ppt.getUrl()];
//  var name = ["Folder","活动","活动ppt"]
//  
//  
//  for(var i=0;i<3;i++){
//    var table = tables[5].getCell(0,0);
//    table.appendParagraph(name[i]+" Url: ");
//    table.appendParagraph(url[i]);
//    table.appendParagraph("")
//  }
//  
//}
//
//function getSheetById(id) {
//  return SpreadsheetApp.getActive().getSheets().filter(
//    function(s) {return s.getSheetId() === id;}
//  )[0];
//}
////
//function ss(){
//  DriveApp.getFolderById("1ynPpUw3WR5UjWbeNnOatnoLwvmBR6AYm").setOwner("ooirayshua@gmail.com");
//  
////  Logger.log(d)
//
//
//}
