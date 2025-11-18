function id(){
  var s = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(206660123); //data0
  var s2 = getSheetById(0); //create
  var s3 = getSheetById(21476299); //data
  
  var code = s3.getRange(1,1).getValue();
  if(code==""){
    ui.alert("ERROR", "DON'T LEAVE IT BLANK", ui.ButtonSet.OK);
  }
  else{
    var name = s2.getRange(2,1).getValue();
    var colorcode = s1.getRange(1,7).getValue();
    var date =new Date();  
    var f = DriveApp.getFolderById('1zCz8aWH5ik1w2q7DTf0on3WENsCL8lJ9').createFolder(name);
    f.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    var document = DriveApp.getFileById('1EPkrh_04fRKgl6UjDzA77dvJaVOerwr63a0pDuVEFXQ').makeCopy();
    var ppt = DriveApp.getFileById('1S0dklPMgf37gKUQfIZVS3MQJ_I2DkI9PO_fKQR_LdHE').makeCopy();

    DriveApp.getFileById(document.getId()).setName(name).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    DriveApp.getFileById(ppt.getId()).setName(name).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    
    var s1val = [[date,f.getUrl(),document.getUrl(),ppt.getUrl()]];
    var s2val = [[f.getUrl()],[document.getUrl()],[ppt.getUrl()]];
    //save
    s1.insertRowBefore(2);
    s1.getRange(2,1,1,4).setValues(s1val);
    s3.getRange('A1').copyTo(s1.getRange('E2'), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    
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
    
    //settitle
    settitle(name,document);
    
    //reset dropdown
    s2.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
                                                 .setAllowInvalid(false)
                                                 .requireValueInRange(s3.getRange('$G$2:$G'), true)
                                                 .build());
  }
}


function settitle(name,document){
  var doc = DocumentApp.openByUrl(document.getUrl());
  var body = doc.getBody();
  var para = body.getParagraphs()
  
  para[0].setText(name);
  para[0].setAttributes({FONT_SIZE:24})
}

function getSheetById(id) {
  return SpreadsheetApp.getActive().getSheets().filter(
    function(s) {return s.getSheetId() === id;}
  )[0];
}

