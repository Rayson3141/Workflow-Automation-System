function createdoc(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1355093305);
  var s2 = getSheetById(1620429395);
  var s3 = getSheetById(411658923);
  var s4 = getSheetById(1035613208);
  var checkblank =  s1.getRange(8,1).getValue();
  
  if(checkblank==""){
    ui.alert("哪一年??","请输入一个合适的年份", ui.ButtonSet.OK);  
  }
  else{ 
    //clear past
    s2.getRange(3,1).clearContent();
    s2.getRange(3,1).setValue(2);
    
    //copy doc
    var document = DriveApp.getFileById('1QfonxrHrxJPuMBAgdnMDAnxBuKGuGZSnI-eafgAAPrE').makeCopy();
    var documentUrl = document.getUrl();
    var documentId = document.getId();
    var plantitle = s1.getRange(6,1).getValue();
    var f = DriveApp.getFolderById('1KjAmypcJru20zgUMlupqeDKFzn1UYc29');
    DriveApp.getFileById(documentId).setName(plantitle).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    
    //save
    var s3lastrow = s3.getMaxRows();
    s3.getRange(2,6,s3lastrow,4)
    .copyTo(s3.getRange(2,2,s3lastrow,4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    s3.getRange(2,2).setValue(new Date());
    s3.getRange(2,5).setValue(documentUrl)
    
    
    //create
    var doc = DocumentApp.openByUrl(documentUrl);
    var body = doc.getBody();
    
    writedocument(s1,body);
    s2.getRange(3,1).setValue(1);
    
    //color code
    if(s1.getRange(9,1).getValue()==1){
      s1.getRange(9,1).setValue(2);
    }
    else{
      s1.getRange(9,1).setValue(1);
    }
    s2.getRange(3,1).setValue(documentUrl);
    s2.getRange(2,1).clearContent();
    
    //reset dropdown
    s2.getRange('A2').setDataValidation(SpreadsheetApp.newDataValidation()
                                        .setAllowInvalid(false)
                                        .requireValueInRange(s4.getRange('$Q$2:$Q'), true)
                                        .build());
  }
}



function writedocument(s1,body){  
  var w = [[30.0, 170.25, 247.25, 74.25],[30.0, 170.25, 247.25, 74.25],[229.5, 292.5],[88.5, 73.5, 86.25, 274.5]]
  var titles = [s1.getRange(5,1).getValue(),"人手分配表"]
  var col = [10,16,22,25]
  
  var constant = [s1.getRange(1,col[0]).getValue(),
                  s1.getRange(1,col[1]).getValue(),
                  s1.getRange(1,col[2]).getValue(),
                  s1.getRange(1,col[3]).getValue()];
  
  var arrays = [s1.getRange(2,col[0]+1,constant[0],4).getValues(),
                s1.getRange(2,col[1]+1,constant[1],4).getValues(),
                s1.getRange(1,col[2]+1,constant[2]+1,2).getValues(),
                s1.getRange(1,col[3]+2,constant[3]+1,4).getValues()]
  
  //title 1
  var title1 = body.appendParagraph(titles[0]);
  body.appendParagraph("");
  
  //body
  for(var i=0;i<4;i++){
    if(i==3){
      var title2 = body.appendParagraph(titles[1]);
      body.appendParagraph("");
    }
    body.appendTable(arrays[i]);
    if(i<4){
      body.appendPageBreak();
    }
  }
  
  //beautify
  title1.setAttributes({FONT_FAMILY:'Arial',FONT_SIZE:22,HORIZONTAL_ALIGNMENT:DocumentApp.HorizontalAlignment.CENTER});
  title2.setAttributes({FONT_FAMILY:'Arial',FONT_SIZE:22,HORIZONTAL_ALIGNMENT:DocumentApp.HorizontalAlignment.CENTER});
  
  //resize & add header
  var y = body.getTables();
  for(var t=0;t<2;t++){
    y[t+1].insertTableRow(0,y[0].getRow(1).copy())
    y[t+1].insertTableRow(0,y[0].getRow(0).copy())
    if(t==1){
      y[2].getCell(1,1).setText("下半年")
    }
  }
  
  for(var e=0;e<4;e++){
    for(var p=0;p<w[e].length;p++){
      y[e+1].setColumnWidth(p,w[e][p])
    }
  }
  
  //delete
  for(var u=0;u<3;u++){
    body.getChild(2-u).removeFromParent();
  }
  
}

function ttie(){
  var doc = DocumentApp.openByUrl("https://docs.google.com/document/d/16w2Ca6upiKbgKyC98hjjgT_AB-wjpC-hS5IKTPtkRik/edit");
  var body = doc.getBody();
  var y = body.getTables();
  
  for(var t=0;t<2;t++){
    y[t+1].insertTableRow(0,y[0].getRow(1).copy())
    y[t+1].insertTableRow(0,y[0].getRow(0).copy())
    if(t==1){
      y[2].getCell(1,1).setText("下半年")
    }
  }
}

