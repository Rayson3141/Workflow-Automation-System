function createppt(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1355093305); //prep;
  var s2 = getSheetById(1620429395); //Gen;
  var s3 = getSheetById(411658923); //export
  var checkblank =  s1.getRange(8,1).getValue();
  
  if(checkblank==""){
    ui.alert("第几次学会活动??","请输入一个合适的学会活动", ui.ButtonSet.OK);  
  }
  else{ 
    //clear past
    s2.getRange(3,1).clearContent();
    s2.getRange(3,1).setValue(2);
    
    //copy ppt 
    var presentation = DriveApp.getFileById('1_sUK0lQpb8PjL2keesn1wq6uqPlxE0Ubc9Pref5tI5g').makeCopy();
    var presentationtUrl = presentation.getUrl();
    var presentationId = presentation.getId();
    var plantitle = s1.getRange(10,1).getValue();
    var f = DriveApp.getFolderById('1j1NEkZ_xTOzKYZ69k39d7aViO0E3rup-');
    DriveApp.getFileById(presentationId).setName(plantitle).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    
    //save
    var s3lastrow = s3.getMaxRows();
    s3.getRange(2,15,s3lastrow,4)
    .copyTo(s3.getRange(2,11,s3lastrow,4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    s3.getRange(2,11).setValue(new Date());
    s3.getRange(2,14).setValue(presentationtUrl)
    
    
    var goalppt = SlidesApp.openByUrl(presentationtUrl);
    var num = s1.getRange(1,2).getValue();
    
    s2.getRange(3,1).setValue(1);
    writeppt(s1,num,goalppt);
    
    //color code
    if(s1.getRange(9,1).getValue()==1){
      s1.getRange(9,1).setValue(2);
    }
    else{
      s1.getRange(9,1).setValue(1);
    }
    s2.getRange(3,1).setValue(presentationtUrl);
    s2.getRange(2,1).clearContent();
  }
}


function writeppt(s1,num,goalppt){
  var array = s1.getRange(1,4,num+1,6).getValues();
  var check = array.map(function(d){ return d[0];});
  
  var slides = goalppt.getSlides();
  var name = s1.getRange(3,1).getValue();
  var subfrontpage = slides[0];
  
  
  var frontpage = slides[1].getShapes();
  frontpage[0].getText().appendText(name);
  
  for(var i=1;(i)*(i-num-1)<0;i=check.indexOf(1,i)+1){
    var z =check.indexOf(1,i);
    
    if(z>-1){
      var z1 = check.indexOf(1,z+1);
      var subname = array[z][3];
      var subslide = goalppt.appendSlide(subfrontpage);
      var subshapes = subslide.getShapes();
      subshapes[0].getText().appendText(subname);
      
      if(z1>-1){
        var n = z1;
      }else{
        var n = num+1
      }
      //id prep
      for(var ii =z+1;ii<n;ii++){
        var ID = array[ii][5];
        var ppt = SlidesApp.openByUrl(ID);
        var pptslides = ppt.getSlides();        
        
        for(var l=0;l<pptslides.length;l++){
          goalppt.appendSlide(pptslides[l])
        }
      }
      
    }   
    
  }
  
  slides[0].remove();
}  

function y(){
  getSheetById(1035613208); //data 1035613208
  getSheetById(1033467041); //data0 1033467041
  getSheetById(1355093305); //prep 1355093305
  getSheetById(411658923); //export 411658923
  getSheetById(0); //活动planning 
  getSheetById(430900926); //筹主，主讲人 430900926
  getSheetById(1620429395); //Gen 1620429395
  getSheetById();
  
}
