function onForm(e) {
  var ss1 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sfwGemyVBbN3C5PpIsMJtFYzHHbWNkZ0slaOIV5UHDU/edit#gid=2091781786");
  var ss2 = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1JbvkRR6dkZSBrs4N3pX0fZKpfdbAbsO7XuQcHK3vlEk/edit#gid=206660123");
  var s1 = ss1.getSheetByName("Form Responses 1");
  var s2 = ss2.getSheetByName("data");
  
  var form = FormApp.openByUrl("https://docs.google.com/forms/d/1-mjIblBK9DcHxuSEUHk92YlK9pW-03jfGkYPpJa7aqQ/edit");
  var response = form.getResponses();
  var items = form.getItems();
  var subnum = response.length;
  var Ids=[];
  
  var formResponse = e.response;
  var itemResponses = formResponse.getItemResponses();
  var itemnum = itemResponses.length;
  
  var editUrls = s1.getRange(2,12,subnum,1).getValues().map(function(d){ return d[0];});
  var actnum = editUrls.indexOf(formResponse.getEditResponseUrl());
  
  if(actnum<0){
    var actcode = subnum;
    var c =1
    }
  else{
    var actcode = actnum+1
    var c = 2
    }
  
  //set constant
  if(c==1){
    s1.getRange(actcode+1,1).copyTo(s1.getRange(actcode+1,10), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  }
  s1.getRange(actcode+1,11).setValue(actcode);

  //THE CHECK
  if(itemnum>6){
    var approve_array =[];
    var approve_prearray = items[10].asMultipleChoiceItem().getChoices();
    for(var z = 0;z<approve_prearray.length;z++){
      approve_array.push(approve_prearray[z].getValue());
    }
    var approve = itemResponses[itemnum-1].getResponse();
    var check =  approve_array.indexOf(approve);
    
    if(check<0){
      check = 3
    }
    else{
      check = check+1
    }
    
    //set constant
    s1.getRange(actcode+1,6).setValue(check);
    
    if(check==1){
      var range = s2.getRange("A2:A").getValues().map(function(d){ return d[0];}).indexOf("");
      if(range<0){
        var span = s2.getMaxRows()-1;
      }
      else{
        var span = range;
      }
      var vlookup = s2.getRange(2,2,span,3).getValues();
      var searchkey_array = vlookup.map(function(d){ return d[0];})
      
      //folder & name 
      var folderUrl = itemResponses[1].getResponse().trim();
      var name = itemResponses[2].getResponse();
      var type1 = itemResponses[3].getResponse();
      var type2 = itemResponses[4].getResponse();
      
      var index = searchkey_array.indexOf(folderUrl);
      if(index<0){
      }
      else{
        var docUrl = vlookup[index][1];
        var pptUrl = vlookup[index][2];
        var folderId = folderUrl.slice(39,folderUrl.length);
        var title = "活动("+actcode+") ; "+name+" ; "+type1+" ; "+type2+" ; "+approve;
        
        DriveApp.getFolderById(folderId).setName(title);
        DocumentApp.openByUrl(docUrl).setName(title);
        SlidesApp.openByUrl(pptUrl).setName(title);
      }
      
    }
    
  }
  
  //assign Edit url
  assignEditUrls(); 
}

function assignEditUrls() {
  var form = FormApp.openByUrl("https://docs.google.com/forms/d/1-mjIblBK9DcHxuSEUHk92YlK9pW-03jfGkYPpJa7aqQ/edit");
  var sheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1sfwGemyVBbN3C5PpIsMJtFYzHHbWNkZ0slaOIV5UHDU/edit#gid=2091781786").getSheetByName("Form Responses 1");

  var data = sheet.getDataRange().getValues();
  var urlCol = 12; 
  var responses = form.getResponses();
  var timestamps = [], urls = [], resultUrls = [];
  
  for (var i = 0; i < responses.length; i++) {
    timestamps.push(responses[i].getTimestamp().setMilliseconds(0));
    urls.push(responses[i].getEditResponseUrl());
  }
  for (var j = 1; j < data.length; j++) {

    resultUrls.push([data[j][0]?urls[timestamps.indexOf(data[j][0].setMilliseconds(0))]:'']);
  }
  sheet.getRange(2, urlCol, resultUrls.length).setValues(resultUrls);  
}

