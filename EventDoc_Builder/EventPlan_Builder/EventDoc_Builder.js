function createdoc(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var s1 = getSheetById(1355093305); //prep
  var s2 = getSheetById(1620429395); //Gen
  var s3 = getSheetById(411658923); //export
  var s4 = getSheetById(1035613208); //data;
  var checkblank =  s1.getRange(8,1).getValue();
  
  if(checkblank==""){
    ui.alert("第几次学会活动??","请输入一个合适的学会活动", ui.ButtonSet.OK);  
  }
  else{ 
    //clear past
    s2.getRange(3,1).clearContent();
    s2.getRange(3,1).setValue(4);
    
    //copy doc
    var document = DriveApp.getFileById('1l0jnwdJCFzyRpCUJajrbVcD1-rk45q-aYLcxTeK2CTg').makeCopy();
    var documentUrl = document.getUrl();
    var documentId = document.getId();
    var plantitle = s1.getRange(6,1).getValue();
    var f = DriveApp.getFolderById('1NnCbRSiHLPcpxj80nVblzrkJM_X7IqKx');
    DriveApp.getFileById(documentId).setName(plantitle).moveTo(f).setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.VIEW);
    
    //save
    var s3lastrow = s3.getMaxRows();
    s3.getRange(2,6,s3lastrow,4)
    .copyTo(s3.getRange(2,2,s3lastrow,4), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    s3.getRange(2,2).setValue(new Date());
    s3.getRange(2,5).setValue(documentUrl)
    
    
    var doc = DocumentApp.openByUrl(documentUrl);
    var body = doc.getBody();
    var num_group = s1.getRange(1,1).getValue();
    var num = s1.getRange(1,2).getValue();
    
    tables(s1,s2,body,num,num_group);
    s2.getRange(3,1).setValue(3);
    bodies(s1,s2,body,num,num_group);
    s2.getRange(3,1).setValue(2);
    images(s1,s2,body,num,num_group);
    s2.getRange(3,1).setValue(1);
    flow(s1,s2,body,num,num_group);
    frontpage(s1,s2,body,num,num_group);
    
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
                                                 .requireValueInRange(s4.getRange('$N$2:$N'), true)
                                                 .build());
  }
}


function images(s1,s2,body,num,num_group){
  var y = body.getTables();

  //create merge rows
  var df = y[0].copy();
  df.getCell(0,0).removeFromParent();
  for(var qr=0;qr<9;qr++){
    df.getCell(0,1).removeFromParent();
  }
  
  var p1 = [s1.getRange(2,1).getValue(),s1.getRange(2,2).getValue()];
  var p2 =  s1.getRange(1,8,2,num_group+1).getValues();
  var array = s1.getRange(1,4,num+1,5).getValues();
  var check = array.map(function(d){ return d[0];});
  
  for(var i1= 0;i1<2;i1++){
    var c = 0
    for(var i=1;(i)*(i-num-1)<0;i=check.indexOf(1,i)+1){
      var z =check.indexOf(1,i);
      
      if(z>-1){
        var z1 = check.indexOf(1,z+1);
        
        if(z1>-1){
          var n = z1;
        }else{
          var n = num+1
        }
        //id prep
        for(var ii =z+1;ii<n;ii++){
          var ID = array[n+z-ii][4];
          var doc_prep = DocumentApp.openByUrl(ID);
          var body_prep = doc_prep.getBody();
          var tables_prep = body_prep.getTables();
          var table_prep = [tables_prep[1],tables_prep[3]]
          var table = table_prep[1-i1];
          
          if(table.findElement(DocumentApp.ElementType.INLINE_IMAGE)!==null){
            var numrows = table.getNumRows();
            var title = array[n+z-ii][3];
            var insert_table_num = body.getChildIndex(y[(1-i1)*num_group+c+2])
            var tablef = body.insertTable(insert_table_num+1,[[title]]);
            for(var g=0;g<numrows;g++){
              var copyrow = table.getRow(g);
              tablef.appendTableRow(copyrow.copy());
            }
            
            body.insertPageBreak(insert_table_num+1);
          }
        }
        
        
      } 
      
      c=c+1      
    }
    
  }
  
}



function frontpage(s1,s2,body,num,num_group){
  var y = body.getTables();
  
  //create merge rows
  var df = y[0].getRow(0).copy();
  for(var qr=0;qr<11-(num_group+1);qr++){
    df.removeCell(1+num_group);
  }
  
  var toptitle = s1.getRange(2,1).getValue();
  var p1 = [s1.getRange(4,1).getValue(),s1.getRange(4,2).getValue()];
  var p2 =  s1.getRange(1,10,2,num_group+1).getValues();
  
  body.getChild(0).asParagraph().setText(toptitle);
  body.getChild(0).asParagraph().setAttributes({FONT_FAMILY:'Arial',FONT_SIZE:22})
  var table = body.insertTable(2,p2);
  var p1row = table.insertTableRow(1,df.copy()).getCell(0).setText(p1[0]);
  table.getRow(1).getCell(1).setText(p1[1])
  body.removeChild(y[0]);
  body.insertParagraph(1,"")
}



function bodies(s1,s2,body,num,num_group){
  var y = body.getTables();

  //create merge rows
  var df = y[0].getRow(0).copy();
  df.setAttributes({BORDER_WIDTH:1.0});
  var df1 = y[0].getRow(0).copy();
  var df2 = y[0].getRow(0).copy();
  df.removeCell(0);
  df1.removeCell(0);
  df2.removeCell(0);
  for(var qw=0;qw<7;qw++){
    df.removeCell(3);
  }
  for(var qe=0;qe<8;qe++){
    df1.removeCell(2);
  }
  for(var qr=0;qr<9;qr++){
    df2.removeCell(1);
  }
  var m_row = [df2,df1,df]
  
  var p1 = [s1.getRange(2,1).getValue(),s1.getRange(2,2).getValue()];
  var p2 =  s1.getRange(1,8,2,num_group+1).getValues();
  var array = s1.getRange(1,4,num+1,5).getValues();
  var check = array.map(function(d){ return d[0];});

  var c = 0
  
  for(var i=1;(i)*(i-num-1)<0;i=check.indexOf(1,i)+1){
    var z =check.indexOf(1,i);
    
    if(z>-1){
      var z1 = check.indexOf(1,z+1);
      var num_per_group = array[z][2]
      
      if(z1>-1){
        var n = z1;
      }else{
        var n = num+1
      }
      //id prep
      for(var ii =z+1;ii<n;ii++){
        var ID = array[ii][4];
        var doc_prep = DocumentApp.openByUrl(ID);
        var body_prep = doc_prep.getBody();
        var tables_prep = body_prep.getTables();
        var table_prep = [tables_prep[0],tables_prep[2],tables_prep[4]]
        var title = array[ii][3];
        var num_per_act = array[ii][2]
        
        //edit body[1,2,3]
        for(var i1=0;i1<3;i1++){
          var table = table_prep[2-i1];
          var num_r_table = table.getNumRows();
          var insert_table_num = (num_group)*(2-i1)+c+2
          
          ///body[3]
          if(2-i1== 2){
            y[insert_table_num].appendTableRow(m_row[2].copy()).getCell(0).setText(title+" （活动数量= "+num_per_act+" ）");
            
            for(var i2=1;i2<num_r_table;i2++){
              var copy_row = table.getRow(i2);
              var insert_row = y[insert_table_num].appendTableRow(copy_row.copy());
              var amount = copy_row.getCell(1).getText();
              insert_row.appendTableCell(amount*num_per_act+"")
            }
            
          }
          
          //body[1]
          if(2-i1==0){
            y[insert_table_num].appendTableRow(m_row[1].copy()).getCell(0).setText(title);            
            
            for(var i2=1;i2<num_r_table;i2++){
              var copy_row = table.getRow(i2);
              var insert_row = y[insert_table_num].appendTableRow(copy_row.copy());
            }
          }
          
          //body[2]
          if(2-i1==1){
            y[insert_table_num].appendTableRow(m_row[0].copy()).getCell(0).setText(title);            
            
            for(var i2=0;i2<num_r_table;i2++){
              var copy_row = table.getRow(i2);
              var insert_row = y[insert_table_num].appendTableRow(copy_row.copy());
            }
            if(array[ii][0]==2){
              y[insert_table_num].removeRow(0);
            }
            
          }
          
//          y[insert_table_num].setBorderWidth(1);
          
        }
        
      }   
      
      
      c=c+1
    }
  }
}



function tables(s1,s2,body,num,num_group){
  var y = body.getTables();
  var child = [8,11,14];
  
  var array = s1.getRange(1,4,num+1,4).getValues();
  var group = array.filter(function(d){ return d[0]===1});
  
  var samp_table = y.slice(2,5);
  
  for(var i=0;i<3;i++){
    var c = 0
    var table_c = samp_table[2-i]
    var index = child[2-i];
    
    if(2-i==2){
      for(var f=0;f<num_group;f++){
        var title = group[f][1];
        var num_act = group[f][2];
        body.insertParagraph(index+4*c,"");
        body.insertParagraph(index+1+4*c,title+" （活动数量= "+num_act+" ）");
        body.insertPageBreak(index+1+4*c);
        var table = body.insertTable(index+3+4*c,table_c.copy());
        c=c+1
      }
    }else{
      for(var f=0;f<num_group;f++){
        var title = group[f][1]
        body.insertParagraph(index+4*c,"");
        body.insertParagraph(index+1+4*c,title);
        body.insertPageBreak(index+1+4*c);
        var table = body.insertTable(index+3+4*c,table_c.copy());
        c=c+1
      }
    }
//    table_c.clear();
    body.removeChild(body.getChild(index+1));
    body.removeChild(body.getChild(index));
    body.removeChild(body.getChild(index-1));
//    body.getChild(index).removeFromParent();
  }
}



function flow(s1,s2,body,num,num_group){
  var y = body.getTables();
  
  var p1 = [s1.getRange(2,1).getValue(),s1.getRange(2,2).getValue()];
  var p2 =  s1.getRange(1,8,2,num_group+1).getValues();
  var array = s1.getRange(1,4,num+1,4).getValues();
  var check = array.map(function(d){ return d[0];});
  var samp_table = y[1].copy();
  var samp_row = y[1].getRow(1).copy();
  var c = 0

  for(var i=1;(i)*(i-num-1)<0;i=check.indexOf(1,i)+1){
    var z =check.indexOf(1,i);
    
    if(z>-1){
      var z1 = check.indexOf(1,z+1);
      var title = array[z][1];
      
      body.insertParagraph(5+3*c,"");
      body.insertParagraph(6+3*c,title);
      var table = body.insertTable(7+3*c,samp_table.copy());
      c=c+1
      
      if(z1>-1){
        var n = z1;
      }else{
        var n = num+1
      }
      
      for(var ii =z+1;ii<n;ii++){
        var val = array[n+z-ii][1];
        var i_row1 = table.insertTableRow(2,samp_row.copy());
        var i_row = table.insertTableRow(2,samp_row.copy());
        i_row1.getCell(0).setText("");
        i_row1.getCell(1).setText("进行活动（"+val+"）");
        i_row.getCell(0).setText("");
        i_row.getCell(1).setText("主讲人解释活动（"+val+"），组长准备材料。");
      }   
    }
  }
  
  body.removeChild(body.getChild(5));
  body.removeChild(y[1])
  
}