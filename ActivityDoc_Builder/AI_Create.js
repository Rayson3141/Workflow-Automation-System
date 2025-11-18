function AI_Create(prompt,document){
  // var s = SpreadsheetApp.getActiveSpreadsheet();
  var doc = DocumentApp.openById(document)
  // var s1 = s.getSheetByName("Sheet5"); //data
  

  // var prompt = s1.getRange(1,4).getValue();
  var prompt1 = "请通过此'"+prompt+"'，打出该活动的步骤。请使用中文打出来。"
  var prompt2 = "请通过此'"+prompt+"'，打出该活动的注意事项。请使用中文打出来与“*”列出来。"
  var prompt3 = "请通过此'"+prompt+"'，打出该活动的理想结果。请使用中文打出来。"
  var prompt4 = "请通过此'"+prompt+"',用这四列材料、数量、单位、备注生成一个材料表。请使用中文打出来。"
  var prompt5 = "请通过此'"+prompt+"'，列出该实验的科学原理的关键词。请使用中文打出来。"
  var prompt6 = "请通过此'"+prompt+"'，打出该实验的科学原理。请使用中文打出来。"


  // 活动内容
  TableInput(doc,removeNumberingArray(GemPrompt(prompt1).split("\n")),0)
  TableInput(doc,removeBulletArray(GemPrompt(prompt2).split("\n")),1)

  //理想结果
  TableTextInput(doc,2,0,GemPrompt(prompt3))

  //材料表
  TableListInput(doc,ArrayFy(GemPrompt(prompt4)))

  //原理
  TableTextInput(doc,5,10,GemPrompt(prompt5),true)
  TableTextInput(doc,5,11,GemPrompt(prompt6))

  // x = GemPrompt(prompt2).split("\n")
  // Logger.log(GemPrompt(prompt3))

}


function GemPrompt(reviewtext){
  var apiKey = "AIzaSyBEa9ipg6WtJUeJ_VxIUpGPGykl8OXLOec";
  var apiURL = "https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent"
  var url = apiURL +"?key="+apiKey;

  var headers = {
    "Content-Type": "application/json"

  };

  var requestBody = {
    "contents": [
        {
            "parts": [
                {
                    "text": reviewtext
                }
            ]
        }
    ]
  };

  
  var options = {

    "method": "POST",
    "headers": headers,
    "payload": JSON.stringify(requestBody)
  };

  var response = UrlFetchApp.fetch(url,options);
  var data = JSON.parse(response.getContentText());
  var output = data.candidates[0].content.parts[0].text

  return output
}

function TableListInput(doc,inputArray) {
  var tables = doc.getBody().getTables()
  var MaterialTable = tables[tables.length-2]
  const row = MaterialTable.getChild(1).copy();


  // Append a new row for each row in the matrix.
  for (var i = 0; i < inputArray.length; i++) {
    const newRow = row.copy();
    MaterialTable.appendTableRow(newRow);
    LastRow = MaterialTable.getNumRows()-1
    // row_edit = MaterialTable.getChild(MaterialTable.getNumChildren-1)

    // Append a new cell for each column in the matrix.
    for (var j = 0; j < inputArray[i].length; j++) {
      MaterialTable.getChild(LastRow).getCell(j).setText(inputArray[i][j]);
    }
  }
  MaterialTable.removeRow(1)
}

function TableInput(doc,inputArray,cellnum) {
  // Get the document and the table
  var tables = doc.getBody().getTables();
  var table = tables[0]

  // Get the target cell
  var cell = table.getChild(1).getCell(cellnum)

  // Split the input string into an array of items
  var itemList = inputArray;
  var length = itemList.length

  // Iterate through the items and insert them as a numbered list
  for (var i = 0; i < length; i++) {
    // Insert the item as a numbered list item
    var listItem = cell.getChild(0).copy()
    listItem.setText(itemList[length-i-1]);
    cell.insertListItem(0,listItem) 
  }

}

function TableTextInput(doc,tablenum,row,text,connect=false){
  var tables = doc.getBody().getTables()
  var PripTable = tables[tablenum]
  if (connect){
   var k = text.split("\n")
   text = k.join(",")
  }
  PripTable.getChild(row).getCell(0).editAsText().appendText(text)
}


function removeNumberingArray(Array) {
  for (var i = 0; i < Array.length; i++) {
    Array[i]=removeNumbering(Array[i])
  }
  return Array.filter(array => array.length > 0);
}

function removeBulletArray(Array) {
  for (var i = 0; i < Array.length; i++) {
    Array[i]=removeBullet(Array[i])
  }
  return Array.filter(array => array.length > 0);
}

function removeNumbering(text) {
  // Replace all of the numbers with an empty string.
  var x =text.split(".")
  x.shift()

  // Return the result.
  return x.join(".").trim()
}

function removeBullet(text) {
  // Replace all of the numbers with an empty string.
  var x =text.split("*")
  // x.shift()

  // Return the result.
  return x.join("").trim()
}

function ArrayFy(result){
  const matrix = []

  for(const row of result.split("\n")){
    x = row.split("|")
    if(x.length==6){
      x.shift()
      x.pop()
      matrix.push(trimArrayElements(x))
    }
  }
  matrix.shift()
  matrix.shift()

  return matrix
}

function withRetry(operation, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      return operation; // Execute the operation
    } 
    catch (error) {
      console.error(`Error in attempt ${attempt}: ${error.message}`);
      // Optional: Implement additional error handling or logging here
    }
    }
  throw new Error("Failed operation after all retries");
}

function trimArrayElements(array) {
  // Use map() to create a new array with trimmed elements
  return array.map(element => element.trim());
}


