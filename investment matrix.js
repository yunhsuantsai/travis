// This sctipr will use loop to copy stock one by one in order to get a list of investment matrix.
// Travis Tsai @ Google App Script
// file: https://docs.google.com/spreadsheets/d/1EmN8Ezv9mdcMTfYEPaPACA518u_KTdBwJEmbckjNY8A/edit#gid=1019243330


function loop() {    // copy stockid one by one from from stock data to matrices sheet  
  var ss = SpreadsheetApp.getActive(),
      target = ss.getSheetByName('獲利矩陣')    // Matrices is here
      result = ss.getSheetByName('獲利矩陣一覽表')    // Results shows here
  var nextRow = getFirstEmptyRow('B');    // Get the Row number if Row B is empty
  var lastRow = result.getLastRow();    // Get the Row number if Row A is empty
  var distance = lastRow - nextRow;
  var data = result.getRange(nextRow, 1, distance, 1).getValues();  // Get database in this line
  for (var i = 0; i < data.length; i++) {    // Loop start from here
    target.getRange('B1').setValue(data[i][0]); 
    SpreadsheetApp.flush();
    Utilities.sleep(5000);      // wait n sec. to make sure all data are updated
    copyresult();
  }
} 

function copyresult() {    // This function is to copy a cell/range from source sheet to targeted sheet 
  var ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName('獲利矩陣')    //source sheet to copy
  var score = sheet.getRange("G4:G20").getDisplayValues();  // Define source cell or range in this line
  var stock = sheet.getRange("B1").getDisplayValues();  // get current stock name & id
  var nextRow = getFirstEmptyRow('B');
      tt = ss.getSheetByName('獲利矩陣一覽表')  //targeted sheet to paste
  var range1 = tt.getRange(nextRow, 1);
  range1.setValues(stock);
  var range2 = tt.getRange(nextRow, 2, 1,17);
  range2.setValues(transpose(score));
}

function getFirstEmptyRow(columnLetter) {  // this function will check the first empty row of the targeted sheet 
  columnLetter = columnLetter || 'A';
  var rangeA1 = columnLetter + ':' + columnLetter;
  var ss = SpreadsheetApp.getActive(),
      sheet = ss.getSheetByName('獲利矩陣一覽表')  //targeted sheet here
  var column = sheet.getRange(rangeA1);
  var values = column.getValues();  
  var ct = 0;
  while ( values[ct][0] != "" ) {
    ct++;
  }
  return (ct+1);    // +1 for compatibility with spreadsheet functions
}

function transpose(a) {    // Change vertical score into horizontal ones
  return Object.keys(a[0]).map(function (c) { return a.map(function (r) { return r[c]; }); });
}
