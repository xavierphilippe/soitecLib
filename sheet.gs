/**
* function to import a data range into a spreadsheet
*
*
* @param  {object} data a 2 dimensions array data[][]
* @param {string} sheetName name of the sheet 
* @param {string} optSpreadsheetId ID of the spreadsheet (optional - default value : current spreadsheet)
*/

function sheetImportDataToSpreadsheet(data,sheetName, optSpreadsheetId){
  
  //manage optionnal parameter and create sheet object
  var sheet;
  switch (arguments.length - 2) {case 0:  optSpreadsheetId = -1; }
  if (optSpreadsheetId == -1) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  } else {
    sheet = SpreadsheetApp.openById(optSpreadsheetId).getSheetByName(sheetName);
  }
  
  //logging System Start - log private
  logWriteLogSystemSpreadsheet_( 'Info', 'Start', 'sheetImportDataToSpreadsheet');
  
  try{    
    var LastRow = sheet.getLastRow();
    sheet.getRange(
      LastRow+1, /* first row */
      1,  /* first column */
      data.length, /* rows */
      data[0].length /* columns */
    ).setValues(data);   
  }
  catch (e){
    Browser.msgBox('Error in function sheetImportDataToSpreadsheet : ' + e.message)
    logWriteLogSystemSpreadsheet_('Error', 'function importDataToSpreadsheet - ' + e.message, 'sheetImportDataToSpreadsheet');
  }
  SpreadsheetApp.flush();
  logWriteLogSystemSpreadsheet_( 'Info',  'End', 'sheetImportDataToSpreadsheet' );
}






/**
* this function returns the value of the parameter defined in setup Sheet<br/>
* parameter name is stored in column A and value is stored in column B
*
* <pre>
* Usage example : 
*
* var columnEntity = findColIndex('Sheet1','Entity Name');
* </pre>
* 
* @param {string} colname the value of the volname
* @param  {string} sheetname the name of the sheet
* @param  {string} optSpreadsheetId Id of the spreadsheet, if not defined, the Active Spreadsheet will be consider 
* @return {integer} column index,  or  -1 if value not found
*/
function sheetGetColumnIndexByColName(colname,sheetName,optSpreadsheetId) {
  logWriteLogSystemSpreadsheet_( 'Info',  'Start', 'sheetGetColumnIndexByColName' );
  //manage optionnal parameter and create sheet object
  switch (arguments.length - 2) {case 0:  optSpreadsheetId = -1; }
  var spreadsheet;  
  if (optSpreadsheetId == -1) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    spreadsheet = SpreadsheetApp.openById(optSpreadsheetId);
  }

  var data = spreadsheet.getSheetByName(sheetName).getRange('1:1').getValues();
  var getColumnIndexByColNameValue = -1; // default value in case of Error

  for (var i = 0; i < data[0].length; i++ ) {
    if(data[0][i]==colname) {
      getColumnIndexByColNameValue = i+1;
      break;
    }
  }
  if (getColumnIndexByColNameValue == -1){
    logWriteLogSystemSpreadsheet_( 'Warning',  'Column index not found for header = ' + colname, 'sheetGetColumnIndexByColName' );
  }
  return getColumnIndexByColNameValue;
  logWriteLogSystemSpreadsheet_( 'Info',  'End', 'sheetGetColumnIndexByColName' );
}


/**
* this function returns the number of the last row (based on standard GetLastRow columns, which doesn't work fine
*
* <pre>
* Usage example : 
*
* usage : var n = sheetGetLastRow('Sheet1');
* </pre>
* 
* @param  {string} sheetname the name of the sheet
* @param  {string} optSpreadsheetId Id of the spreadsheet, if not defined, the Active Spreadsheet will be consider 
* @return {integer} row index
*/
function sheetGetLastRow(sheetname,optSpreadsheetId){
  logWriteLogSystemSpreadsheet_( 'Info',  'Start', 'sheetGetLastRow' );
  
  //manage optionnal parameter and create sheet object
  switch (arguments.length - 1) {case 0:  optSpreadsheetId = -1; }
  var spreadsheet;  
  if (optSpreadsheetId == -1) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  } else {
    spreadsheet = SpreadsheetApp.openById(optSpreadsheetId);
  }
  
  try{
    var sheet = spreadsheet.getSheetByName(sheetname);
    var lastrow = sheet.getLastRow();
    var getLastRowValue = 0;
    for(i=lastrow;i>0;i--){
      if ((sheet.getRange(i,1).getValue() + sheet.getRange(i,2).getValue()) !=''){
        getLastRowValue = i;
        break;
      }
    }
    return getLastRowValue;  
  }
  catch(e)
  {
    Browser.msgBox('ERROR - spreadsheetname' + sheetname + '-' + e.message);
    logWriteLogSystemSpreadsheet_('Error', e.message, 'sheetGetLastRow');
  }
  logWriteLogSystemSpreadsheet_( 'Info',  'End', 'sheetGetLastRow' );
}
