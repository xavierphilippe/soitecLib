/**
* create a log - type Error
* result : a new row in the target spreadsheet : Date / Type / ActiveUser /  Message
*
*<pre>
* Usage example : writeLogSpreadsheet('Error', 'function MyTest, wrong parameter', 'mylog', 'EeIzKI_SazP-RfbqBpH');
*  or with default value : writeLogSpreadsheet ('Warning','missing value in MyFunction'); ==> result message is inserted in 'log' sheet
*</pre>
* @param  {string} type expected values : 'Warning' or 'Error' or 'Info'
* @param  {string} message content to be recorded into sheet
* @param {string} optSheetName name of the sheet that contains the log 
* @param {string} optSpreadsheetId ID of the spreadsheet (optional - default value : current spreadsheet)
*/
function logWriteLogSpreadsheet(type, message, optSheetName, optSpreadsheetId) {    
  
  //test type parameters
  if (((type != 'Warning')&&(type != 'Error')&&(type != 'Info'))){
    Browser.msgBox ('Wrong parameter in writeLogSpreadsheet function - type:expected values : Warning or Error or Info');
    return;
  }
  //manage optional parameters
  var sheet;
  switch (arguments.length - 2) {case 0:  optSheetName = 'log'; case 1:  optSpreadsheetId = -1; }
  if (optSpreadsheetId == -1) {
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(optSheetName);
  } else {
    sheet = SpreadsheetApp.openById(optSpreadsheetId).getSheetByName(optSheetName);
  }
  
  //if sheet doesn't exist - creation of 'log' spreadsheet
  if (sheet == null){ 
    if (optSpreadsheetId == -1) {
      sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet('log');
    } else {
      sheet = SpreadsheetApp.openById(optSpreadsheetId).insertSheet('log');;
    }
  }
  
  // create a new row + write the message 
  sheet.appendRow([new  Date(), type, Session.getActiveUser(), message]);
}


/**
* Private function
* --------------------------
* create a system log for soitec administrator
* then spreadsheetID and sheetName are defined by soitec Administrator
* tis function is reserved for function created by Soitec
* <pre>
* example : LoggingAddLogErrorSpreadsheet('log','Error in function MyTest, wrong parameter', 'MyFunction')
* </pref>
*
* @param  {string} type expected values : 'Warning' or 'Error' or 'Info'
* @param  {string} message content to be recorded into sheet
* @param {string} optFunction Name of the function executed (Optional - default value : Empty)
*/
function logWriteLogSystemSpreadsheet_(type,message,optFunction) {  
  var spreadsheetId = '1q3OaCUmhDAt5HBCDOEBxaz_nh9p1tRWNrSOaAQ4pyqQ';
  var sheetName = 'log';
  var activeSpreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  switch (arguments.length - 2) {case 0:  optFunction = '';  }  
  sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);  
  sheet.appendRow([new Date(), type, Session.getActiveUser(), Session.getEffectiveUser(), activeSpreadsheetUrl, optFunction, message]);
}
