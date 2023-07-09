// Constants for spreadsheet data sheet
const DATA_SHEET_NAME = "Custom Field Config";
const FIELD_NAME_COLUMN = 1;
const FIELD_TYPE_COLUMN = 2;
const CATEGORY_COLUMN = 3;
const OPTIONS_COLUMN = 4;
const STATUS_COLUMN = 5;
const TOTAL_COLUMNS = 5;

// Constants for config sheet
const CONFIG_SHEET_NAME = "API Config";
const DOMAIN_NAME = "B1";
const API_TOKEN_CELL = "B2";

// Get the active spreadsheet
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// Get the data sheet
var dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
// Get the config sheet
var configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

//Calling the API Configuration sheet data
var [domainName,apiToken] = getAPIConfig(configSheet);


// Function to create a custom field in Pipedrive
function processFields() {
  
  // Get all the values from Column 1 to Column N
  var range = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1,TOTAL_COLUMNS);
  var values = range.getValues();
  
  
  
  // Iterate through the each field values in the values object
  for (var i = 0; i < values.length; i++) {
    let fieldName = values[i][FIELD_NAME_COLUMN - 1];
    let fieldType = values[i][FIELD_TYPE_COLUMN - 1 ];
    let category = values[i][CATEGORY_COLUMN - 1];
    let option = values[i][OPTIONS_COLUMN - 1]; 
    let status = values[i][STATUS_COLUMN - 1];

    if(status != "1"){

      if (createCustomField(fieldName,fieldType,category,option)) 
      {
         var cell = dataSheet.getRange(i+2 , STATUS_COLUMN); // Get the specific cell in the desired column and row
         cell.setValue("1");
         cell.setFontColor("#34a853");
      }

      else
      {
         var cell = dataSheet.getRange(i+2 , STATUS_COLUMN); // Get the specific cell in the desired column and row
         cell.setValue("0");
         cell.setFontColor("#ea4335");
      }

    }
  }
}

function getAPIConfig(configSheet){
  var api_Token = configSheet.getRange(API_TOKEN_CELL).getValue();
  var domain_Name = configSheet.getRange(DOMAIN_NAME).getValue();
  return [domain_Name,api_Token];
}

function createCustomField(fieldName,fieldType,category,option){
  // Create the custom field using the Pipedrive API
      var apiUrl = `https://${domainName}.pipedrive.com/api/v1/${category}?api_token=${apiToken}`;

      if(fieldType == "set" || fieldType == "enum")
      {
        var payload = {
            name:`${fieldName}`,
            field_type:`${fieldType}`,
            options: option.split(",")
      };
      }
      
      else{
        var payload = {
            name:`${fieldName}`,
            field_type:`${fieldType}`
      };
      }

      
      var option_s = {
        "method": "POST",
        "contentType": "application/json",
        "payload": JSON.stringify(payload)
      };

      try{
        var response = UrlFetchApp.fetch(apiUrl, option_s);
        var responseData = JSON.parse(response.getContentText());

        if (responseData.success && responseData.data.id !== '') 
        {
          return true;
        }

        else
        {
          return false;
        }
      } catch (error) {
        // Code to handle the exception
        Logger.log("An error occurred: " + error.message);
      }
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Pipedrive')
    .addItem('Create Fields', 'processFields')
    .addToUi();
}
