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
const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
// Get the data sheet
const dataSheet = spreadsheet.getSheetByName(DATA_SHEET_NAME);
// Get the config sheet
const configSheet = spreadsheet.getSheetByName(CONFIG_SHEET_NAME);

// Calling the API Configuration sheet data
const [domainName, apiToken] = getAPIConfig(configSheet);

// Function to create a custom field in Pipedrive
function processFields() {
  // Get all the values from Column 1 to Column N
  const range = dataSheet.getRange(2, 1, dataSheet.getLastRow() - 1, TOTAL_COLUMNS);
  const values = range.getValues();

  // Iterate through each field values in the values object
  values.forEach((row, i) => {
    const [fieldName, fieldType, category, option, status] = row;

    if (status !== "1") {
      if (createCustomField(fieldName, fieldType, category, option)) {
        setCellStatus(i + 2, "1", "#34a853");
      } else {
        setCellStatus(i + 2, "0", "#ea4335");
      }
    }
  });
}

function getAPIConfig(configSheet) {
  const apiToken = configSheet.getRange(API_TOKEN_CELL).getValue();
  const domainName = configSheet.getRange(DOMAIN_NAME).getValue();
  return [domainName, apiToken];
}

function createCustomField(fieldName, fieldType, category, option) {
  // Create the custom field using the Pipedrive API
  const apiUrl = `https://${domainName}.pipedrive.com/api/v1/${category}?api_token=${apiToken}`;

  let payload = {
    name: fieldName,
    field_type: fieldType
  };

  if (fieldType === "set" || fieldType === "enum") {
    payload.options = option.split(",");
  }

  const options = {
    method: "POST",
    contentType: "application/json",
    payload: JSON.stringify(payload)
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, options);
    const responseData = JSON.parse(response.getContentText());

    if (responseData.success && responseData.data.id !== '') {
      return true;
    } else {
      return false;
    }
  } catch (error) {
    // Code to handle the exception
    Logger.log("An error occurred: " + error.message);
    return false;
  }
}

function setCellStatus(row, status, fontColor) {
  const cell = dataSheet.getRange(row, STATUS_COLUMN);
  cell.setValue(status).setFontColor(fontColor);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Pipedrive')
    .addItem('Create Fields', 'processFields')
    .addToUi();
}

