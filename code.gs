// Add the menu to the Spreadsheet interface
function onOpen(e) {
  SpreadsheetApp.getUi()
      .createMenu('Block Urls')
      .addItem('Deploy', 'getOUFromUser')
      .addToUi();
}

//Promt user for list and OU target
function getOUFromUser() {
  const htmlOutput = HtmlService
                       .createHtmlOutput(
                           '<p>Vælg liste og<br>hvilket OU, blokeringen skal gælde for</p><label for="shNames">Vælg Liste:</label><select name="shNames" id="shNames">'+getAllSheetNamesAsOptions()+'</select><br><br><label for="orgunits">Vælg OU:</label><select name="orgunits" id="orgunits">'+getOrgUnitIds()+'</select><br><br><button onclick="google.script.run.batchModifyUrlBlock(document.getElementById(\'shNames\').value,document.getElementById(\'orgunits\').value);google.script.host.close()">Deploy</button>'
                           )
                       .setWidth(400)
                       .setHeight(200);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Vælg blokering og OU');
}


// Deploy URLBlocklist
function batchModifyUrlBlock(sheetName, orgUnitId) {
  var orgUnitIds = [orgUnitId]; // Replace with your org unit IDs.
  
  //var sheetName = "Samlet"; // Replace with your sheet name
  var columnName = "URL"; // Replace with your column name
  var ui = SpreadsheetApp.getUi();

  var blockedUrls = getColumnValuesAsArray(sheetName, columnName); // Replace with your blocked URLs.

  var result = batchModifyUrlBlockPolicy(orgUnitIds, blockedUrls);
  if(result){
    ui.alert("batchModifyUrlBlockPolicy ran successfully, check logs for details");
  } else {
    ui.alert("batchModifyUrlBlockPolicy failed, check logs for details");
  }
}

//GET  STUFF FROM SHEETS
//Get the list of URLs to block from the column with the header "URL"
function getColumnValuesAsArray(sheetName, columnName) {
  try {
    // Get the spreadsheet and sheet
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log("Sheet '" + sheetName + "' not found.");
      return null;
    }

    // Get the header row to find the column index
    var headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var columnIndex = headerRow.indexOf(columnName) + 1; // +1 because index is 1-based

    if (columnIndex === 0) {
      Logger.log("Column '" + columnName + "' not found in header row.");
      return null;
    }

    // Get the last row with data in the specified column
    var lastRow = sheet.getLastRow();

    // Get the values from the specified column
    var columnValues = sheet.getRange(2, columnIndex, lastRow - 1, 1).getValues(); // Start from row 2 (after header)

    // Flatten the 2D array to a 1D array
    var resultArray = columnValues.map(function(row) {
      return row[0];
    });

    return resultArray;

  } catch (e) {
    Logger.log("Error: " + e.toString());
    return null;
  }
}

//Get all sheet names as a HTML list of options
function getAllSheetNamesAsOptions() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ss.getSheets();
  var optionsHTML = "";

  for (var i = 0; i < sheets.length; i++) {
    var shName = sheets[i].getName();
    optionsHTML+='<option value="'+shName+'">'+shName+'</option>';
  }

  return optionsHTML;
}

//Get all orgunits and ids as a html list of options
function getOrgUnitIds() {
  try {
    var optionsHTML = "";
    var pageToken;
    var orgUnits;
    do {
      orgUnits = AdminDirectory.Orgunits.list('my_customer',{
        pageToken: pageToken,
        type: 'all' // include sub org units.
      });

      if (orgUnits.organizationUnits) {
        orgUnits.organizationUnits.sort((a, b) => {
          // Split the paths into components
          const pathA = a.orgUnitPath.split("/");
          const pathB = b.orgUnitPath.split("/");

          // Compare paths component by component
          for (let i = 0; i < Math.min(pathA.length, pathB.length); i++) {
            if (pathA[i] < pathB[i]) return -1; 
            if (pathA[i] > pathB[i]) return 1; 
          }

          // If all components match, sort by name (case-insensitive)
          if (pathA.length === pathB.length) {
            return a.name.toLowerCase().localeCompare(b.name.toLowerCase());
          }

          // If all components match and names are equal, shorter path comes first (unlikely)
          return pathA.length - pathB.length;
        });
        for (var i = 0; i < orgUnits.organizationUnits.length; i++) {
          var orgUnit = orgUnits.organizationUnits[i];
          optionsHTML+='<option value="'+orgUnit.orgUnitId.substring(3)+'">'+orgUnit.orgUnitPath+'</option>';
        }
        pageToken = orgUnits.nextPageToken;
      } else {
        break; // No more org units to check
      }
    } while (pageToken);

    Logger.log(optionsHTML);
    //Logger.log('Organizational unit with path "' + orgUnitPath + '" not found.');
    return optionsHTML;
  } catch (e) {
    Logger.log('Error: ' + e.toString());
    return null;
  }
}

//Deploy the list and change options in Google Admin
function batchModifyUrlBlockPolicy(orgUnitIds, blockedUrls) {
  try {
    // Replace 'my_customer' with your actual customer ID or keep 'customers/my_customer'.
    var customerId = 'customers/my_customer';

    var requests = orgUnitIds.map(function(orgUnitId) {
      return {
        policyTargetKey: {
          targetResource: 'orgunits/' + orgUnitId
        },
        policyValue: {
          policySchema: 'chrome.users.UrlBlocking',
          value: {
            urlBlocklist: blockedUrls
          }
        },
        updateMask: 'urlBlocklist'
      };
    });

    var payload = {
      requests: requests
    };

    var options = {
      method: 'POST',
      contentType: 'application/json',
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    var url = 'https://chromepolicy.googleapis.com/v1/' + customerId + '/policies/orgunits:batchModify';
    var response = UrlFetchApp.fetch(url, options);
    var responseCode = response.getResponseCode();
    var responseBody = response.getContentText();

    if (responseCode >= 200 && responseCode < 300) {
      Logger.log('Batch modify successful: ' + responseBody);
      return JSON.parse(responseBody);
    } else {
      Logger.log('Batch modify failed. Response code: ' + responseCode + ', Response body: ' + responseBody);
      return null;
    }

  } catch (e) {
    Logger.log('Error: ' + e.toString());
    return null;
  }
}


//Not in use: Copy from Aarhus Municipality/Søren Torp
function checkWebsiteStatus() {
  let url = "<WEBSITE_URL>";

  // Record time so we can track how long the website
  // takes to load.
  let start = new Date();
  let response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  let end = new Date();

  let responseCode = response.getResponseCode();
  let loadTimeMs = end - start;

  // Record a log of the website's status to the spreadsheet.
  SpreadsheetApp.getActive().getSheetByName("Data").appendRow([start, responseCode, loadTimeMs]);

  // Send email notification if 
  if(response.getResponseCode() != 200) {
    let email = "<EMAIL_ADDRESS>";
    let subject = "[ACTION REQUIRED] Website may be down - " + new Date();
    let body = `The URL ${url} may be down. Expected response code 200 but got ${responseCode} instead.`;
    MailApp.sendEmail(email, subject,body);
  }
}
