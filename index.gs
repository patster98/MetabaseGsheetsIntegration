function importMBData(){

  /////////////////////////////////VARIABLES////////////////////////////////////
  // Declare general variables
  var workingSheet = "Your sheet name"; // Sheet where you want to paste/retrieve data
  var inputCol = "A"; // Column where the input value for the query is stored
  var validCol = "B"; // Column for validation (while this column is empty code will iterate)
  var fillCol1 = "C"; // Column 1 to fill values from querry (adjust to the number of columns your querry returns)
  var fillCol2 = "D"; // Column 2 to fill values from querry (adjust to the number of columns your querry returns)
  var fillCol3 = "E"; // Column 3 to fill values from querry (adjust to the number of columns your querry returns)
  // Continue adding these lines if you need more columns/ data to paste depending on the data you retrieve from your query
  // Replace the letters with the actual letters from the columns you are working with

  // Declare MB credentials
  var usuarioMB = "yourUsername@yourMail.com"; // Metabase user credentials
  var pswdMB = "yourPassword"; // Metabase user pswd
  var queryNumber = "YourQueryNumber"; // Metabase query number (you can find it in the query URL)
  /////////////////////////////////////////////////////////////////////////////

  // Obtengo el token de MB
  var baseUrl = "https://yourCompanyURL.metabaseapp.com/" // Base Url where your querry is stored
  var sessionUrl = baseUrl + "api/session"; // Url defined by MB API
  var options = {
    "method": "post",
    "headers": {
      "Content-Type": "application/json"
    },
    "payload": JSON.stringify({
      username: usuarioMB,
      password: pswdMB
    })
  };
  var response;
  try {
    response = UrlFetchApp.fetch(sessionUrl, options);
  } catch (e) {
    throw (e);
  }
  var token = JSON.parse(response).id;


  // Search data input for the query
  var sheet =  SpreadsheetApp.getActive().getSheetByName(workingSheet)
  var i = 2 // First line is a header, begin iteration for every line in your spreadSheet
  var request_id = sheet.getRange(inputCol + i).getValue() // The input for your query
  
  
  // Send query to MB and iterate while input data column isn't empty
  while(request_id != ""){
    if(sheet.getRange("L" + i).getValue() == ""){

      // Get these parameters from Network tab --> Payload on Chrome DevTools when running query
      var parameter = '[{"type": "category","value": "' + request_id + '","target": ["variable",["template-tag","external_request_id"]],"id": "fbe92907-56f1-ee91-60e1-6b558fdbc427"}]';
      var encoded = encodeURIComponent(parameter); // Encode parameters to append to url format
      var questionUrl = baseUrl + "api/card/" + queryNumber + "/query/json?parameters=" + encoded + "/json"; // Url format when passing an input to the query

      // POST method options as defined by MB API, can also be retrieved from Network tab
      var options = {
        "method": "post",
        "headers": {
          "Content-Type": "application/json",
          "X-Metabase-Session": token
        },
        "muteHttpExceptions": true
      };

      var response;
      try {
        response = UrlFetchApp.fetch(questionUrl, options); // Get response from query via POST to MB API
      } catch (e) {
        return {
          'success': false,
          'error': e
        };
      }

      var statusCode = response.getResponseCode();

      if (statusCode == 200 || statusCode == 202) {

        // Check that status code isn't an error and paste values on desired sheet cells
        var values = JSON.parse(response.getContentText()); // Convert JSON format
        //console.log(response.getContentText())
        //console.log(values[0].value1Name)
        
        // Here the "value1Name" and all these values should be named exactly as the name in your query's column 
        sheet.getRange(validCol + i).setValue(values[0].value1Name) // Get desired values from JSON eg: values[arrayPositionWhereTheValueIsStored].desiredValue1
        sheet.getRange(fillCol1 + i).setValue(values[0].value2Name) // Get desired values from JSON eg: values[arrayPositionWhereTheValueIsStored].desiredValue2
        sheet.getRange(fillCol2 + i).setValue(values[0].value3Name) // idem
        sheet.getRange(fillCol3 + i).setValue(values[0].value4Name) // idem
        //Continue adding this same line with the column letter and value you want to fill from your querry
        // eg: For pasting the account_id column's value from your query on the "G" column on your sheet use:
        //  sheet.getRange("G" + i).setValue(values[0].account_id) 
        
      } else {
        console.log("Error when retrieving MB data")
      }

    }

    // Continue iteration to next line
    i++
    request_id = sheet.getRange(inputCol + i).getValue()
  }

  SpreadsheetApp.getUi().alert("Process finished!")


}


function onInstall() {
  onOpen();
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Metabase Integration')
    .addItem('Import query with input parameter', 'importMBData')
    .addToUi();
}
