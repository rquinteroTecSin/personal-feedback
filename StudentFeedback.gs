/*
Sergio Ruiz-Loza, Ph.D.
sergio.ruiz.loza@tec.mx
Computer Department. CCM.
201913
Code.gs
Sends batch emails with feedback to students.

Base code: https://www.benlcollins.com/spreadsheets/google-forms-survey-tool/
*/

// add menu to Sheet
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu("Send Emails")
  .addItem("Send Email Batch","createEmail")
  .addToUi();
}

/**
 * take the range of data in sheet
 * use it to build an HTML email body
 */
function createEmail() {
  
  
  
  var SHEET_NAME  = 'Sheet4';                               // <<<< 1. Set source SHEET here. The script will work using this sheet.
  var STATUS_COL  = "STATUS1";                              // <<<< 2. Set STATUS# date to update here (1-3). This is the day that the email was sent. Updated automatically.
  var MESSAGE_COL = "MESSAGE1";                             // <<<< 3. Set MESSAGE# to send here (1-3). Choose the partial feedback number to send.
  var FORM_URL    = 'https://forms.gle/N2ybiWThmzLTapAU6';  // <<<< 4. URL for the Google Form to collect the student reflection. You should create one for each partial!
  var PROFESSOR   = 'Sergio Ruiz-Loza, Ph.D.';              // <<<< 5. Professor name for the email text.
  
  var thisWorkbook = SpreadsheetApp.getActiveSpreadsheet();
  var thisSheet = thisWorkbook.getSheetByName(SHEET_NAME);

  // get the data range of the sheet
  var allRange = thisSheet.getDataRange();
  
  // get all the data in this range
  var allData = allRange.getValues();
  
  // get the header row
  var headers = allData.shift();
  
  // create header index map
  var headerIndexes = indexifyHeaders(headers);
  
  allData.forEach(function(row,i) {
    /*
    When the STATUS date is already set, don't send a message to that student.
    
    When I downloaded student lists from the system, some 'Retired' students appeared
    at the bottom. The second condition will ignore those.
    */
    if (!row[headerIndexes[STATUS_COL]] && row[headerIndexes["RETIRED"]] == "NO") { 
      var   htmlBody = 
        "Hello " + row[headerIndexes["NAME"]] +",<br><br>" +
         "Your feedback for the partial evaluation is as follows:<br><br>" +
          row[headerIndexes[MESSAGE_COL]] + 
           "<br/><br/>" + 
             "<b>The next step is required and is a part of your final grade:</b><br/>"+
             "Considering this feedback, reflect by entering this " +
             "<a href=\""+FORM_URL+"\" target=_blank>form</a>" +
             " and filling the required spaces. Please do so before next class.<br/>"  +
               "<b>Important</b>: Remember to use your TEC email and password to enter the form.<br/>" +
                 "<br/>Regards,<br/>"+PROFESSOR+"<br/>";
      
      var timestamp = sendEmail(row[headerIndexes["EMAIL"]], htmlBody);
      thisSheet.getRange(i + 2, headerIndexes[STATUS_COL] + 1).setValue(timestamp);
    }
    else {
      Logger.log("No email sent for this row: " + i + 1);
    }
  });
}

/**
 * create index from column headings
 * @param {[object]} headers is an array of column headings
 * @return {{object}} object of column headings as key value pairs with index number
 */
function indexifyHeaders(headers) {
  
  var index = 0;
  return headers.reduce (
    // callback function
    function(p,c) {
      //skip cols with blank headers
      if (c) {
        // can't have duplicate column names
        if (p.hasOwnProperty(c)) {
          throw new Error('duplicate column name: ' + c);
        }
        p[c] = index;
      }
      index++;
      return p;
    },
    {} // initial value for reduce function to use as first argument
  );
}

/**
 * send email from GmailApp service
 * @param {string} recipient is the email address to send email to
 * @param {string} body is the html body of the email
 * @return {object} new date object to write into spreadsheet to confirm email sent
 */
function sendEmail(recipient,body) {
  
  GmailApp.sendEmail(
    recipient,
    "Computational Thinking for Engineering. Partial exam feedback", 
    "",
    {
      htmlBody: body
    }
  );
  
  return new Date();
}
