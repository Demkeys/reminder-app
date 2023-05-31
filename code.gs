function myFunction() {
  CheckRecords();  
  
}

// This method is called through a Trigger, either OnOpen of the 
// spreadsheet, or some Time-based Trigger.
function onOpenTrigg() {

  // Check all the records and do 
  CheckRecords();

  // Finally display an alert message so user knows that this
  // script has been executed and they can now close the document.
  SpreadsheetApp.getUi().alert('Reminder App',
`Welcome to Reminder App. This app keeps track of legal documents and reminds the user if a document is close to expiring. 
You can use the WarningDays column to decide when the app should start reminding.
RemainingColor column tells you whether a document is close to expiring.
RemainingDays columns will tells you the number of days till it expires.
If a document is close to expiring, you will recieve an email.
If you are seeing this message, all records have been checked and you can now close this spreadsheet.`,
    SpreadsheetApp.getUi().ButtonSet.OK
    );
}

function CheckRecords() {
  // Converts Milliseconds to days. Returns number.
  function MillisecondsToDays(val) {
    return (((val*0.001)/60)/60)/24;
  }

  // If this flag is true, atleast one doc is close to expiring, so
  // read docExpiryMsgArr to form the email message.
  var someDocsExpiring = false;
  
  // String array containing messages of each expiring doc.
  // This array will later be concatanated into one long message
  // and sent as one email.
  var docExpiryMsgArr = []; 

  // Column index starts at 1
  const colIndex = {
    RemainingColor: 1,
    RemainingDays: 2,
    EmployeeName: 3,
    DocName: 4,
    DocInfo: 5,
    DocDoI: 6,
    DocDoE: 7,
    WarningDays: 8
  };
  
  // Total rows containing records. This does not include Row 1
  // which contains labels.
  const totalRows = 6;
  
  // Row 1 is labels, so offset and then start count.
  const totalRowsWithOffset = 1+totalRows;

  // Get current spreadsheet and sheet.
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheets()[0];

  // Blank constructor gives today's date.
  var dateToday = new Date();

  // Set dateToday's time to 0 because retrieved dates will
  // also have 0 (Midnight) time.
  dateToday.setHours(0, 0, 0, 0);

  // Pull data from sheet.
  for(var i = 2; i <= totalRowsWithOffset; i++)
  {
    // Create rowData object with data from 'i' row.
    const rowData = {
      EmployeeName: sheet.getRange(i, colIndex.EmployeeName).getValue(),
      DocName: sheet.getRange(i, colIndex.DocName).getValue(),
      DocInfo: sheet.getRange(i, colIndex.DocInfo).getValue(),
      DocDoI: sheet.getRange(i, colIndex.DocDoI).getValue(),
      DocDoE: sheet.getRange(i, colIndex.DocDoE).getValue(),
      WarningDays: sheet.getRange(i, colIndex.WarningDays).getValue()
    };
    
    // Retrieved date could be string or object, so check type,
    // check parse accordingly and return Date object.
    // UPDATE: This method may not be necessary anymore since we
    // have changed the date format in the spreadsheet to dd-MMM-yyyy.
    // This means that the data will pretty much always be parsed 
    // as a Date object, so the string condition isn't necessary. But
    // for lazy reasons, leave this methods as is. It works.
    function ParseDate(dt) {
      // If dt is an object, pass it back out unchanged.
      if(typeof(dt) == "object") {
        return dt;
      }
      // If dt is a string, split it and construct Date object out 
      // of it. Months start at 0, so subtract by 1.
      else if(typeof(dt) == 'string') {
        var [d, m, y] = dt.split('/');
        return new Date(parseInt(y), parseInt(m)-1, parseInt(d));
      }
    }
    
    var docDoI = ParseDate(rowData.DocDoI); // Unused.
    var docDoE = ParseDate(rowData.DocDoE);

    // Calculate and set RemainingDays and RemainingColor cells.
    var remDays = MillisecondsToDays(docDoE.getTime()-dateToday.getTime());
    sheet.getRange(i, colIndex.RemainingDays).setValue(remDays);
    
    // If remDays is greater than WarningDays, set cell color to green.
    // Nothing more needs to be done.
    if(remDays > parseInt(rowData.WarningDays)){
      sheet.getRange(i, colIndex.RemainingColor)
        .setBackgroundRGB(0,255,0);
    }
    // If remDays is less than or equal to WarningDays, we are 
    // approaching expiry date, so user must be warned. Set cell color 
    // to red and then calculate step days and see if remDays falls on
    // one of the step days. If it does, create a message chunk.
    else if(remDays <= parseInt(rowData.WarningDays)){
      sheet.getRange(i, colIndex.RemainingColor)
        .setBackgroundRGB(255,0,0);

      const noOfSteps = 4; // Total number of steps.
      const stepVal = rowData.WarningDays/4; // Value of one stepped day in WarningDays.

      /* In each iteration of noOfSteps, calculate step days of 
      WarningDays using noOfSteps. Then use Math.ceil to get the upper 
      integer value. Then compare that value to remDays. If true, remDays
      is at one of the step days, so create a message chunk for this
      record. For example:
      WarningDays = 30.
      Steps = 4.
      Step Day Values = [0, 7.5, 15, 22.5, 30]
      Ceil Step Day Values = [0, 8, 15, 23, 30]
      Steps is 4, but we run the loop to 4+1, so that 30 gets included
      in the calculation, so the user can start getting reminders at 30
      instead of starting at 23.
      */
      for(var j = 0; j <= noOfSteps; j++)
      {
        var ceilStepVal = Math.ceil(stepVal*j);
        console.log(ceilStepVal);
        if(remDays == ceilStepVal)
        {
          // Since a doc is close to expiring, form a message chunk and
          // push it to docExpiryMsgArr.
          docExpiryMsgArr.push(
`Employee: ${rowData.EmployeeName}
Document: ${rowData.DocName}
Expires In: ${remDays} Days`);
          // Then set someDocsExpiring flag to true to after this loop ends
          // a full message can be formed and emailed.
          someDocsExpiring = true;
        }
      }
      
    }
    
    // console.log(docDoE);
    // console.log(typeof(docDoI));
    // console.log(typeof(docDoE));
    // console.log(remDays);
    
  }

  // If any docs are expiring, form message and send email.
  // console.log(docExpiryMsgArr);
  // return;
  if(someDocsExpiring)
  {
    var msgBody = 
`Some documents are expiring soon. Renew them, and then update their Date of Issue and Date of Expiry in the Reminder App spreadsheet.
-----------------------------------
`;
    for(var i = 0; i < docExpiryMsgArr.length; i++)
    {
      msgBody +=
`${docExpiryMsgArr[i]}
***
`;
    }
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      'Reminder App: Some documents expiring soon.',
      msgBody
    );
  }
  // console.log(docExpiryMsgArr);
}

function myFunction02() {
  var dateToday = new Date();
  
  var dt04 = new Date(2023,9,20,0,0,0,0);
  var dt05 = new Date(2023,5,29,0,0,0,0);
  // dt04.setHours(30);
  console.log(
    ((((dt04.getTime()-dt05.getTime())/1000)/60)/60)/24
    );

  // Checks if compVal is close to compTarget. The distance between
  // compVal and compTarget should be less than closeThreshold for
  // compVal to be considered close.
  function IsValueCloseToTarget(compVal, compTarget, closeThreshold) {
    return (Math.abs(compTarget-compVal)<closeThreshold);
  }
  console.log(Math.ceil(7.5)==8);

}
