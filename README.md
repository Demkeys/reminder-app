# Reminder App
This is an app that keeps track of records in a Google Spreadsheet and emails the user when a record's expiry date is approaching. The emails are sent at intervals instead of daily. More details in the Overview section.

In this repo I'm only sharing the script, since that's the important part, but there are a couple of other things that need to be setup in order for this to work. I explain that in the Instructions section.

___Note: I made this app to help someone out, but figured I'd share it on GitHub, for whoever it helps.___

# Overview
How this app works is, user creates a record containing various details of a document. These details include:
- EmployeeName: Containing the Employee's Name
- DocName: Containing the name of the document
- DocInfo: Containing info about the type of document
- DocDoI: Containing the Date of Issue of the document
- DocDoE: Containing the Date of Expiry of the document
- WarningDays: Mentioning how many days before expiry we should start warning the user.
User can create multiple records. User can setup a Trigger (OnOpen, Time-based, etc.) that will execute the CheckRecords function from the script. The CheckRecords function iterates over each record, and if the remaining days till expiry is less than WarningDays, it will notify the user about it via email.

More about notifying the user:
- All the soon expiring records are collected into one message, so the user is only sent one email with all the info, instead of multiple emails.
- Emails are not sent out daily. WarningDays is divided into step intervals, and when the RemainingDays hits one of those intervals (set by 'noOfStep' variable in code), the email is sent out. For example, if WarningDays is 30 and noOfSteps is 4, this will be divided into 5 steps [0, 8, 15, 23, 30]. When RemainingDays is at any of these days, the user will get an email. In total the user should only receive 4-5 emails until expiry.

Below are instructions on how I think you should set up the system but to be honest there is no right way. Set it up according to your needs. The three key components here are the Spreadsheet, Script and Triggers. 

___Important Note: Performing certain actions (for example sending an email) from script requires permissions. If you don't have the required permission to perform the action, trying to do it via Trigger might cause the execution to fail. To get around this, execute the function via the Apps Script Code Editor window. At that point a popup will appear asking you to grant permission to the app. Once you've granted the app permission for those actions, you should be able to execute them via Triggers with no issues.___

# Instructions
As mentioned above, this is how I would suggest setting up the app because this is what worked for my scenario. Do what works for you.
- Create Google Spreadsheet with the following columns:
![](https://github.com/Demkeys/reminder-app/blob/main/sheetformat.png)
  - The names of the columns don't matter, but their index matters, because that is how the script reads them. 
  - RemainingColor and RemainingDays values are calculated and set from script, so you don't need to add any data there.
  - Set DocDoI and DocDoE columns to a custom Date Time format using Format > Number > Custom date and time. The format must be dd-MMM-yyyy (eg. 04-Apr-2023). This ensures that dates are parsed correctly in the script. DocDoI isn't used in the script, so you don't need to worry too much about it.
- Open the code editor window using Extensions > Apps Script. This is script that is contained within the spreadsheet, so it has access to various aspects of the spreadsheet. Copy the code from [code.gs](https://github.com/Demkeys/reminder-app/blob/main/code.gs) file and paste it into your 'code.gs' file.
  - In the code:
     - Set totalRows value to the count of your rows (not including the labels row).
     - Set noOfSteps value to how many intervals you want to divide WarningDays into (read above for explanation on noOfSteps).
- Set up Triggers from the Triggers window. I have set up two Triggers for my purpose, which I will describe here. You can set up your Triggers according to your need.
  - Trigger 1 (This trigger will execute once a day, at the time range you specify)
    - Function to run: CheckRecords
    - Event Source: Time-driven
    - Event Type: Day Timer (set whatever time you want)
  - Trigger 2 (This trigger will execute once when you open the spreadsheet. It will execute CheckRecords and then pop up a message for the user)
    - Function to run: OnOpenTrigg
    - Event Source: From spreadsheet
    - Event Type: On open

