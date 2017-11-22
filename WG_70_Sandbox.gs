function fcnTestEmail(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Config Sheet to get options
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Get Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
  var Service = shtConfig.getRange(20,1).getValue();
  
  // Set Email Subject
  var EmailSubjectEN = "Test Email"
  
  // Start of Email Message
  var EmailMessageEN = '<html><body>';
  
  // Add Signature
  EmailMessageEN += "Test<br><br>This is a email to test the daily quota"+
    "<br><br>Thank you for using TCG Booster League Manager from Turn 1 Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessageEN += '</body></html>';
  
  var Recipients = "ericbouchard9@gmail.com"
//  var Recipients = "ericbouchard9@gmail.com, turn1glt@gmail.com"
    
  if(Service == "Mail"){
    MailApp.sendEmail(Recipients, EmailSubjectEN, "",{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessageEN});
    // Post Quota to Log Sheet
    var MailQuota = "Remaining Emails to send: " + MailApp.getRemainingDailyQuota();
  }
  
  if(Service == "Gmail"){
    GmailApp.sendEmail(Recipients, EmailSubjectEN, "",{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessageEN});
    // Post Quota to Log Sheet
    var MailQuota = "Remaining Emails to send: " + MailApp.getRemainingDailyQuota();
  }
      
  subPostLog(shtLog, MailQuota);

}