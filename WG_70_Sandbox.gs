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

// **********************************************
// function subTestContactGroup()
//
//
// **********************************************

function subTestContactGroup(){

  var PlyrContactInfo = new Array(4);
  
  PlyrContactInfo[0]= 'Eric';
  PlyrContactInfo[1]= 'Bouchard';
  PlyrContactInfo[2]= 'ericbouchard9@gmail.com';
//  PlyrContactInfo[3]= 'English';
  PlyrContactInfo[3]= 'Français';
  
  var shtConfig = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Config');
  
  //  PlyrContactInfo[0]= First Name
  //  PlyrContactInfo[1]= Last Name
  //  PlyrContactInfo[2]= Email
  //  PlyrContactInfo[3]= Language
  
  var cfgEvntParam = shtConfig.getRange(4,4,32,1).getValues();
  
  // Event Parameters
  var evntLocation = cfgEvntParam[0][0];
  var evntNameEN = cfgEvntParam[7][0];
  var evntNameFR = cfgEvntParam[8][0];
  var evntCntctGrpNameEN = evntLocation + " " + evntNameEN;
  var evntCntctGrpNameFR = evntLocation + " " + evntNameFR;
  var ContactGroupEN;
  var ContactGroupFR;
  
  var Status;
    
  // Delete Contact Groups
  // Get Contact Group
  ContactGroupEN = ContactsApp.getContactGroup(evntCntctGrpNameEN);
  ContactGroupFR = ContactsApp.getContactGroup(evntCntctGrpNameFR);
  // If Contact Group exists, Delete it
  if(ContactGroupEN != null) ContactsApp.deleteContactGroup(ContactGroupEN);
  if(ContactGroupFR != null) ContactsApp.deleteContactGroup(ContactGroupFR);
  
  
}