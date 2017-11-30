function tstCreateMember(){

  var Member = subCreateArray(16,1);
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Full Name
  //  Member[ 2] = Member First Name
  //  Member[ 3] = Member Last Name
  //  Member[ 4] = Member Email
  //  Member[ 5] = Member Language
  //  Member[ 6] = Member Phone Number
  //  Member[ 7] = Member Record File ID
  //  Member[ 8] = Member Record File Link
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  Member[ 0] = ""
  Member[ 1] = "Eric Bouchard" // Member Full Name
  Member[ 2] = "Eric" // Member First Name
  Member[ 3] = "Bouchard" // Member Last Name
  Member[ 4] = "ericbouchard9@gmail.com" // Member Email
  Member[ 5] = "English" // Member Language
  Member[ 6] = "514-318-7571" // Member Phone Number
  Member[ 7] = ""
  Member[ 8] = ""
  
  fcnCreateMember(Member);
  
  

}

function fcnCreateMember(Member, GameType) {
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Full Name
  //  Member[ 2] = Member First Name
  //  Member[ 3] = Member Last Name
  //  Member[ 4] = Member Email
  //  Member[ 5] = Member Language
  //  Member[ 6] = Member Phone Number
  //  Member[ 7] = Member Record File ID
  //  Member[ 8] = Member Record File Link
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  // Create Player Record Spreadsheet
  var ss = SpreadsheetApp.create(Member[1]);
  var id = ss.getId();
  // Get File
  var file = DriveApp.getFileById(id);
  var folders = DriveApp.getFoldersByName('Members');
  while (folders.hasNext()) {
    var folder = folders.next();
  }
  // Move File to Folder
  folder.addFile(file);
  // Remove File from Root Folder
  DriveApp.getRootFolder().removeFile(file);
  
  // Open Member List (File _Member List)
  var shtList = SpreadsheetApp.openById("1wQVwIIFJoRTNfK3oV4ioeFU3hw6JnEA4MXmc-GI8Vgw").getSheetByName("List");
  var listNbMembers = shtList.getRange(2,1).getValue();
  var listMaxCol = shtList.getMaxColumns();
  var listMaxRow = shtList.getMaxRows();
  var listNextMemberRow = listNbMembers + 3; // First Member is at Row 3
  
  var NewMemberID = listNbMembers + 1;
  
  // Open Member Spreadsheet
  var ssMember = SpreadsheetApp.openById(id);
  
  // Copy Templates to Member Spreadsheet
  var ssTemplates = SpreadsheetApp.openById("1HPnrUdIen2X0YeV1R2eNf5CEOZ9Yq_GYJEWTfAuIAMU");
  ssTemplates.getSheetByName("Member Info").copyTo(ssMember);
  ssTemplates.getSheetByName("WG Record Template").copyTo(ssMember);
  ssTemplates.getSheetByName("TCG Record Template").copyTo(ssMember);
  ssTemplates.getSheetByName("BG Record Template").copyTo(ssMember);
  
  // Delete Sheet1
  ssMember.deleteSheet(ssMember.getSheets()[0]);
  
  // Get Number of Sheets
  var memberNbSheets = ssMember.getNumSheets();
  var memberSheets = ssMember.getSheets();
  var SheetName;
  var Sheet;
  
  // Rename all Copied Sheets
  for(var sht= 0; sht < memberNbSheets; sht++){
    Sheet = memberSheets[sht];
    SheetName = Sheet.getName();
    Logger.log("sht: %s - SheetName: %s",sht,SheetName);
    // Rename all tabs
    switch(SheetName){
      case "Copy of Member Info"         : Sheet.setName("Info"); break;
      case "Copy of WG Record Template"  : Sheet.setName("WG Record"); break;
      case "Copy of TCG Record Template" : Sheet.setName("TCG Record"); break;
      case "Copy of BG Record Template"  : Sheet.setName("BG Record"); break;
    }
  }
  
  // Update Member Info with New Data
  Member[0] = NewMemberID;
  Member[7] = id;
  Member[8] = 'https://docs.google.com/spreadsheets/d/' + id;
  
  
  // Write Member Values in _Member List
  var shtInfo = ssMember.getSheetByName("Info");
  var valMemberSheet = shtInfo.getRange(1, 2, 16, 1).getValues();
  var valMemberList =  shtList.getRange(listNextMemberRow,1, 1, listMaxCol).getValues();
  
  // Set Member Info Values for both sheets
  for(var i= 0; i<16; i++){
    valMemberList[0][i+1] = Member[i];
    valMemberSheet[i][0]  = Member[i];
    
  }
  // Add New Member to _Member List
  shtList.insertRowBefore(listMaxRow);
  shtList.getRange(listNextMemberRow+1,1, 1, listMaxCol).setValues(valMemberList);
  shtList.getRange(listNextMemberRow+1, 1).setValue('=if(INDIRECT("R[0]C[1]",false)<>"",1,0)');
  
  // Write Member Info to Member Spreadsheet
  shtInfo.getRange(1, 2, 16, 1).setValues(valMemberSheet);
  
  Logger.log(valMemberList);
  Logger.log(valMemberSheet);
  return Member;
}


function testSearch(){

  var PlayerName = 'Russ';
  var SearchStatus = SearchFiles(PlayerName);
  Logger.log(SearchStatus[0]);
  Logger.log(SearchStatus[1]);
  Logger.log(SearchStatus[2]);
}


function SearchFiles(PlayerName) {
  
  var Status = subCreateArray(3,1);
  
  Status[0] = 'Member Not Found';
  
  var searchFor ='title contains "' + PlayerName + '"';
  var names =[];
  var fileIds=[];
  var files = DriveApp.searchFiles(searchFor);
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();// To get FileId of the file
    fileIds.push(fileId);
    var name = file.getName();
    names.push(name);

  }

  if(names.length == 0){
    Status[0] = 'Member Not Found';
    Status[1] = '';
    Status[2] = '';
  }
  
  if(names.length > 0){
    Status[0] = 'Member Found';
    Status[1] = names[0];
    Status[2] = 'https://docs.google.com/spreadsheets/d/' + fileIds[0];
  }
  
  return Status;

}



// **********************************************
// function fcnCreateGlobalPlayerRecord()
//
// This function creates all Players Records 
// from the Config File
//
// **********************************************

function fcnCreateGlobalPlayerRecord(){

  Logger.log("Routine: fcnCreateGlobalPlayerRecord");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange(4,28,30,1).getValues();
  
  // Player Log Spreadsheet
  var ssPlayerRecord = SpreadsheetApp.openById(shtIDs[13][0]);
  var shtTemplate = ssPlayerRecord.getSheetByName('Template');
  var shtArmyListNum;
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrLang = cfgColShtPlyr[5][0];
  
  // Sheets Values
  var NbSheet = ssPlayerRecord.getNumSheets();
  var ssSheets = ssPlayerRecord.getSheets();
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  // Get Players Data
  // [0]= Player Name, [colShtPlyrLang - colShtPlyrName]= Language Preference 
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, colShtPlyrLang-colShtPlyrName+1).getValues();
  
  // Routine Variables
  var shtPlyr;
  var PlyrName;
  var SheetName;
  var GlobalHdr;
  var HstryHdr;
  var PlayerFound = 0;
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Gets the Player Name 
    PlyrName = PlayerData[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
          
    // If Player is not found, add a tab
    if(PlayerFound == 0){
      
      var ssPlayerFile = fcnCreateMember(PlyrName);
      Logger.log(ssPlayerFile.getName());
      
//      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
//      shtTemplate.copyTo(ssPlayerFile).activate();
//      //ssPlayerFile.insertSheet(PlyrName, NbSheet-1, {template: shtTemplate});
//      shtPlyr = ssPlayerFile.getSheetByName(PlyrName);
//      
//      // Opens the new sheet and modify appropriate data (Player Name, Header)
//      shtPlyr.getRange(2,1).setValue(PlyrName);
//      
//      // Translate Header if Player's Language Preference is French
//      if(PlayerData[plyr][colShtPlyrLang - colShtPlyrName] == 'Français'){
//        // Set Global Header
//        GlobalHdr = shtPlyr.getRange(3,1,1,6).getValues();
//        GlobalHdr[0][0] = 'Joué';       // Played
//        GlobalHdr[0][1] = 'Victoires';  // Win
//        GlobalHdr[0][2] = 'Défaites';   // Loss
//        GlobalHdr[0][3] = 'Nulles';     // Tie
//        GlobalHdr[0][4] = 'Points';     // Points
//        GlobalHdr[0][5] = '% Victoire'; // Win%
//        shtPlyr.getRange(3,1,1,6).setValues(GlobalHdr);
//      
//        // Set Hstry Header
//        HstryHdr = shtPlyr.getRange(6,1,1,7).getValues();
//        HstryHdr[0][0] = 'Événement';   // Event Name
//        HstryHdr[0][1] = ''; // Event Name (merged cell)
//        HstryHdr[0][2] = 'Ronde'; // Round
//        HstryHdr[0][3] = 'Résultat'; // Match Result
//        HstryHdr[0][4] = 'Joué contre'; // Played vs
//        HstryHdr[0][5] = ''; // Played vs (merged cell)
//        HstryHdr[0][6] = 'Points Marqués'; // Scored Points
//        shtPlyr.getRange(6,1,1,7).setValues(HstryHdr);
//      }
    }
  }
//  // English Version
//  ssPlayerRecord.setActiveSheet(ssPlayerRecord.getSheets()[0]);
//  ssPlayerRecord.getSheetByName('Template').hideSheet();

}








function testFunctionCall(){

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Sheet to get options
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  
  fcnUpdateStandings(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData);

  fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
  
}


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
  
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  
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