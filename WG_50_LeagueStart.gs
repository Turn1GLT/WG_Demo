// **********************************************
// function fcnUpdateLinksIDs()
//
// This function updates all sheets Links and IDs  
// in the Config File
//
// **********************************************

function fcnUpdateLinksIDs(){
  
  Logger.log("Routine: fcnUpdateLinksIDs");
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Copy Log Spreadsheet
  var shtCopyLogID = shtConfig.getRange(9,15).getValue();
  var LinksStatus = shtConfig.getRange(9,16).getValue();
  
  if (shtCopyLogID != '' && LinksStatus =='') {
    var shtCopyLog = SpreadsheetApp.openById(shtCopyLogID).getSheets()[0];
  
    var CopyLogNbFiles = shtCopyLog.getRange(2, 6).getValue();
    var rowStartCopyLog = 5;
    var rowStartConfig = 4;
    var colShtId = 7;
    var colShtUrl = 11;
    
    var CopyLogVal = shtCopyLog.getRange(rowStartCopyLog, 2, CopyLogNbFiles, 3).getValues();
    // [0]= Sheet Name, [1]= Sheet URL, [2]= Sheet ID
    
    var FileName;
    var Link;
    var Formula;
    var rowCfg = 'Not Found';
    
    // Clear Sheet IDs
    shtConfig.getRange(rowStartConfig, colShtId,20,1).clearContent();
    // Clear Sheet URLs
    shtConfig.getRange(rowStartConfig,colShtUrl,20,1).clearContent();
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master WG Event' :
          rowCfg = rowStartConfig + 0; break;
        case 'Master WG Log' :
          rowCfg = rowStartConfig + 1; break;
        case 'Master WG Army DB' :
          rowCfg = rowStartConfig + 2; break;
        case 'Master WG Army Lists EN' :
          rowCfg = rowStartConfig + 3; break;
        case 'Master WG Army Lists FR' :
          rowCfg = rowStartConfig + 4; break;
        case 'Master WG Standings EN' :
          rowCfg = rowStartConfig + 5; break;
        case 'Master WG Standings FR' :
          rowCfg = rowStartConfig + 6; break;	
        case 'Master WG Player Records' :
          rowCfg = rowStartConfig + 13; break;
        case 'Master WG Player List & Round Bonus' :
          rowCfg = rowStartConfig + 14; break;
        case 'Master WG Starting Pool' :
          rowCfg = rowStartConfig + 15; break;        
        default : 
          rowStartConfig = 'Not Found'; break;
      }
      
      // Set the Appropriate Sheet ID Value and URL in the Config File
      if (rowCfg != 'Not Found') {
        shtConfig.getRange(rowCfg, colShtId).setValue(CopyLogVal[row][2]);
        // Opens Spreadsheet by ID to get URL
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();        
        shtConfig.getRange(rowCfg, colShtUrl).setValue(Link);
      }
    }
  }
}

// **********************************************
// function fcnInitializeEvent()
//
// This function clears all data from sheets  
// to start a new Event (League / Tournament)
//
// **********************************************

function fcnInitializeEvent(){
  
  Logger.log("Routine: fcnInitializeEvent");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var cfgEventType = shtConfig.getRange(7,4).getValue();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "Clear "+ cfgEventType +" Data Confirmation";
  var msg = "Click OK to clear all "+ cfgEventType +" Data to start a new " + cfgEventType;
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
//  if(cfgEventType == "League" || cfgEventType == "Tournament"){
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
    var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
    // Registration Form Construction 
    // Column 1 = Category Name
    // Column 2 = Category Order in Form
    // Column 3 = Column Value in Player/Team Sheet
    var cfgRegFormCnstrVal = shtConfig.getRange(4,26,20,3).getValues();
    
    // Event Parameters
    var evntLocation = cfgEvntParam[0][0];
    var evntNameEN = cfgEvntParam[7][0];
    var evntNameFR = cfgEvntParam[8][0];
    var evntCntctGrpNameEN = evntLocation + " " + evntNameEN;
    var evntCntctGrpNameFR = evntLocation + " " + evntNameFR;
    var ContactGroupEN;
    var ContactGroupFR;
    
    // Columns from Config File
    var colRspMatchID        = cfgColRspSht[1][0];
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    
    // Column Round Sheets
    var colRndMP             = cfgColRndSht[2][0];
    
    var colPlyrName   = cfgRegFormCnstrVal[ 2][2];
    var colPlyrStatus = cfgRegFormCnstrVal[16][2];
    
    // Sheets
    var shtStandings =   ss.getSheetByName('Standings');
    var shtRound       = ss.getSheetByName('Round1');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtResponses   = ss.getSheetByName('Responses');
    var shtMatchRespEN = ss.getSheetByName('MatchResp EN');
    var shtMatchRespFR = ss.getSheetByName('MatchResp FR');
    var shtPlayers =     ss.getSheetByName('Players');
    var ssExtPlayers = SpreadsheetApp.openById(shtIDs[14][0]); // External Player List Spreadsheet
    var shtExtPlayers = ssExtPlayers.getSheetByName('Players');// External Player List Sheet
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtMatchRespEN.getMaxRows();
    var MaxColRspnEN = shtMatchRespEN.getMaxColumns();
    var MaxRowRspnFR = shtMatchRespFR.getMaxRows();
    var MaxColRspnFR = shtMatchRespFR.getMaxColumns();
    var MaxRowRndSht = shtRound.getMaxRows();
    var MaxColRndSht = shtRound.getMaxColumns();
    var MaxRowPlayers = shtPlayers.getMaxRows();
    var MaxColPlayers = shtPlayers.getMaxColumns();
        
    // Clear Data
    // Standings
    shtStandings.getRange(6,2,MaxRowStdg-5,MaxColStdg-1).clearContent();
    // Match Results (does not clear the last column)
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-1).clearContent();
    // Responses
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtMatchRespEN.getRange(2,1,MaxRowRspnEN-1,MaxColRspnEN).clearContent();
    shtMatchRespFR.getRange(2,1,MaxRowRspnFR-1,MaxColRspnFR).clearContent()
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndMP,MaxRowRndSht-4,MaxColRndSht-colRndMP+1).clearContent();
    }
    Logger.log('Event Data Cleared');
    
    // Clear Player List
    // From Player Name to Status
    shtPlayers.getRange(3, 2, MaxRowPlayers-2, colPlyrStatus-colPlyrName).clearContent();
    // From Status to rest of File
    shtPlayers.getRange(3, colPlyrStatus+1, MaxRowPlayers-2, MaxColPlayers-colPlyrStatus).clearContent();
    shtExtPlayers.getRange(3, 2, MaxRowPlayers-2, MaxColPlayers-1).clearContent();
    Logger.log('Player List Cleared');
    
    // Delete Contact Groups
    // Get Contact Group
    ContactGroupEN = ContactsApp.getContactGroup(evntCntctGrpNameEN);
    ContactGroupFR = ContactsApp.getContactGroup(evntCntctGrpNameFR);
    // If Contact Group exists, Delete it
    if(ContactGroupEN != null) ContactsApp.deleteContactGroup(ContactGroupEN);
    if(ContactGroupFR != null) ContactsApp.deleteContactGroup(ContactGroupFR);
    Logger.log('Contact Groups Deleted');
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
    Logger.log('Standings Updated');
    
    // Clear Players DB and Card Pools
    fcnDelPlayerArmyDB();
    fcnDelPlayerArmyList();
    fcnDelEventRecord();
    Logger.log('Army DB and Army Lists Cleared');
        
    title = cfgEventType +" Data Cleared";
    msg = "All " + cfgEventType +" Data has been cleared. You are now ready to start a new " + cfgEventType;
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);    
  }
}

// **********************************************
// function fcnClearMatchResults()
//
// This function clears all Results data but
// does not clear Responses
//
// **********************************************

function fcnClearMatchResults(){
  
  Logger.log("Routine: fcnClearMatchResults");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var cfgEventType = shtConfig.getRange(7,4).getValue();
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title = "Reset " + cfgEventType + " Match Results";
  var msg = "Click OK to clear all "+ cfgEventType +" match results";
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
    
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
    var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
    
    // Columns from Config File
    var colRspMatchID        = cfgColRspSht[1][0];
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    
    // Column Round Sheets
    var colRndMP             = cfgColRndSht[2][0];
    
    // Sheets
    var shtStandings   = ss.getSheetByName('Standings');
    var shtRound       = ss.getSheetByName('Round1');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtResponses   = ss.getSheetByName('Responses');
    var shtMatchRespEN = ss.getSheetByName('MatchResp EN');
    var shtMatchRespFR = ss.getSheetByName('MatchResp FR');
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtMatchRespEN.getMaxRows();
    var MaxColRspnEN = shtMatchRespEN.getMaxColumns();
    var MaxRowRspnFR = shtMatchRespFR.getMaxRows();
    var MaxColRspnFR = shtMatchRespFR.getMaxColumns();
    var MaxRowRndSht = shtRound.getMaxRows();
    var MaxColRndSht = shtRound.getMaxColumns();
    
    // Clear Data
    shtStandings.getRange(6,2,MaxRowStdg-5,MaxColStdg-1).clearContent();
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-1).clearContent();
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtMatchRespEN.getRange(2,colRspMatchID,MaxRowRspnEN-1,MaxColRspnEN-colRspMatchID+1).clearContent();
    shtMatchRespFR.getRange(2,colRspMatchID,MaxRowRspnFR-1,MaxColRspnFR-colRspMatchID+1).clearContent();
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndMP,MaxRowRndSht-4,MaxColRndSht-colRndMP+1).clearContent();
    }
    
    // Clear Event Records
    fcnClrEvntRecord();
    
    Logger.log('Match Data Cleared');
    
    // Update Standings Copies
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, 0, 1);
    Logger.log('Standings Updated');
    
    title = "Match Results Cleared";
    msg = "All Match Results have been cleared. You are now ready to submit Match Reports";
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
}

// **********************************************
// function fcnCrtEvntRecord()
//
// This function generates all Players Records 
// from the Config File
//
// **********************************************

function fcnCrtEvntRecord(){
  
  Logger.log("Routine: fcnCrtEvntRecord");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam =  shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange( 4,28,30,1).getValues();
  var cfgColShtTeam = shtConfig.getRange(24,28,30,1).getValues();
  
  // Event Log Spreadsheet
  var ssEventRecord = SpreadsheetApp.openById(shtIDs[13][0]);
  var shtTemplate = ssEventRecord.getSheetByName('Template');
  var shtArmyListNum;
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrLang = cfgColShtPlyr[5][0];
  
  var colShtTeamName = cfgColShtTeam[7][0];
  var colShtTeamLang = cfgColShtTeam[5][0];
  
  // Event Parameters
  var evntFormat = cfgEvntParam[ 9][0];
  
  // Sheets Values
  var NbSheet = ssEventRecord.getNumSheets();
  var ssSheets = ssEventRecord.getSheets();
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  // Get Players Names and Languages 
  var PlyrNames = shtPlayers.getRange(2,colShtTeamName, NbPlayers+1, 1).getValues();
  var PlyrLang =  shtPlayers.getRange(2,colShtTeamLang, NbPlayers+1, 1).getValues();
  
  // Teams 
  var shtTeams = ss.getSheetByName('Teams'); 
  var NbTeams = shtTeams.getRange(2,1).getValue();
  // Get Teams Names and Languages 
  var TeamNames = shtTeams.getRange(2,colShtTeamName, NbTeams+1, 1).getValues();
  var TeamLang =  shtTeams.getRange(2,colShtTeamLang, NbTeams+1, 1).getValues();
  
  // Routine Variables
  var shtPT;
  var namePT;
  var langPT;
  var nameSheet;
  var GlobalHdr;
  var HstryHdr;
  var LoopMax;
  var PTFound = 0;
  
  // Defines Loop Parameters
  if(evntFormat == "Single") LoopMax = NbPlayers;
  
  if(evntFormat == "Team") LoopMax = NbTeams;
  
  // Loops through each player starting from the Last
  for (var PT = LoopMax; PT > 0; PT--){
    
    // Gets the Player/Team Name and Language
    if(evntFormat == "Single"){
      namePT = PlyrNames[PT][0];
      langPT = PlyrLang[PT][0];
    }
    if(evntFormat == "Team"){
      namePT = TeamNames[PT][0];
      langPT = TeamLang[PT][0];
    }
    
    // Resets the Player/Team Found flag before searching
    PTFound = 0;
    // Look if Player/Team exists, if yes, skip, if not, create Player/Team
    for(var sheet = NbSheet; sheet > 0; sheet --){
      nameSheet = ssSheets[sheet-1].getSheetName();
      if (nameSheet == namePT) PTFound = 1;
    }
          
    // If Player/Team is not found, add a tab
    if(PTFound == 0){
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssEventRecord.insertSheet(namePT, NbSheet-1, {template: shtTemplate});
      shtPT = ssEventRecord.getSheetByName(namePT);
      shtPT.showSheet();
      
      // Updates the number of sheets
      NbSheet = ssEventRecord.getNumSheets();
      ssSheets = ssEventRecord.getSheets();
      
      // Opens the new sheet and modify appropriate data (Player/Team Name, Header)
      shtPT.getRange(2,1).setValue(namePT);
      
      // Translate Header if Player/Team Language Preference is French
      if(langPT == 'Français'){
        // Set Global Header
        GlobalHdr = shtPT.getRange(3,1,1,9).getValues();
        GlobalHdr[0][0] = 'Joué';           // Played
        GlobalHdr[0][1] = 'Victoires';      // Win
        GlobalHdr[0][2] = 'Défaites';       // Loss
        GlobalHdr[0][3] = 'Nulles';         // Tie
        GlobalHdr[0][4] = 'Pts Marqués';    // Pts Scored
        GlobalHdr[0][5] = 'Pts Alloués';    // Pts Allowed
        GlobalHdr[0][6] = '% Victoire';     // Win%
        GlobalHdr[0][7] = 'Pts Mrq/Match'; // Pts Scored / Match
        GlobalHdr[0][8] = 'Pts All/Match'; // Pts Allowed / Match
        shtPT.getRange(3,1,1,9).setValues(GlobalHdr);
      
        // Set History Header
        HstryHdr = shtPT.getRange(6,1,1,9).getValues();
        HstryHdr[0][0] = 'Événement';      // Event Name
        HstryHdr[0][1] = '';               // Event Name (merged cell)
        HstryHdr[0][2] = 'Jeu';            // Game
        HstryHdr[0][3] = 'Ronde';          // Round
        HstryHdr[0][4] = 'Résultat';       // Match Result
        HstryHdr[0][5] = 'Joué contre';    // Played vs
        HstryHdr[0][6] = '';               // Played vs (merged cell)
        HstryHdr[0][7] = 'Points Marqués'; // Points Scored
        HstryHdr[0][8] = 'Points Alloués'; // Points Allowed
        shtPT.getRange(6,1,1,9).setValues(HstryHdr);
      }
    }
  }
  // English Version
  ssEventRecord.setActiveSheet(ssEventRecord.getSheets()[0]);
  ssEventRecord.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnDelEventRecord()
//
// This function deletes all Players/Teams Record Sheets
// from the Config File
//
// **********************************************

function fcnDelEventRecord(){
  
  Logger.log("Routine: fcnDelEventRecord");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[13][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Event Records
  subDelPlayerSheets(shtIDs[13][0]);

}

// **********************************************
// function fcnClrEvntRecord()
//
// This function clears all data in Player Record Sheets
//
// **********************************************

function fcnClrEvntRecord(){

  Logger.log("Routine: fcnClrEvntRecord");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Get Player Log Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var ssEvntPlyrRec = SpreadsheetApp.openById(shtIDs[13][0]);
  var rngRecord = "A4:I4";
  var evntPlyrRecNbSheets = ssEvntPlyrRec.getNumSheets();
  var evntPlyrSheets = ssEvntPlyrRec.getSheets();
  var evntPlyrRowStart = 7;
  
  // Routine Variables
  var sheet;
  var shtMaxCol;
  var shtMaxRow;
  
  // Loop through all Players Sheets
  for(var sht = 0; sht < evntPlyrRecNbSheets; sht++){
    // Get Sheet
    sheet = evntPlyrSheets[sht];
    shtMaxCol = sheet.getMaxColumns();
    shtMaxRow = sheet.getMaxRows();
    
    // Clear Player Record
    sheet.getRange(rngRecord).clearContent();
    
    // Delete all History Rows from Row 8 to Max Row
    if(shtMaxRow > evntPlyrRowStart) sheet.deleteRows(evntPlyrRowStart+1, shtMaxRow-evntPlyrRowStart);
    
    // Clear Player History
    sheet.getRange(evntPlyrRowStart, 1, 1, shtMaxCol).clearContent();
  }
}



// **********************************************
// function fcnCrtPlayerArmyDB()
//
// This function generates all Army DB for all 
// players from the Config File
//
// **********************************************

function fcnCrtPlayerArmyDB(){
  
  Logger.log("Routine: fcnCrtPlayerArmyDB");
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Configuration Data
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange(4,28,20,1).getValues();
  
  // Column Values
  var colShtPlyrName =     cfgColShtPlyr[ 2][0];
  var colShtPlyrTeam =     cfgColShtPlyr[ 5][0];
  var colShtPlyrArmy =     cfgColShtPlyr[10][0];
  var colShtPlyrFaction1 = cfgColShtPlyr[11][0];
  var colShtPlyrFaction2 = cfgColShtPlyr[12][0];
  var colShtPlyrWarlord =  cfgColShtPlyr[13][0];
  
  // Current Army Values (Power Level or Points
  var armyBldRatingMode = cfgArmyBuild[0][0];
  var armyBldArmyValue = cfgArmyBuild[1][0];
  
  // Army DB Spreadsheet
  var ssArmyDB = SpreadsheetApp.openById(shtIDs[2][0]);
  var shtTemplate = ssArmyDB.getSheetByName('Template');
  var NbSheet = ssArmyDB.getNumSheets();
  var ssSheets = ssArmyDB.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var MaxColPlayers = shtPlayers.getMaxColumns();
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  // Get Players Data
  var PlayerData = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, MaxColPlayers-1).getValues();
     
  var shtPlyr;
  var PlyrName;
  var PlyrArmy;
  var PlyrFaction1;
  var PlyrFaction2;
  var PlyrWarlord;
  
  var ArmyDefOffset = 2;
  
  // Loops through each player starting from the last
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
    
    // If Player is not found, create a tab with the player's name
    if(PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssArmyDB.getNumSheets();
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssArmyDB.insertSheet(PlyrName, NbSheet-1, {template: shtTemplate});
      shtPlyr = ssArmyDB.getSheetByName(PlyrName);
      
      // Updates the number of sheets
      NbSheet = ssArmyDB.getNumSheets();
      ssSheets = ssArmyDB.getSheets();
      
      // Get the Player Data
      PlyrArmy =     PlayerData[plyr][colShtPlyrArmy-ArmyDefOffset];
      PlyrFaction1 = PlayerData[plyr][colShtPlyrFaction1-ArmyDefOffset];
      PlyrFaction2 = PlayerData[plyr][colShtPlyrFaction2-ArmyDefOffset];
      PlyrWarlord =  PlayerData[plyr][colShtPlyrWarlord-ArmyDefOffset];
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyr.getRange(3,3).setValue(PlyrName);
      shtPlyr.getRange(4,3).setValue(PlyrArmy);
      if(PlyrFaction2 == '') shtPlyr.getRange(5,3).setValue(PlyrFaction1);
      if(PlyrFaction2 != '') shtPlyr.getRange(5,3).setValue(PlyrFaction1 + ', ' + PlyrFaction2);
      shtPlyr.getRange(6,3).setValue(PlyrWarlord);
      //
      if (armyBldRatingMode == 'Power Level') shtPlyr.getRange(5,9).setValue(armyBldArmyValue);
      if (armyBldRatingMode == 'Points')      shtPlyr.getRange(5,11).setValue(armyBldArmyValue);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyr.hideColumns( 6, 3);
        shtPlyr.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
        shtPlyr.hideColumns( 5, 1);
        shtPlyr.hideColumns( 9, 2);
      }
    }
  }
  shtPlyr = ssArmyDB.getSheets()[0];
  ssArmyDB.setActiveSheet(shtPlyr);
}


// **********************************************
// function fcnCrtPlayerArmyList()
//
// This function generates all accessible Army Lists
// for all players from the Config File
//
// **********************************************

function fcnCrtPlayerArmyList(){
  
  Logger.log("Routine: fcnCrtPlayerArmyDB");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Configuration Data
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange(4,28,30,1).getValues();
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  
  // Current Army Values (Power Level or Points
  var armyBldRatingMode = cfgArmyBuild[0][0];
  var armyBldArmyValue = cfgArmyBuild[1][0];
  
  // Army DB Spreadsheet
  var ssArmyDB = SpreadsheetApp.openById(shtIDs[2][0]); 
  
  // Army Lists Spreadsheet
  var ssArmyListEN = SpreadsheetApp.openById(shtIDs[3][0]);
  var ssArmyListFR = SpreadsheetApp.openById(shtIDs[4][0]);
  var shtTemplateEN = ssArmyListEN.getSheetByName('Template');
  var shtTemplateFR = ssArmyListFR.getSheetByName('Template');
  var shtArmyListNum;
  
  var NbSheet = ssArmyListEN.getNumSheets();
  var ssSheets = ssArmyListEN.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayerNames = shtPlayers.getRange(2,colShtPlyrName, NbPlayers+1, 1).getValues();
    
  var shtPlyrEN;
  var shtPlyrFR;
  var PlyrName;
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerNames[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet; sheet > 0; sheet --){
      SheetName = ssSheets[sheet-1].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
          
    // If Player is not found, add a tab
    if(PlayerFound == 0){
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      // English Version
      ssArmyListEN.insertSheet(PlyrName, NbSheet-1, {template: shtTemplateEN});
      shtPlyrEN = ssArmyListEN.getSheetByName(PlyrName);
      shtPlyrEN.showSheet();
      
      // Updates the number of sheets
      NbSheet = ssArmyListEN.getNumSheets();
      ssSheets = ssArmyListEN.getSheets();
      
      // French Version
      ssArmyListFR.insertSheet(PlyrName, NbSheet-1, {template: shtTemplateFR});
      shtPlyrFR = ssArmyListFR.getSheetByName(PlyrName);
      shtPlyrFR.showSheet();
      
      // Get Player Army DB Values
      fcnCopyArmyDBtoArmyList(shtConfig,PlyrName);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyrEN.hideColumns( 6, 3);
        shtPlyrEN.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
        shtPlyrEN.hideColumns( 5, 1);
        shtPlyrEN.hideColumns( 9, 2);
      }
       
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyrFR.hideColumns( 6, 3);
        shtPlyrFR.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
        shtPlyrFR.hideColumns( 5, 1);
        shtPlyrFR.hideColumns( 9, 2);
      }
    }
  }
  // English Version
  ssArmyListEN.setActiveSheet(ssArmyListEN.getSheets()[0]);
  ssArmyListEN.getSheetByName('Template').hideSheet();
  
  // French Version
  ssArmyListFR.setActiveSheet(ssArmyListFR.getSheets()[0]);
  ssArmyListFR.getSheetByName('Template').hideSheet();
}


// **********************************************
// function fcnCopyArmyDBtoArmyList()
//
// This function copies all data from the Army DB 
// to the Army List for selected Player
//
// **********************************************

function fcnCopyArmyDBtoArmyList(shtConfig,PlyrName){
  
  Logger.log("Routine: fcnCopyArmyDBtoArmyList");

  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Army DB Spreadsheet
  var ssArmyDB = SpreadsheetApp.openById(shtIDs[2][0]); 
  // Army Lists Spreadsheet
  var ssArmyListEN = SpreadsheetApp.openById(shtIDs[3][0]);
  var ssArmyListFR = SpreadsheetApp.openById(shtIDs[4][0]);
  
  // Player Sheet 
  var shtPlyrDB = ssArmyDB.getSheetByName(PlyrName);
  var shtPlyrListEN = ssArmyListEN.getSheetByName(PlyrName);
  var shtPlyrListFR = ssArmyListFR.getSheetByName(PlyrName);
  
  // Army Sheet Range Values
  var rngPlyrHdr   = 'C3:C6';
  var rngRating    = 'I5:L5';
  var rngDtch1Hdr  = 'B10:D10';
  var rngDtch1Data = 'B13:H25';
  var rngDtch2Hdr  = 'B30:D30';
  var rngDtch2Data = 'B33:H45';
  var rngDtch3Hdr  = 'B50:D50';
  var rngDtch3Data = 'B53:H65';

  // Player Header
  var PlyrHdr = shtPlyrDB.getRange(rngPlyrHdr).getValues();
  shtPlyrListEN.getRange(rngPlyrHdr).setValues(PlyrHdr);
  shtPlyrListFR.getRange(rngPlyrHdr).setValues(PlyrHdr);
  
  // Army Rating Level
  var Rating = shtPlyrDB.getRange(rngRating).getValues();
  shtPlyrListEN.getRange(rngRating).setValues(Rating);
  shtPlyrListFR.getRange(rngRating).setValues(Rating);
  
  // Detachment 1
  var Dtch1Hdr = shtPlyrDB.getRange(rngDtch1Hdr).getValues();
  shtPlyrListEN.getRange(rngDtch1Hdr).setValues(Dtch1Hdr);
  shtPlyrListFR.getRange(rngDtch1Hdr).setValues(Dtch1Hdr);
  
  var Dtch1Data = shtPlyrDB.getRange(rngDtch1Data).getValues();
  shtPlyrListEN.getRange(rngDtch1Data).setValues(Dtch1Data);
  shtPlyrListFR.getRange(rngDtch1Data).setValues(Dtch1Data);
    
  // Detachment 2
  var Dtch2Hdr = shtPlyrDB.getRange(rngDtch2Hdr).getValues();
  shtPlyrListEN.getRange(rngDtch2Hdr).setValues(Dtch2Hdr);
  shtPlyrListFR.getRange(rngDtch2Hdr).setValues(Dtch2Hdr);
  
  var Dtch2Data = shtPlyrDB.getRange(rngDtch2Data).getValues();
  shtPlyrListEN.getRange(rngDtch2Data).setValues(Dtch2Data);
  shtPlyrListFR.getRange(rngDtch2Data).setValues(Dtch2Data);
  
  // Detachment 3
  var Dtch3Hdr = shtPlyrDB.getRange(rngDtch3Hdr).getValues();
  shtPlyrListEN.getRange(rngDtch3Hdr).setValues(Dtch3Hdr);
  shtPlyrListFR.getRange(rngDtch3Hdr).setValues(Dtch3Hdr);
  
  var Dtch3Data = shtPlyrDB.getRange(rngDtch3Data).getValues();
  shtPlyrListEN.getRange(rngDtch3Data).setValues(Dtch3Data);
  shtPlyrListFR.getRange(rngDtch3Data).setValues(Dtch3Data);
}




// **********************************************
// function fcnDelPlayerArmyDB()
//
// This function deletes all Players Army DB Sheets 
// from the Config File
//
// **********************************************

function fcnDelPlayerArmyDB(){
  
  Logger.log("Routine: fcnDelPlayerArmyDB");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
   
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[2][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
    
  // Delete Player Army DB
  subDelPlayerSheets(shtIDs[2][0]);
 }


// **********************************************
// function fcnDelPlayerArmyList()
//
// This function deletes all Players Army List Sheets
// from the Config File
//
// **********************************************

function fcnDelPlayerArmyList(){
  
  Logger.log("Routine: fcnDelPlayerArmyList");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[3][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Template Sheet
  var ssDel = SpreadsheetApp.openById(shtIDs[4][0]);
  var shtTemplate = ssDel.getSheetByName('Template').showSheet();
  
  // Delete Players Army Lists EN
  subDelPlayerSheets(shtIDs[3][0]);
  
  // Delete Players Army Lists FR
  subDelPlayerSheets(shtIDs[4][0]);
}


