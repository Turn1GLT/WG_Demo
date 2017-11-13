// **********************************************
// function fcnInitLeague()
//
// This function clears all data from sheets  
// to start a new league
//
// **********************************************

function fcnInitLeague(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Open Spreadsheets
  var shtConfig = ss.getSheetByName('Config');
  var shtStandingsEN = ss.getSheetByName('Standings');
  var shtStandingsFR = ss.getSheetByName('Standings');
  var shtMatchRslt   = ss.getSheetByName('Match Results');
  var shtRound;
  var shtResponses   = ss.getSheetByName('Responses');
  var shtResponsesEN = ss.getSheetByName('Responses EN');
  var shtResponsesFR = ss.getSheetByName('Responses FR');
  
  var MaxRowRslt = shtMatchRslt.getMaxRows();
  var MaxColRslt = shtMatchRslt.getMaxColumns();
  var MaxRowRspn = shtResponses.getMaxRows();
  var MaxColRspn = shtResponses.getMaxColumns();
  var MaxRowRspnEN = shtResponsesEN.getMaxRows();
  var MaxColRspnEN = shtResponsesEN.getMaxColumns();
  var MaxRowRspnFR = shtResponsesFR.getMaxRows();
  var MaxColRspnFR = shtResponsesFR.getMaxColumns();
  
  var ColMatchID = shtConfig.getRange(17,9).getValue();
  var ColMatchIDLastVal = shtConfig.getRange(22,9).getValue();
  
  // Clear Data
  shtStandingsEN.getRange(6,2,32,7).clearContent();
  shtStandingsFR.getRange(6,2,32,7).clearContent();
  shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
  shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
  shtResponses.getRange(1,ColMatchIDLastVal).setValue(0);
  shtResponsesEN.getRange(2,1,MaxRowRspnEN-1,MaxColRspnEN).clearContent();
  shtResponsesFR.getRange(2,1,MaxRowRspnFR-1,MaxColRspnFR).clearContent()
  
  // Round Results
  for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
    shtRound = ss.getSheetByName('Round'+RoundNum);
    shtRound.getRange(5,5,32,3).clearContent();
    shtRound.getRange(5,9,32,3).clearContent();
  }

  Logger.log('League Data Cleared');
  
  // Update Standings Copies
  fcnCopyStandingsResults(ss, shtConfig)
  Logger.log('Standings Updated');
  
  // Clear Players DB and Card Pools
  fcnDelPlayerArmyDB();
  fcnDelPlayerArmyList();
  Logger.log('Army DB and Army Lists Cleared');
  
  // Generate Players DB and Card Pools
  fcnGenPlayerArmyDB();
  fcnGenPlayerArmyList();
  Logger.log('Army DB and Army Lists Generated');
}

// **********************************************
// function fcnResetLeagueMatch()
//
// This function clears all Results data but
// does not clear Responses
//
// **********************************************

function fcnResetLeagueMatch(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Open Spreadsheets
  var shtConfig = ss.getSheetByName('Config');
  var shtStandings   = ss.getSheetByName('Standings');
  var shtMatchRslt   = ss.getSheetByName('Match Results');
  var shtRound;
  var shtResponses   = ss.getSheetByName('Responses');
  var shtResponsesEN = ss.getSheetByName('Responses EN');
  var shtResponsesFR = ss.getSheetByName('Responses FR');
  
  var MaxRowRslt = shtMatchRslt.getMaxRows();
  var MaxColRslt = shtMatchRslt.getMaxColumns();
  var MaxRowRspn = shtResponses.getMaxRows();
  var MaxColRspn = shtResponses.getMaxColumns();
  var MaxRowRspnEN = shtResponsesEN.getMaxRows();
  var MaxColRspnEN = shtResponsesEN.getMaxColumns();
  var MaxRowRspnFR = shtResponsesFR.getMaxRows();
  var MaxColRspnFR = shtResponsesFR.getMaxColumns();

  var cfgRoundRound = shtConfig.getRange(13,9).getValue();  
  var ColMatchID = shtConfig.getRange(17,9).getValue();
  var ColMatchIDLastVal = shtConfig.getRange(22,9).getValue();
  
  // Clear Data
  shtStandings.getRange(6,2,32,7).clearContent();
  shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
  shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
  shtResponses.getRange(1,ColMatchIDLastVal).setValue(0);
  shtResponsesEN.getRange(2,ColMatchID,MaxRowRspnEN-1,7).clearContent();
  shtResponsesFR.getRange(2,ColMatchID,MaxRowRspnFR-1,7).clearContent();
  
  // Round Results
  for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
    shtRound = ss.getSheetByName('Round'+RoundNum);
    shtRound.getRange(5,5,32,3).clearContent();
    shtRound.getRange(5,9,32,3).clearContent();
  }

  Logger.log('League Data Cleared');
  
  // Update Standings Copies
  fcnCopyStandingsResults(ss, shtConfig)
  Logger.log('Standings Updated');
 
}



// **********************************************
// function fcnUpdateLinksIDs()
//
// This function updates all sheets Links and IDs  
// in the Config File
//
// **********************************************

function fcnUpdateLinksIDs(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Copy Log Spreadsheet
  var shtCopyLogID = shtConfig.getRange(27, 2).getValue();
  
  if (shtCopyLogID != '') {
    var shtCopyLog = SpreadsheetApp.openById(shtCopyLogID).getSheets()[0];
  
    var CopyLogNbFiles = shtCopyLog.getRange(2, 6).getValue();
    var StartRowCopyLog = 5;
    var StartRowConfigId = 30
    var StartRowConfigLink = 17;
    
    var CopyLogVal = shtCopyLog.getRange(StartRowCopyLog, 2, CopyLogNbFiles, 3).getValues();
    
    var FileName;
    var Link;
    var Formula;
    var ConfigRowID = 'Not Found';
    var ConfigRowLk = 'Not Found';
    
    // Clear Configuration File
    shtConfig.getRange(17,2,6,1).clearContent();
    shtConfig.getRange(30,2,7,1).clearContent();
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master WG League' :
          ConfigRowID = StartRowConfigId + 0;
          ConfigRowLk = 'Not Found'; break;
        case 'Master WG League Army DB' :
          ConfigRowID = StartRowConfigId + 1; 
          ConfigRowLk = 'Not Found'; break;
        case 'Master WG League Army Lists EN' :
          ConfigRowID = StartRowConfigId + 2; 
          ConfigRowLk = StartRowConfigLink + 1; break;
        case 'Master WG League Army Lists FR' :
          ConfigRowID = StartRowConfigId + 3; 
          ConfigRowLk = StartRowConfigLink + 4; break;
        case 'Master WG League Standings EN' :
          ConfigRowID = StartRowConfigId + 4; 
          ConfigRowLk = StartRowConfigLink + 0; break;
        case 'Master WG League Standings FR' :
          ConfigRowID = StartRowConfigId + 5; 
          ConfigRowLk = StartRowConfigLink + 3; break;
        case 'Master WG League Match Reporter EN' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;
        case 'Master WG League Match Reporter FR' :
          ConfigRowID = 'Not Found';
          ConfigRowLk = 'Not Found'; break;	
        default : 
          ConfigRowID = 'Not Found'; 
          ConfigRowLk = 'Not Found'; break;
      }
      
      // Set the Appropriate Sheet ID Value in the Config File
      if (ConfigRowID != 'Not Found') {
        shtConfig.getRange(ConfigRowID, 2).setValue(CopyLogVal[row][2]);
      }
      // Set tthe Appropriate Sheet ID Value in the Config File
      if (ConfigRowLk != 'Not Found') {
        // Opens Spreadsheet by ID
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();
        Logger.log(Link); 
        
        shtConfig.getRange(ConfigRowLk, 2).setValue(Link);
      }
    }
  }
}


// **********************************************
// function fcnCrtPlayerArmyDB()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnCrtPlayerArmyDB(){
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Army DB Spreadsheet
  var ArmyDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssArmyDB = SpreadsheetApp.openById(ArmyDBShtID);
  var shtArmyDB = ssArmyDB.getSheetByName('Template');
  var shtArmyDBNum;
  var NumSheet = ssArmyDB.getNumSheets();
  var SheetsArmyDB = ssArmyDB.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,6).getValue();
  
  var shtPlyrArmyDB;
  var shtPlyrName;
  var SetNum;
  var PlyrRow;

  var shtPlyrArmyName;
  var shtPlyrFaction1;
  var shtPlyrFaction2;
  var shtPlyrWarlord;
  
  // Current Army Values (Power Level or Points
  var shtConfigValueMode = shtConfig.getRange(6,7).getValue();
  var shtConfigArmyValue = shtConfig.getRange(10,7).getValue();
  
  
  // Gets the Player Info from the Player Sheet to Populate the Header
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName     = shtPlayers.getRange(PlyrRow, 2).getValue();
    shtPlyrArmyName = shtPlayers.getRange(PlyrRow, 7).getValue();
    shtPlyrFaction1 = shtPlayers.getRange(PlyrRow, 8).getValue();
    shtPlyrFaction2 = shtPlayers.getRange(PlyrRow, 9).getValue();
    shtPlyrWarlord   = shtPlayers.getRange(PlyrRow,10).getValue();
    
    // Resets the Player Found flag before searching
    PlayerFound = 0;
            
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsArmyDB[sheet].getSheetName();
      Logger.log(SheetName);
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }

    Logger.log('PlayerFound:%s',PlayerFound);
    
    // If Player is not found, add a tab
    if(PlayerFound == 0){
      // INSERTS TAB BEFORE "Card DB" TAB
      ssArmyDB.insertSheet(shtPlyrName, 0, {template: shtArmyDB});
      shtPlyrArmyDB = ssArmyDB.getSheets()[0];
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrArmyDB.getRange(3,3).setValue(shtPlyrName);
      shtPlyrArmyDB.getRange(4,3).setValue(shtPlyrArmyName);
      shtPlyrArmyDB.getRange(5,3).setValue(shtPlyrFaction1 + ', ' + shtPlyrFaction2);
      shtPlyrArmyDB.getRange(6,3).setValue(shtPlyrWarlord);
      if (shtConfigValueMode == 'Power Level') shtPlyrArmyDB.getRange(5,9).setValue(shtConfigArmyValue);
      if (shtConfigValueMode == 'Points')      shtPlyrArmyDB.getRange(5,11).setValue(shtConfigArmyValue);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (shtConfigValueMode == 'Power Level') {
        shtPlyrArmyDB.hideColumns( 6, 3);
        shtPlyrArmyDB.hideColumns(11, 2);
      }
      if (shtConfigValueMode == 'Points') {
        shtPlyrArmyDB.hideColumns( 5, 1);
        shtPlyrArmyDB.hideColumns( 9, 2);
      }
      
    }
  }
  shtPlyrArmyDB = ssArmyDB.getSheets()[0];
  ssArmyDB.setActiveSheet(shtPlyrArmyDB);
}


// **********************************************
// function fcnCrtPlayerArmyList()
//
// This function generates all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnCrtPlayerArmyList(){
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtConfigValueMode = shtConfig.getRange(6,7).getValue();
  
  // Army DB Spreadsheet
  var ArmyDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssArmyDB = SpreadsheetApp.openById(ArmyDBShtID);
  var shtArmyDBTemplate = ssArmyDB.getSheetByName('Template');
  var shtArmyDBMaxRows = shtArmyDBTemplate.getMaxRows();
  var shtArmyDBMaxCols = shtArmyDBTemplate.getMaxColumns();
  Logger.log('shtArmyDBMaxRows: %s',shtArmyDBMaxRows);
  Logger.log('shtArmyDBMaxCols: %s',shtArmyDBMaxCols);
  
  var shtArmyDBVal;
  
  // Army Lists Spreadsheet
  var ArmyListShtEnID = shtConfig.getRange(32, 2).getValue();
  var ArmyListShtFrID = shtConfig.getRange(33, 2).getValue();
  var ssArmyListEn = SpreadsheetApp.openById(ArmyListShtEnID);
  var ssArmyListFr = SpreadsheetApp.openById(ArmyListShtFrID);
  var shtArmyListEn = ssArmyListEn.getSheetByName('Template');
  var shtArmyListFr = ssArmyListFr.getSheetByName('Template');
  var shtArmyListNum;
  
  var NumSheet = ssArmyListEn.getNumSheets();
  var SheetsArmyDB = ssArmyListEn.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,6).getValue();
  
  var shtPlyrArmyListEn;
  var shtPlyrArmyListFr;
  var shtPlyrName;
  var PlyrRow
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    
    // Update the Player Row and Get Player Name from Config File
    PlyrRow = plyr + 2; // 2 is the row where the player list starts
    shtPlyrName = shtPlayers.getRange(PlyrRow, 2).getValue();
    
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NumSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsArmyDB[sheet].getSheetName();
      if (SheetName == shtPlyrName) PlayerFound = 1;
    }
          
    // If Player is not found, add a tab
    if(PlayerFound == 0){
      // INSERTS TAB BEFORE "Card DB" TAB
      
      // English Version
      shtArmyDBVal = ssArmyDB.getSheetByName(shtPlyrName).getRange(1, 1, shtArmyDBMaxRows, shtArmyDBMaxCols).getValues();
      ssArmyListEn.insertSheet(shtPlyrName, 0, {template: shtArmyListEn});
      shtPlyrArmyListEn = ssArmyListEn.getSheets()[0];
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrArmyListEn.getRange(1, 1, shtArmyDBMaxRows, shtArmyDBMaxCols).setValues(shtArmyDBVal);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (shtConfigValueMode == 'Power Level') {
        shtPlyrArmyListEn.hideColumns( 6, 3);
        shtPlyrArmyListEn.hideColumns(11, 2);
      }
      if (shtConfigValueMode == 'Points') {
        shtPlyrArmyListEn.hideColumns( 5, 1);
        shtPlyrArmyListEn.hideColumns( 9, 2);
      }
      
      // French Version
      ssArmyListFr.insertSheet(shtPlyrName, 0, {template: shtArmyListFr});
      shtPlyrArmyListFr = ssArmyListFr.getSheets()[0];
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrArmyListFr.getRange(1, 1, shtArmyDBMaxRows, shtArmyDBMaxCols).setValues(shtArmyDBVal);  
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (shtConfigValueMode == 'Power Level') {
        shtPlyrArmyListFr.hideColumns( 6, 3);
        shtPlyrArmyListFr.hideColumns(11, 2);
      }
      if (shtConfigValueMode == 'Points') {
        shtPlyrArmyListFr.hideColumns( 5, 1);
        shtPlyrArmyListFr.hideColumns( 9, 2);
      }
    }
  }
  // English Version
  shtPlyrArmyListEn = ssArmyListEn.getSheets()[0];
  ssArmyListEn.setActiveSheet(shtPlyrArmyListEn);
  ssArmyListEn.getSheetByName('Template').hideSheet();
  
  // French Version
  shtPlyrArmyListFr = ssArmyListFr.getSheets()[0];
  ssArmyListFr.setActiveSheet(shtPlyrArmyListFr);
  ssArmyListFr.getSheetByName('Template').hideSheet();

}


// **********************************************
// function fcnDelPlayerArmyDB()
//
// This function deletes all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerArmyDB(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card DB Spreadsheet
  var ArmyDBShtID = shtConfig.getRange(31, 2).getValue();
  var ssArmyDB = SpreadsheetApp.openById(ArmyDBShtID);
  var shtTemplate = ssArmyDB.getSheetByName('Template');
  var ssNbSheet = ssArmyDB.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Activates Template Sheet
  ssArmyDB.setActiveSheet(shtTemplate);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    shtCurr = ssArmyDB.getSheets()[0];
    shtCurrName = shtCurr.getName();
    if( shtCurrName != 'Template') ssArmyDB.deleteSheet(shtCurr);
    }
}

// **********************************************
// function fcnDelPlayerArmyList()
//
// This function deletes all Card DB for all 
// players from the Config File
//
// **********************************************

function fcnDelPlayerArmyList(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Card Pool Spreadsheet
  var ArmyListShtIDEn = shtConfig.getRange(32, 2).getValue();
  var ArmyListShtIDFr = shtConfig.getRange(33, 2).getValue();
  var ssArmyListEn = SpreadsheetApp.openById(ArmyListShtIDEn);
  var ssArmyListFr = SpreadsheetApp.openById(ArmyListShtIDFr);
  var shtTemplateEn = ssArmyListEn.getSheetByName('Template');
  var shtTemplateFr = ssArmyListFr.getSheetByName('Template');
  var ssNbSheet = ssArmyListEn.getNumSheets();
  
  // Routine Variables
  var shtCurrEn;
  var shtCurrNameEn;
  var shtCurrFr;
  var shtCurrNameFr;  
  
  // Show Template sheet
  shtTemplateEn.showSheet();
  shtTemplateFr.showSheet();
  
  // Activates Template Sheet
  ssArmyListEn.setActiveSheet(shtTemplateEn);
  ssArmyListFr.setActiveSheet(shtTemplateFr);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    
    // English Version
    shtCurrEn = ssArmyListEn.getSheets()[0];
    shtCurrNameEn = shtCurrEn.getName();
    if( shtCurrNameEn != 'Template') ssArmyListEn.deleteSheet(shtCurrEn);
    
    // French Version   
    shtCurrFr = ssArmyListFr.getSheets()[0];
    shtCurrNameFr = shtCurrFr.getName();
    if( shtCurrNameFr != 'Template') ssArmyListFr.deleteSheet(shtCurrFr);
  }
}

// **********************************************
// function fcnSetupResponseSht()
//
// This function sets up the new Responses sheets 
// and deletes the old ones
//
// **********************************************

function fcnSetupResponseSht(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Responses Sheets
  var shtOldRespEN = ss.getSheetByName('Responses EN');
  var shtOldRespFR = ss.getSheetByName('Responses FR');
  var shtNewRespEN = ss.getSheetByName('New Responses EN');
  var shtNewRespFR = ss.getSheetByName('New Responses FR');
    
  var OldRespMaxCol = shtOldRespEN.getMaxColumns();
  var NewRespMaxRow = shtNewRespEN.getMaxRows();
  var DataLastCol = 7; // Response Data Last Column 
  var ColWidth;
  
  // Copy Header from Old to New sheet - Loop to Copy Value and Format from cell to cell, copy formula (or set) in last cell
  for (var col = 1; col <= OldRespMaxCol; col++){
    // Insert Column if it doesn't exist (col >=24)
    if (col >= DataLastCol && col < OldRespMaxCol){
      shtNewRespEN.insertColumnAfter(col);
      shtNewRespFR.insertColumnAfter(col);
    }
    // Set New Response Sheet Values 
    shtOldRespEN.getRange(1,col).copyTo(shtNewRespEN.getRange(1,col));
    shtOldRespFR.getRange(1,col).copyTo(shtNewRespFR.getRange(1,col));
    ColWidth = shtOldRespEN.getColumnWidth(col);
    shtNewRespEN.setColumnWidth(col,ColWidth);
    shtNewRespFR.setColumnWidth(col,ColWidth);
  }
  // Hides Columns 25, 27-30
//  shtNewRespEN.hideColumns(25);
//  shtNewRespEN.hideColumns(27,4);
//  shtNewRespFR.hideColumns(25);
//  shtNewRespFR.hideColumns(27,4);
  
  // Delete Old Sheets
  ss.deleteSheet(shtOldRespEN);
  ss.deleteSheet(shtOldRespFR);
  
  // Rename New Sheets
  shtNewRespEN.setName('Responses EN');
  shtNewRespFR.setName('Responses FR');

}
