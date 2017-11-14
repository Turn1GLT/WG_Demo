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
    var rowCopyLogStart = 5;
    var rowShtIdStart = 4;
    var colShtId = 7;
    var rowShtUrlStart = 4;
    var colShtUrl = 11;
    
    var CopyLogVal = shtCopyLog.getRange(rowCopyLogStart, 2, CopyLogNbFiles, 3).getValues();
    // [0]= Sheet Name, [1]= Sheet URL, [2]= Sheet ID
    
    var FileName;
    var Link;
    var Formula;
    var rowID = 'Not Found';
    var rowUrl = 'Not Found';
    
    // Clear Sheet IDs
    shtConfig.getRange(rowShtIdStart, colShtId,20,1).clearContent();
    // Clear Sheet URLs
    shtConfig.getRange(rowShtUrlStart,colShtUrl,20,1).clearContent();
    
    // Loop through all Copied Sheets and get their Link and ID
    for (var row = 0; row < CopyLogNbFiles; row++){
      // Get File Name
      FileName = CopyLogVal[row][0];
      
      switch(FileName){
        case 'Master WG League' :
          rowID = rowShtIdStart + 0;
          rowUrl = rowShtUrlStart + 0; break;
        case 'Master WG Log' :
          rowID = rowShtIdStart + 1;
          rowUrl = rowShtUrlStart + 1; break;
        case 'Master WG League Army DB' :
          rowID = rowShtIdStart + 2; 
          rowUrl = rowShtUrlStart + 2; break;
        case 'Master WG League Army Lists EN' :
          rowID = rowShtIdStart + 3; 
          rowUrl = rowShtUrlStart + 3; break;
        case 'Master WG League Army Lists FR' :
          rowID = rowShtIdStart + 4; 
          rowUrl = rowShtUrlStart + 4; break;
        case 'Master WG League Standings EN' :
          rowID = rowShtIdStart + 5; 
          rowUrl = rowShtUrlStart + 5; break;
        case 'Master WG League Standings FR' :
          rowID = rowShtIdStart + 6; 
          rowUrl = rowShtUrlStart + 6; break;	
        default : 
          rowID = 'Not Found'; 
          rowUrl = 'Not Found'; break;
      }
      
      // Set the Appropriate Sheet ID Value in the Config File
      if (rowID != 'Not Found') {
        shtConfig.getRange(rowID, colShtId).setValue(CopyLogVal[row][2]);
      }
      // Set tthe Appropriate Sheet ID Value in the Config File
      if (rowUrl != 'Not Found') {
        // Opens Spreadsheet by ID
        Link = SpreadsheetApp.openById(CopyLogVal[row][2]).getUrl();        
        shtConfig.getRange(rowUrl, colShtUrl).setValue(Link);
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
  
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
    
    // Columns from Config File
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    var colRndWin = cfgColRndSht[3][0];
    var colRndMatchLoc = cfgColRndSht[8][0];
    
    // Sheets
    var shtStandings = ss.getSheetByName('Standings');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtRound;
    var shtResponses   = ss.getSheetByName('Responses');
    var shtResponsesEN = ss.getSheetByName('Responses EN');
    var shtResponsesFR = ss.getSheetByName('Responses FR');
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtResponsesEN.getMaxRows();
    var MaxColRspnEN = shtResponsesEN.getMaxColumns();
    var MaxRowRspnFR = shtResponsesFR.getMaxRows();
    var MaxColRspnFR = shtResponsesFR.getMaxColumns();
    
    // Clear Data
    // Standings
    shtStandings.getRange(6,2,MaxRowStdg-5,MaxColStdg-1).clearContent();
    // Match Results (does not clear the last column)
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
    // Responses
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtResponsesEN.getRange(2,1,MaxRowRspnEN-1,MaxColRspnEN).clearContent();
    shtResponsesFR.getRange(2,1,MaxRowRspnFR-1,MaxColRspnFR).clearContent()
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndWin,32,4).clearContent();
      shtRound.getRange(5,colRndMatchLoc,32,4).clearContent();
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
  var msg = "Click OK to clear all "+ cfgEventType +"match results";
  var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
    
  // If Confirmed (OK), Initialize all League Data
  if(uiResponse == "OK"){
    
    // Config Sheet to get options
    var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
    var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
    var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
    
    // Columns from Config File
    var colRspMatchID = cfgColRspSht[1][0];
    var colRspMatchIDLastVal = cfgColRspSht[6][0];
    var colRndWin = cfgColRndSht[3][0];
    var colRndMatchLoc = cfgColRndSht[8][0];
    
    // Sheets
    var shtStandings   = ss.getSheetByName('Standings');
    var shtMatchRslt   = ss.getSheetByName('Match Results');
    var shtRound;
    var shtResponses   = ss.getSheetByName('Responses');
    var shtResponsesEN = ss.getSheetByName('Responses EN');
    var shtResponsesFR = ss.getSheetByName('Responses FR');
    
    // Max Rows / Columns
    var MaxRowStdg = shtStandings.getMaxRows();
    var MaxColStdg = shtStandings.getMaxColumns();
    var MaxRowRslt = shtMatchRslt.getMaxRows();
    var MaxColRslt = shtMatchRslt.getMaxColumns();
    var MaxRowRspn = shtResponses.getMaxRows();
    var MaxColRspn = shtResponses.getMaxColumns();
    var MaxRowRspnEN = shtResponsesEN.getMaxRows();
    var MaxColRspnEN = shtResponsesEN.getMaxColumns();
    var MaxRowRspnFR = shtResponsesFR.getMaxRows();
    var MaxColRspnFR = shtResponsesFR.getMaxColumns();
    
    // Clear Data
    shtStandings.getRange(6,2,32,7).clearContent();
    shtMatchRslt.getRange(6,2,MaxRowRslt-5,MaxColRslt-2).clearContent();
    shtResponses.getRange(2,1,MaxRowRspn-1,MaxColRspn).clearContent();
    shtResponses.getRange(1,colRspMatchIDLastVal).setValue(0);
    shtResponsesEN.getRange(2,colRspMatchID,MaxRowRspnEN-1,MaxColRspnEN-colRspMatchID+1).clearContent();
    shtResponsesFR.getRange(2,colRspMatchID,MaxRowRspnFR-1,MaxColRspnFR-colRspMatchID+1).clearContent();
    
    // Round Results
    for (var RoundNum = 1; RoundNum <= 8; RoundNum++){
      shtRound = ss.getSheetByName('Round'+RoundNum);
      shtRound.getRange(5,colRndWin,32,4).clearContent();
      shtRound.getRange(5,colRndMatchLoc,32,4).clearContent();
    }
    
    Logger.log('League Data Cleared');
    
    // Update Standings Copies
    fcnCopyStandingsResults(ss, shtConfig)
    Logger.log('Standings Updated');
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
  
  Logger.log("Routine: fcnCrtPlayerArmyDB");
  
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Configuration Data
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange(5,28,16,1).getValues();
  
  // Column Values
  var colShtPlyrName = cfgColShtPlyr[2][0];
  var colShtPlyrTeam = cfgColShtPlyr[5][0];
  var colShtPlyrArmy = cfgColShtPlyr[7][0];
  var colShtPlyrFaction1 = cfgColShtPlyr[8][0];
  var colShtPlyrFaction2 = cfgColShtPlyr[9][0];
  var colShtPlyrWarlord = cfgColShtPlyr[10][0];
  
  // Current Army Values (Power Level or Points
  var armyBldRatingMode = cfgArmyBuild[0][0];
  var armyBldArmyValue = cfgArmyBuild[1][0];
  
  // Army DB Spreadsheet
  var ssArmyDB = SpreadsheetApp.openById(shtIDs[2][0]);
  var shtTemplate = ssArmyDB.getSheetByName('Template');
  var NbSheet = ssArmyDB.getNumSheets();
  var SheetsArmyDB = ssArmyDB.getSheets();
  var SheetName;
  var PlayerFound = 0;

  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayerNames = shtPlayers.getRange(3,colShtPlyrName, NbPlayers, 1).getValues();
    
  var shtPlyrArmyDB;
  var PlyrName;
  var PlyrArmy;
  var PlyrFaction1;
  var PlyrFaction2;
  var PlyrWarlord;
  
  // Loops through each player starting from the last
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerNames[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;            
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsArmyDB[sheet].getSheetName();
      Logger.log(SheetName);
      if (SheetName == PlyrName) PlayerFound = 1;
    }

    Logger.log('PlayerFound:%s',PlayerFound);
    
    // If Player is not found, create a tab with the player's name
    if(PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssArmyDB.getNumSheets();
      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      ssArmyDB.insertSheet(PlyrName, NbSheet-2, {template: shtTemplate});
      shtPlyrArmyDB = ssArmyDB.getSheetByName(PlyrName);
      
      // Opens the new sheet and modify appropriate data (Player Name, Header)
      shtPlyrArmyDB.getRange(3,3).setValue(PlyrName);
      shtPlyrArmyDB.getRange(4,3).setValue(PlyrArmy);
      if(PlyrFaction2 == '') shtPlyrArmyDB.getRange(5,3).setValue(PlyrFaction1);
      if(PlyrFaction2 != '') shtPlyrArmyDB.getRange(5,3).setValue(PlyrFaction1 + ', ' + PlyrFaction2);
      shtPlyrArmyDB.getRange(6,3).setValue(PlyrWarlord);
      //
      if (armyBldRatingMode == 'Power Level') shtPlyrArmyDB.getRange(5,9).setValue(armyBldArmyValue);
      if (armyBldRatingMode == 'Points')      shtPlyrArmyDB.getRange(5,11).setValue(armyBldArmyValue);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyrArmyDB.hideColumns( 6, 3);
        shtPlyrArmyDB.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
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
  
  Logger.log("Routine: fcnCrtPlayerArmyDB");
    
  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Configuration Data
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  var cfgColShtPlyr = shtConfig.getRange(5,28,16,1).getValues();
  
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
  var shtTemplateEN = ssArmyListEn.getSheetByName('Template');
  var shtTemplateFR = ssArmyListFr.getSheetByName('Template');
  var shtArmyListNum;
  
  var NbSheet = ssArmyListEN.getNumSheets();
  var SheetsArmyDB = ssArmyListEN.getSheets();
  var SheetName;
  var PlayerFound = 0;
  
  // Players 
  var shtPlayers = ss.getSheetByName('Players'); 
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayerNames = shtPlayers.getRange(3,colShtPlyrName, NbPlayers, 1).getValues();
    
  var shtPlyrArmyListEN;
  var shtPlyrArmyListFR;
  var PlyrName;
  
  // Loops through each player starting from the first
  for (var plyr = NbPlayers; plyr > 0; plyr--){
    // Gets the Player Name 
    PlyrName = PlayerNames[plyr][0];
    // Resets the Player Found flag before searching
    PlayerFound = 0;
    // Look if player exists, if yes, skip, if not, create player
    for(var sheet = NbSheet - 1; sheet >= 0; sheet --){
      SheetName = SheetsArmyDB[sheet].getSheetName();
      if (SheetName == PlyrName) PlayerFound = 1;
    }
          
    // If Player is not found, add a tab
    if(PlayerFound == 0){
      // Get the Template sheet index
      NbSheet = ssArmyListEN.getNumSheets();

      // Inserts Sheet before Template (Last Sheet in Spreadsheet)
      // English Version
      ssArmyListEN.insertSheet(PlyrName, NbSheet-2, {template: shtTemplateEN});
      shtPlyrArmyListEN = ssArmyListEN.getSheetByName(PlyrName);
      
      // French Version
      ssArmyListFR.insertSheet(PlyrName, NbSheet-2, {template: shtTemplateFR});
      shtPlyrArmyListFR = ssArmyListFR.getSheetByName(PlyrName);
            
      // Get Player Army DB Values
      fcnCopyArmyDBtoArmyList(shtConfig,PlyrName);
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyrArmyListEN.hideColumns( 6, 3);
        shtPlyrArmyListEN.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
        shtPlyrArmyListEN.hideColumns( 5, 1);
        shtPlyrArmyListEN.hideColumns( 9, 2);
      }
       
      
      // Hides the unused columns according to the Army Value (Power Level or Points)
      if (armyBldRatingMode == 'Power Level') {
        shtPlyrArmyListFR.hideColumns( 6, 3);
        shtPlyrArmyListFR.hideColumns(11, 2);
      }
      if (armyBldRatingMode == 'Points') {
        shtPlyrArmyListFR.hideColumns( 5, 1);
        shtPlyrArmyListFR.hideColumns( 9, 2);
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
  var shtPlyrArmyDB = ssArmyDB.getSheetByName(PlyrName);
  var shtPlyrArmyListEN = ssArmyListEN.getSheetByName(PlyrName);
  var shtPlyrArmyListFR = ssArmyListFR.getSheetByName(PlyrName);
  
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
  var PlyrHdr = shtPlyrArmyDB.getRange(rngPlyrHdr).getValues();
  shtPlyrArmyListEN.getRange(rngPlyrHdr).setValues(PlyrHdr);
  shtPlyrArmyListFR.getRange(rngPlyrHdr).setValues(PlyrHdr);
  
  // Army Rating Level
  var Rating = shtPlyrArmyDB.getRange(rngRating).getValues();
  shtPlyrArmyListEN.getRange(rngRating).setValues(Rating);
  shtPlyrArmyListFR.getRange(rngRating).setValues(Rating);
  
  // Detachment 1
  var Dtch1Hdr = shtPlyrArmyDB.getRange(rngDtch1Hdr).getValues();
  shtPlyrArmyListEN.getRange(rngDtch1Hdr).setValues(Dtch1Hdr);
  shtPlyrArmyListFR.getRange(rngDtch1Hdr).setValues(Dtch1Hdr);
  
  var Dtch1Data = shtPlyrArmyDB.getRange(rngDtch1Data).getValues();
  shtPlyrArmyListEN.getRange(rngDtch1Data).setValues(Dtch1Data);
  shtPlyrArmyListFR.getRange(rngDtch1Data).setValues(Dtch1Data);
    
  // Detachment 2
  var Dtch2Hdr = shtPlyrArmyDB.getRange(rngDtch2Hdr).getValues();
  shtPlyrArmyListEN.getRange(rngDtch2Hdr).setValues(Dtch2Hdr);
  shtPlyrArmyListFR.getRange(rngDtch2Hdr).setValues(Dtch2Hdr);
  
  var Dtch2Data = shtPlyrArmyDB.getRange(rngDtch2Data).getValues();
  shtPlyrArmyListEN.getRange(rngDtch2Data).setValues(Dtch2Data);
  shtPlyrArmyListFR.getRange(rngDtch2Data).setValues(Dtch2Data);
  
  // Detachment 3
  var Dtch3Hdr = shtPlyrArmyDB.getRange(rngDtch3Hdr).getValues();
  shtPlyrArmyListEN.getRange(rngDtch3Hdr).setValues(Dtch3Hdr);
  shtPlyrArmyListFR.getRange(rngDtch3Hdr).setValues(Dtch3Hdr);
  
  var Dtch3Data = shtPlyrArmyDB.getRange(rngDtch3Data).getValues();
  shtPlyrArmyListEN.getRange(rngDtch3Data).setValues(Dtch3Data);
  shtPlyrArmyListFR.getRange(rngDtch3Data).setValues(Dtch3Data);
}


// **********************************************
// function fcnDelPlayerArmyDB()
//
// This function deletes all Card DB for all 
// players from the Config File
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
  
  // Army DB Spreadsheet
  var ssArmyDB = SpreadsheetApp.openById(shtIDs[2][0]); 
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
  
  Logger.log("Routine: fcnDelPlayerArmyList");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Config Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Army Lists Spreadsheet
  var ssArmyListEN = SpreadsheetApp.openById(shtIDs[3][0]);
  var ssArmyListFR = SpreadsheetApp.openById(shtIDs[4][0]);
  var shtTemplateEN = ssArmyListEN.getSheetByName('Template');
  var shtTemplateFR = ssArmyListFR.getSheetByName('Template');
  var ssNbSheet = ssArmyListEN.getNumSheets();
  
  // Routine Variables
  var shtCurrEN;
  var shtCurrNameEN;
  var shtCurrFR;
  var shtCurrNameFR;  
  
  // Show Template sheet
  shtTemplateEN.showSheet();
  shtTemplateFR.showSheet();
  
  // Activates Template Sheet
  ssArmyListEN.setActiveSheet(shtTemplateEN);
  ssArmyListFR.setActiveSheet(shtTemplateFR);
  
  for (var sht = 0; sht < ssNbSheet - 1; sht ++){
    
    // English Version
    shtCurrEN = ssArmyListEN.getSheets()[0];
    shtCurrNameEN = shtCurrEN.getName();
    if( shtCurrNameEN != 'Template') ssArmyListEN.deleteSheet(shtCurrEN);
    
    // French Version   
    shtCurrFR = ssArmyListFR.getSheets()[0];
    shtCurrNameFR = shtCurrFR.getName();
    if( shtCurrNameFR != 'Template') ssArmyListFR.deleteSheet(shtCurrFR);
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
  
  Logger.log("Routine: fcnSetupResponseSht");

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Configuration Sheet
  var shtConfig = ss.getSheetByName('Config');
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  
  // Open Responses Sheets
  var shtOldRespEN = ss.getSheetByName('Responses EN');
  var shtOldRespFR = ss.getSheetByName('Responses FR');
  var shtNewRespEN = ss.getSheetByName('New Responses EN');
  var shtNewRespFR = ss.getSheetByName('New Responses FR');
    
  var OldRespMaxCol = shtOldRespEN.getMaxColumns();
  var NewRespMaxRow = shtNewRespEN.getMaxRows();
  var ColWidth;
  
  // Columns Values and Parameters
  var RspnDataInputs = cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var colMatchID = cfgColRspSht[1][0];
  var colPrcsd = cfgColRspSht[2][0];
  var colDataConflict = cfgColRspSht[3][0];
  var colStatus = cfgColRspSht[4][0];
  var colStatusMsg = cfgColRspSht[5][0];
  var colMatchIDLastVal = cfgColRspSht[6][0];
  var colNextEmptyRow = cfgColRspSht[7][0];
  var colNbUnprcsdEntries = cfgColRspSht[8][0];
  
  // Copy Header from Old to New sheet - Loop to Copy Value and Format from cell to cell, copy formula (or set) in last cell
  for (var col = 1; col <= OldRespMaxCol; col++){
    // Insert Column if it doesn't exist
    if (col >= colMatchID-1 && col < OldRespMaxCol){
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
  
  // Hides Columns 
  shtNewRespEN.hideColumn(colMatchID);
  shtNewRespEN.hideColumn(colDataConflict);
  shtNewRespEN.hideColumn(colStatus);
  shtNewRespEN.hideColumn(colStatusMsg);
  shtNewRespEN.hideColumn(colMatchIDLastVal);
  
  shtNewRespFR.hideColumn(colMatchID);
  shtNewRespFR.hideColumn(colDataConflict);
  shtNewRespFR.hideColumn(colStatus);
  shtNewRespFR.hideColumn(colStatusMsg);
  shtNewRespFR.hideColumn(colMatchIDLastVal);
  
  // Delete Old Sheets
  ss.deleteSheet(shtOldRespEN);
  ss.deleteSheet(shtOldRespFR);
  
  // Rename New Sheets
  shtNewRespEN.setName('Responses EN');
  shtNewRespFR.setName('Responses FR');

}
