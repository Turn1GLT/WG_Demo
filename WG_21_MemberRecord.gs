// **********************************************
// function fcnLogPlayerMatch
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnLogPlayerMatch(ss, shtConfig, logStatusPlyr, MatchData){

  var StatusVal  = logStatusPlyr[0];
  var StatusMsg  = logStatusPlyr[1];
  var PlayerName = logStatusPlyr[2];
    
  // Get Players Sheet
  var shtPlayers = ss.getSheetByName("Players");
  
  // Get Player Log Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var shtPlayerLog = SpreadsheetApp.openById(shtIDs[13][0]).getSheetName(PlayerName);
  
  
return logStatusPlyr;
}