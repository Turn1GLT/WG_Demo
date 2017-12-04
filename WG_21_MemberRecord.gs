// **********************************************
// function fcnLogPlayerMatch
//
// This function logs the match to the player's record 
// for this event ant to the player's profile 
// if the Member Option is enabled
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