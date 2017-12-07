// **********************************************
// function fcnLogEvntPlayerMatch
//
// This function logs the match to the player's record 
// for this event
//
// **********************************************

function fcnLogEventMatch(ss, shtConfig, cfgEvntParam, logStatusPlyr, MatchData){

  Logger.log("Routine: fcnLogEventMatch: %s",logStatusPlyr[2]);
 
  // Status Values
  var StatusVal  = logStatusPlyr[0];
  var StatusMsg  = logStatusPlyr[1];
  var PlayerName = logStatusPlyr[2];
  
  
  // Event Parameters
  var evntPtsGainedMatch = cfgEvntParam[27][0];
  
  // Get Players Sheet
  var shtPlayers = ss.getSheetByName("Players");
  
  // Get Player Log Spreadsheet
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var ssEvntPlyrRec = SpreadsheetApp.openById(shtIDs[13][0]);
  var shtEvntPlyrRec = ssEvntPlyrRec.getSheetByName(PlayerName);
  var rngRecord = "A4:I4";
  
  // Match Data Variables
  var matchEventEN = MatchData[0][1];
  var matchEventFR = MatchData[0][2];
  var matchGameSys = MatchData[1][1];
  var matchRound =   MatchData[2][0];
  var matchPT1 =     MatchData[3][0];
  var matchPT1Pts =  MatchData[3][1];
  var matchPT2 =     MatchData[4][0];
  var matchPT2Pts =  MatchData[4][1];
  var matchTie =     MatchData[5][0];
  
  // Routine Variables
  var MatchResult = "";
  
  // Get Player Event Record
  var evntRecPlyr = shtEvntPlyrRec.getRange(rngRecord).getValues();
  
  if (evntRecPlyr[0][0] == '') evntRecPlyr[0][0] = 0; // MP
  if (evntRecPlyr[0][1] == '') evntRecPlyr[0][1] = 0; // Win
  if (evntRecPlyr[0][2] == '') evntRecPlyr[0][2] = 0; // Loss
  if (evntRecPlyr[0][3] == '') evntRecPlyr[0][3] = 0; // Tie
  if (evntRecPlyr[0][4] == '') evntRecPlyr[0][4] = 0; // Points Scored
  if (evntRecPlyr[0][4] == '') evntRecPlyr[0][5] = 0; // Points Allowed
  if (evntRecPlyr[0][5] == '') evntRecPlyr[0][6] = 0; // Win %
  if (evntRecPlyr[0][5] == '') evntRecPlyr[0][7] = 0; // Pts Scored  / Match
  if (evntRecPlyr[0][5] == '') evntRecPlyr[0][8] = 0; // Pts Allowed / Match
  
  // Checks if Match is a Tie
  if((evntPtsGainedMatch == "Enabled" && matchPT1Pts == matchPT2Pts) || (matchTie == "Yes" || matchTie == "Oui")) MatchResult = "Tie";
  
  // Update Player Record
  
  // Update Player Matches Played
  evntRecPlyr[0][0] = evntRecPlyr[0][0] + 1;
  
  // If Points Gained in Match Options is not Used
  if(MatchResult == "" && evntPtsGainedMatch == "Disabled"){
    // Update Player Wins
    if(PlayerName == matchPT1) {
      evntRecPlyr[0][1] = evntRecPlyr[0][1] + 1;
      MatchResult = "Win";
    }
    // Update Player Loss
    if(PlayerName == matchPT2) {
      evntRecPlyr[0][2] = evntRecPlyr[0][2] + 1;
      MatchResult = "Loss";
    }
  }

  // If Points Gained in Match Options is Used
  if(MatchResult == "" && evntPtsGainedMatch == "Enabled"){
    // Update Player Wins
    if((PlayerName == matchPT1 && matchPT1Pts > matchPT2Pts) || (PlayerName == matchPT2 && matchPT2Pts > matchPT1Pts)) {
      evntRecPlyr[0][1] = evntRecPlyr[0][1] + 1;
      MatchResult = "Win";
    }
    // Update Player Loss
    if((PlayerName == matchPT1 && matchPT1Pts < matchPT2Pts) || (PlayerName == matchPT2 && matchPT2Pts < matchPT1Pts)) {
      evntRecPlyr[0][2] = evntRecPlyr[0][2] + 1;
      MatchResult = "Loss";
    }
  }
  
  // Update Player Tie
  if(MatchResult == "Tie" || (evntPtsGainedMatch == "Enabled" && matchPT1Pts == matchPT2Pts)) {
    evntRecPlyr[0][3] = evntRecPlyr[0][3] + 1;
  }
  
  // Update Points
  // If Player 1
  if(PlayerName == matchPT1 && evntPtsGainedMatch == "Enabled") {
    evntRecPlyr[0][4] = evntRecPlyr[0][4] + matchPT1Pts; // Points Scored
    evntRecPlyr[0][5] = evntRecPlyr[0][5] + matchPT2Pts; // Points Allowed
  }
  // If Player 2
  if(PlayerName == matchPT2 && evntPtsGainedMatch == "Enabled") {
    evntRecPlyr[0][4] = evntRecPlyr[0][4] + matchPT2Pts; // Points Scored
    evntRecPlyr[0][5] = evntRecPlyr[0][5] + matchPT1Pts; // Points Allowed
  }
  
  // Update Win Percentage
  if(evntRecPlyr[0][0] > 0) evntRecPlyr[0][6] = evntRecPlyr[0][1] / evntRecPlyr[0][0];

  // Update Points Scored / Match
  if(evntRecPlyr[0][0] > 0) evntRecPlyr[0][7] = evntRecPlyr[0][4] / evntRecPlyr[0][0];
  
  // Update Points Allowed / Match
  if(evntRecPlyr[0][0] > 0) evntRecPlyr[0][8] = evntRecPlyr[0][5] / evntRecPlyr[0][0];  
  
  // Post New Record
  shtEvntPlyrRec.getRange(rngRecord).setValues(evntRecPlyr);
  
  // UPDATE PLAYER HISTORY

  // Get Player Language Preference
  var Language = shtEvntPlyrRec.getRange(6,1).getValue();  
  if(Language == "Event")     Language = "English";
  if(Language == "Événement") Language = "Français";
  
  // Create Values Array (1,6)
  var values = subCreateArray(1,9); 
  // [0][0]= Event Name Cell 1
  // [0][1]= Event Name Cell 2
  // [0][2]= Game System
  // [0][3]= Round
  // [0][4]= Result (Win, Loss, Tie)
  // [0][5]= Played Against Cell 1
  // [0][6]= Played Against Cell 2
  // [0][7]= Points Scored
  // [0][8]= Points Allowed
  
  // Event Name
  if(Language == "English")  values[0][0]= matchEventEN;// Event Name
  if(Language == "Français") values[0][0]= matchEventFR;// Event Name
  
  // Game System
  values[0][2]= matchGameSys;
  
  // Round
  values[0][3]= matchRound;
  
  // Result
  switch(MatchResult){
    case "Win": {
      // Result (Win, Loss, Tie)
      if(Language == "English")  values[0][4]= "Win";
      if(Language == "Français") values[0][4]= "Victoire";
      break;
    }
    case "Loss": {
      // Result (Win, Loss, Tie)
      if(Language == "English")  values[0][4]= "Loss";
      if(Language == "Français") values[0][4]= "Défaite";
      break;
    }
    case "Tie": {
      // Result (Win, Loss, Tie)
      if(Language == "English")  values[0][4]= "Tie";
      if(Language == "Français") values[0][4]= "Nulle";
      break;
    }
  }
  // If Player is Player 1
  if(PlayerName == matchPT1){
    // Played Against
    values[0][5]= matchPT2;
    // Points Scored
    values[0][7]= matchPT1Pts;
    // Points Allowed
    values[0][8]= matchPT2Pts;
  }

  // If Player is Player 2
  if(PlayerName == matchPT2){
    // Played Against
    values[0][5]= matchPT1;
    // Points Scored
    values[0][7]= matchPT2Pts;
    // Points Allowed
    values[0][8]= matchPT1Pts;
  }
      
  // Get Last Row
  var LastRow = shtEvntPlyrRec.getMaxRows();
   
  // Post Match Data to Last Row
  shtEvntPlyrRec.getRange(LastRow,1,1,9).setValues(values);
  
  // Add Row for next Log and Merge Columns 1-2 and 6-7
  shtEvntPlyrRec.insertRowAfter(LastRow);
  shtEvntPlyrRec.getRange(LastRow+1, 1, 1, 2).merge();
  shtEvntPlyrRec.getRange(LastRow+1, 6, 1, 2).merge();
  
  logStatusPlyr[0]= 1;
  logStatusPlyr[1]= "Event Player Record Updated";
      
  return logStatusPlyr;
}