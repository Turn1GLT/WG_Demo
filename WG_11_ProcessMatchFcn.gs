// **********************************************
// function fcnPostMatchResultsWG()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Round X tab
//
// **********************************************

function fcnPostMatchResultsWG(ss, shtConfig, ResponseData, MatchingRspnData, MatchID, MatchData) {
  
  Logger.log("Routine: fcnPostMatchResultsWG");
  
  // Config Sheet to get options
  var cfgEvntParam =    shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColRspSht =    shtConfig.getRange( 4,18,16,1).getValues();
  var cfgColRndSht =    shtConfig.getRange( 4,21,16,1).getValues();
  var cfgExecData  =    shtConfig.getRange( 4,24,16,1).getValues();
  var cfgColMatchRep =  shtConfig.getRange( 4,31,20,1).getValues();
  var cfgColMatchRslt = shtConfig.getRange(21,18,32,1).getValues();
  
  // Code Execution Options
  var exeDualSubmission =      cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var exePlyrMatchValidation = cfgExecData[2][0];
  var exePostRoundResult =     cfgExecData[8][0];
  
  // Event Parameters
  var evntGameType =       cfgEvntParam[4][0];
  var evntFormat =         cfgEvntParam[9][0];
  var evntLocationBonus =  cfgEvntParam[23][0];
  var evntPtsGainedMatch = cfgEvntParam[27][0];
  var evntTiePossible =    cfgEvntParam[31][0];
  
  // Cumulative Results sheet variables
  var shtCumul;
  var BalanceBonusLosr;
  
  // Match Results Sheet Values
  var shtRslt = ss.getSheetByName('Match Results');
  var shtRsltMaxRows = shtRslt.getMaxRows();
  var shtRsltMaxCol = shtRslt.getMaxColumns();
  var RsltLastResultRowRng = shtRslt.getRange(3, 5);
  var RsltLastResultRow = RsltLastResultRowRng.getValue() + 1;
  var RsltRng = shtRslt.getRange(RsltLastResultRow, 1, 1, shtRsltMaxCol);
  var ResultData = RsltRng.getValues();
  
  // Routine Variables
  var MatchValidPT1 = new Array(2); // [0] = Status, [1] = Matches Played by Player/Team 1 used for Error Validation
  var MatchValidPT2 = new Array(2); // [0] = Status, [1] = Matches Played by Player/Team 2 used for Error Validation
  var RsltPlyrData1;
  var RsltPlyrData2;
  
  // Column Values for Data in Response Sheet
  var colArrRspnPwd =     cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =   cfgColMatchRep[ 2][0]-1;
  var colArrRspnPlyr1 =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnTeam1 =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnPts1 =    cfgColMatchRep[ 5][0]-1;
  var colArrRspnPlyr2 =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnTeam2 =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnPts2 =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =     cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =     cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub = cfgColMatchRep[19][0]-1;

  // Column Values for Data in Match Result Sheet
  var colArrRsltResultID = cfgColMatchRslt[ 0][0]-1;
  var colArrRsltMatchCnt = cfgColMatchRslt[ 1][0]-1;
  var colArrRsltMatchID =  cfgColMatchRslt[ 2][0]-1;
  var colArrRsltRound =    cfgColMatchRslt[ 3][0]-1;
  var colArrRsltPT1 =      cfgColMatchRslt[ 4][0]-1;
  var colArrRsltPts1 =     cfgColMatchRslt[ 5][0]-1;
  var colArrRsltPT2 =      cfgColMatchRslt[ 6][0]-1;
  var colArrRsltPts2 =     cfgColMatchRslt[ 7][0]-1;
  var colArrRsltTie =      cfgColMatchRslt[ 8][0]-1;
  var colArrRsltLoc =      cfgColMatchRslt[ 9][0]-1;
  var colArrRsltBal =      cfgColMatchRslt[10][0]-1;  
  
  // Column Values for Rounds Sheets
  var colRndBalBonus = cfgColRndSht[10][0];
  
  var MatchPostedStatus = 0;
  var PT1;
  var PT2;
  
  // Sets which Data set is Player A and Player B. Data[0][1] = Player who posted the data
  if (exeDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][2]) {
    RsltPlyrData1 = ResponseData;
    RsltPlyrData2 = MatchingRspnData;
  }
  
  if (exeDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][3]) {
    RsltPlyrData1 = ResponseData;
    RsltPlyrData2 = MatchingRspnData;
  }
  
  // Copy Players Data
  ResultData[0][colArrRsltRound] = ResponseData[0][colArrRspnRound];  // Round Number
  
  // If Single Event
  if(evntFormat == "Single"){
    ResultData[0][colArrRsltPT1] = ResponseData[0][colArrRspnPlyr1];  // Player 1
    ResultData[0][colArrRsltPT2] = ResponseData[0][colArrRspnPlyr2];  // Player 2
  }
  // If Team Event
  if(evntFormat == "Team"){
    ResultData[0][colArrRsltPT1] = ResponseData[0][colArrRspnTeam1];  // Team 1
    ResultData[0][colArrRsltPT2] = ResponseData[0][colArrRspnTeam2];  // Team 2
  }
  
  // Points
  if(evntPtsGainedMatch == "Enabled"){
    ResultData[0][colArrRsltPts1] = ResponseData[0][colArrRspnPts1]; // Player 1 Points
    ResultData[0][colArrRsltPts2] = ResponseData[0][colArrRspnPts2]; // Player 2 Points
  }
  else{
    ResultData[0][colArrRsltPts1] = "-"; // Winning Player Points
    ResultData[0][colArrRsltPts2] = "-"; // Losing Player Points
  }
  
  // Player/Team 1 and 2 Names
  PT1 = ResultData[0][colArrRsltPT1];
  PT2 = ResultData[0][colArrRsltPT2];
  
  if(evntTiePossible == "Enabled")   ResultData[0][colArrRsltTie] = ResponseData[0][colArrRspnTie];  // Game is Tie
  if(evntLocationBonus == "Enabled") ResultData[0][colArrRsltLoc] = ResponseData[0][colArrRspnLoc];  // Location
  
  // If option is enabled, Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  if (exePlyrMatchValidation == 'Enabled'){
    // Call subroutine to check if players match are valid
    // subPlayerMatchValidation(ss, shtConfig, ParticipantName, MatchValidation)
    MatchValidPT1 = subPlayerMatchValidation(ss, shtConfig, PT1, MatchValidPT1);
    Logger.log('%s Match Validation: %s',PT1, MatchValidPT1[0]);
    
    MatchValidPT2 = subPlayerMatchValidation(ss, shtConfig, PT2, MatchValidPT2);
    Logger.log('%s Match Validation: %s',PT2, MatchValidPT2[0]);
  }

  // If option is disabled, Consider Matches are valid
  if (exePlyrMatchValidation == 'Disabled'){
    MatchValidPT1[0] = 1;
    MatchValidPT2[0] = 1;
  }
  
  // If both players have played a valid match
  if (MatchValidPT1[0] == 1 && MatchValidPT2[0] == 1){
    
    // Copies Match ID
    ResultData[0][colArrRsltMatchID] = MatchID; // Match ID
    
    // Sets Data in Match Result Tab
    ResultData[0][colArrRsltMatchCnt] = '= if(INDIRECT("R[0]C[1]",FALSE)<>"",1,"")';
    RsltRng.setValues(ResultData);
    
    // Update the Match Posted Status
    MatchPostedStatus = 1;
    
    // Post Results in Appropriate Round Number for Both Players
    if(exePostRoundResult == "Enabled") {
      // DataPostedLosr is an Array with [0]=Post Status (1=Success) [1]=Loser Row [2]=Power Level Column
      MatchData = fcnPostRoundResultWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRslt, ResultData, MatchData);
    }
  }
  
  // If Match Validation was not successful, generate Error Status
  
  // returns Error that Winning Player is Eliminated from the League
  if (MatchValidPT1[0] == -1 && MatchValidPT2[0] == 1)  MatchPostedStatus = -11;
  
  // returns Error that Winning Player has played too many matches
  if (MatchValidPT1[0] == -2 && MatchValidPT2[0] == 1)  MatchPostedStatus = -12;  
  
  // returns Error that Losing Player is Eliminated from the League
  if (MatchValidPT2[0] == -1 && MatchValidPT1[0] == 1)  MatchPostedStatus = -21;
  
  // returns Error that Losing Player has played too many matches
  if (MatchValidPT2[0] == -2 && MatchValidPT1[0] == 1)  MatchPostedStatus = -22;
  
  // returns Error that Both Players are Eliminated from the League
  if (MatchValidPT1[0] == -1 && MatchValidPT2[0] == -1) MatchPostedStatus = -31;
  
  // returns Error that Winning Player is Eliminated from the League and Losing Player has played too many matches
  if (MatchValidPT1[0] == -1 && MatchValidPT2[0] == -2) MatchPostedStatus = -32;

  // returns Error that Winning Player has player too many matches and Losing Player is Eliminated from the League
  if (MatchValidPT1[0] == -2 && MatchValidPT2[0] == -1) MatchPostedStatus = -33;
  
  // returns Error that Both Players have played too many matches
  if (MatchValidPT1[0] == -2 && MatchValidPT2[0] == -2) MatchPostedStatus = -34;
  
  // Populates Match Data for Main Routine
  MatchData[4][3] = BalanceBonusLosr;     // Player/Team 2 Balance Bonus Value
  MatchData[25][0] = MatchPostedStatus;
  
  return MatchData;
}


// **********************************************
// function fcnPostRoundResultWG()
//
// Once the Match Data has been posted in the 
// Match Results Tab, the Round X results are updated
// for each player
//
// **********************************************

function fcnPostRoundResultWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRslt, ResultData, MatchData) {
  
  Logger.log("Routine: fcnPostResultRoundWG");

  // Column Values for Rounds Sheets  
  var colRndPlyr =     cfgColRndSht[ 0][0];
  var colRndStatus =   cfgColRndSht[ 1][0];
  var colRndMP =       cfgColRndSht[ 2][0];
  var colRndWins =     cfgColRndSht[ 3][0];
  var colRndLoss =     cfgColRndSht[ 4][0];
  var colRndTie =      cfgColRndSht[ 5][0];
  var colRndPts =      cfgColRndSht[ 6][0];
  var colRndWinPerc =  cfgColRndSht[ 7][0];
  var colRndSports =   cfgColRndSht[ 8][0];
  var colRndLocation = cfgColRndSht[ 9][0];
  var colRndBalBonus = cfgColRndSht[10][0];
  var colRndPenLoss =  cfgColRndSht[11][0];
  var colRndMatchup =  cfgColRndSht[12][0];
  
  // Column Values for Data in Match Result Sheet
  var colArrRsltRound =    cfgColMatchRslt[ 3][0]-1;
  var colArrRsltPT1 =      cfgColMatchRslt[ 4][0]-1;
  var colArrRsltPT1Pts =   cfgColMatchRslt[ 5][0]-1;
  var colArrRsltPT2 =      cfgColMatchRslt[ 6][0]-1;
  var colArrRsltPT2Pts =   cfgColMatchRslt[ 7][0]-1;
  var colArrRsltTie =      cfgColMatchRslt[ 8][0]-1;
  var colArrRsltLoc =      cfgColMatchRslt[ 9][0]-1;  
  
  // League Parameters
  var evntGameType =       cfgEvntParam[ 4][0];
  var evntBalBonus =       cfgEvntParam[21][0];
  var evntBalBonusVal =    cfgEvntParam[22][0];
  var evntLocationBonus =  cfgEvntParam[23][0];  
  var evntPtsGainedMatch = cfgEvntParam[27][0];
  
  // function variables
  var shtRndRslt;
  var shtRndMaxCol;
  var RndRecPT1;
  var RndLocPT1
  var RndRecPT2;
  var RndLocPT2;
  var RndPackData;
  var RndPT1Matchup;
  var RndPT2Matchup;
  
  var RndPT1Row = 0;
  var RndPT2Row = 0;
  var RndMatchTie = 0; // Match is not a Tie by default
  
  var WinrPT;
  var LosrPT;
  var LosrRow;
  var BalBonusVal;
  
  // Match Values
  var MatchRound =      ResultData[0][colArrRsltRound];
  var MatchDataPT1 =    ResultData[0][colArrRsltPT1];
  var MatchDataPT1Pts = ResultData[0][colArrRsltPT1Pts];
  var MatchDataPT2 =    ResultData[0][colArrRsltPT2];
  var MatchDataPT2Pts = ResultData[0][colArrRsltPT2Pts];
  var MatchDataTie  =   ResultData[0][colArrRsltTie];
  var MatchLoc =        ResultData[0][colArrRsltLoc];
  
  var Round = 'Round'+MatchRound;
  shtRndRslt = ss.getSheetByName(Round);
  
  shtRndMaxCol = shtRndRslt.getMaxColumns();

  // Find Player Rows : subFindPlayerRow(sheet, rowStart, colPlyr, length, PlayerName)
  RndPT1Row = subFindPlayerRow(shtRndRslt, 5, colRndPlyr, 32, MatchDataPT1);
  RndPT2Row = subFindPlayerRow(shtRndRslt, 5, colRndPlyr, 32, MatchDataPT2);
  
  // Get Player/Team 1 and 2 Records when both rows have been found
  // Get Player/Team 1 and 2 Match Record, 6 values: Matches Played, Wins, Loss, Ties, Points, Win%
  RndRecPT1 = shtRndRslt.getRange(RndPT1Row,colRndMP,1,6).getValues();
  RndRecPT2 = shtRndRslt.getRange(RndPT2Row,colRndMP,1,6).getValues();
  
  // Update Player/Team 1 and 2 Location Bonus if Applicable
  if(evntLocationBonus == "Enabled" && (MatchLoc == 'Yes' || MatchLoc == 'Oui')){
    RndLocPT1 = shtRndRslt.getRange(RndPT1Row,colRndLocation).getValue() + 1;
    RndLocPT2 = shtRndRslt.getRange(RndPT2Row,colRndLocation).getValue() + 1;
    
    shtRndRslt.getRange(RndPT1Row,colRndLocation).setValue(RndLocPT1);
    shtRndRslt.getRange(RndPT2Row,colRndLocation).setValue(RndLocPT2);
  }
  
  // Match Tie Result
  if((evntPtsGainedMatch == "Enabled" && MatchDataPT1Pts == MatchDataPT2Pts) || (ResultData[0][colArrRsltTie] == 'Yes' || ResultData[0][colArrRsltTie] == 'Oui')){
    RndMatchTie = 1;  
    // Update Player 1
    RndRecPT1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT1, "Tie", MatchDataPT1Pts);
    // Update Player 2
    RndRecPT2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT2, "Tie", MatchDataPT2Pts);
  }
  // If Match is not Tie
  if(RndMatchTie == 0) {
    
    // If Option Points Gained in Match is Enabled
    if(evntPtsGainedMatch == "Disabled"){
      RndRecPT1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT1, "Win" , MatchDataPT1Pts);
      RndRecPT2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT2, "Loss", MatchDataPT2Pts);
      WinrPT = MatchDataPT1;
      LosrPT = MatchDataPT2;
    }
    
    // If Option Points Gained in Match is Enabled
    if(evntPtsGainedMatch == "Enabled"){
      // Player/Team 1 Points > Player/Team 2 Points
      if(MatchDataPT1Pts > MatchDataPT2Pts) {
        RndRecPT1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT1, "Win" , MatchDataPT1Pts);
        RndRecPT2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT2, "Loss", MatchDataPT2Pts);
        WinrPT = MatchDataPT1;
        LosrPT = MatchDataPT2;
      }
      // Player/Team 2 Points > Player/Team 1 Points
      if(MatchDataPT2Pts > MatchDataPT1Pts) {
        RndRecPT1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT1, "Loss", MatchDataPT1Pts);
        RndRecPT2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPT2, "Win" , MatchDataPT2Pts);
        WinrPT = MatchDataPT2;
        LosrPT = MatchDataPT1;
      }
    }
  }

  // Update Round Matchups
  // Player/Team 1
  RndPT1Matchup = shtRndRslt.getRange(RndPT1Row,colRndMatchup).getValue();
  if(RndPT1Matchup == '') RndPT1Matchup = MatchDataPT2;
  else RndPT1Matchup += ', ' + MatchDataPT2;
  
  // Player/Team 2
  RndPT2Matchup = shtRndRslt.getRange(RndPT2Row,colRndMatchup).getValue();
  if(RndPT2Matchup == '') RndPT2Matchup = MatchDataPT1;
  else RndPT2Matchup += ', ' + MatchDataPT1;
  
  // Update the Round Results Sheet
  shtRndRslt.getRange(RndPT1Row,colRndMP,1,6).setValues(RndRecPT1);
  shtRndRslt.getRange(RndPT2Row,colRndMP,1,6).setValues(RndRecPT2);
  shtRndRslt.getRange(RndPT1Row,colRndMatchup).setValue(RndPT1Matchup);
  shtRndRslt.getRange(RndPT2Row,colRndMatchup).setValue(RndPT2Matchup);

  
  // If Match is not a Tie and Balance Bonus is Enabled
  if (RndMatchTie == 0 && evntBalBonus == 'Enabled'){
    // Get Row for Losing Player/Team
    if(LosrPT == MatchDataPT1) LosrRow = RndPT1Row;
    if(LosrPT == MatchDataPT2) LosrRow = RndPT2Row;
    // Get Loser Amount of Balance Bonus Points and Increase by value from Config file
    BalBonusVal = shtRndRslt.getRange(LosrRow,colRndBalBonus).getValue() + evntBalBonusVal;
    shtRndRslt.getRange(LosrRow,colRndBalBonus).setValue(BalBonusVal);
  }
  
  // Update Match Data
  // Player/Team 1
  MatchData[3][1]= MatchDataPT1Pts;  // Points
  MatchData[3][2]= RndRecPT1[0][0];  // Matches Played this Event
  
  // Player/Team 2
  MatchData[4][1]= MatchDataPT2Pts;  // Points
  MatchData[4][2]= RndRecPT2[0][0];  // Matches Played this Event
  MatchData[4][3]= BalBonusVal;
  
  
  return MatchData;
}


// **********************************************
// function fcnFindDuplicateData()
//
// This function searches the entry list to find any 
// duplicate responses. To make sure we do not interfere 
// with the fcnFindMatchingData, we look for a non-zero Match ID
//
// The functions returns the Row number where the matching data was found. 
// 
// If no duplicate data is found, it returns 0;
//
// **********************************************

function fcnFindDuplicateData(ss, shtRspn, cfgEvntParam, cfgColRspSht, cfgColMatchRep, RspnDataInputs, ResponseData, RspnRow, RspnMaxRows) {
  
  Logger.log("Routine: fcnFindDuplicateData");
    
  // Event Parameters
  var evntFormat =   cfgEvntParam[9][0];

  // Column Values for Data in Response Sheet
  var colArrRspnMatchID = cfgColRspSht[1][0]-1;
  var colArrRspnPrcsd =   cfgColRspSht[2][0]-1;
  
  // Column Values for Data in Response Sheet
  var colArrRspnPwd =     cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =   cfgColMatchRep[ 2][0]-1;
  var colArrRspnPlyr1 =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnTeam1 =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnPts1 =    cfgColMatchRep[ 5][0]-1;
  var colArrRspnPlyr2 =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnTeam2 =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnPts2 =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =     cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =     cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub = cfgColMatchRep[19][0]-1;
   
  // Values from Response Data
  var RspnDataRound =     ResponseData[0][colArrRspnRound];   // Round Number
  
  // Winning and Losing Player/Team
  // If Single Event
  if(evntFormat == "Single"){
    var RspnDataPT1 = ResponseData[0][colArrRspnPlyr1];  // Winning Player
    var RspnDataPT2 = ResponseData[0][colArrRspnPlyr2];  // Losing Player
  }
  // If Team Event
  if(evntFormat == "Team"){
    var RspnDataPT1 = ResponseData[0][colArrRspnTeam1];  // Winning Team
    var RspnDataPT2 = ResponseData[0][colArrRspnTeam2];  // Losing Team
  }
  
  // Entry Data
  var EntryRound;
  var EntryPT1;
  var EntryPT2;
  var EntryData;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DuplicateRow = 0;
  
  // Gets Entry Data to analyze
  EntryData = shtRspn.getRange(1, 1, RspnMaxRows, RspnDataInputs).getValues();
  
  // Loop to find if another entry has the same data
  for (var EntryRow = 1; EntryRow < RspnMaxRows; EntryRow++){
    
    // Filters only entries of the same Round the response was posted
    if (EntryData[EntryRow][colArrRspnRound] == RspnDataRound){
      
      EntryRound = EntryData[EntryRow][colArrRspnRound]; // Round

      // Winning and Losing Player/Team
      // If Single Event
      if(evntFormat == "Single"){
        EntryPT1 = EntryData[EntryRow][colArrRspnPlyr1];  // Winning Player
        EntryPT2 = EntryData[EntryRow][colArrRspnPlyr2];  // Losing Player
      }
      // If Team Event
      if(evntFormat == "Team"){
        EntryPT1 = EntryData[EntryRow][colArrRspnTeam1];  // Winning Team
        EntryPT2 = EntryData[EntryRow][colArrRspnTeam2];  // Losing Team
      }
      
      EntryMatchID = EntryData[EntryRow][colArrRspnMatchID]; // Match ID
      EntryPrcssd =  EntryData[EntryRow][colArrRspnPrcsd];   // Entry Processed Flag
            
      // If both rows are different, the Data Entry was processed and was compiled in the Match Results (Match as a Match ID), Look for player entry combination
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID != ''){
        // If combination of players are the same between the entry data and the new response data, duplicate entry was found. Save Row index
        if ((RspnDataPT1 == EntryPT1 && RspnDataPT2 == EntryPT2) || (RspnDataPT1 == EntryPT1 && RspnDataPT2 == EntryPT2)){
          DuplicateRow = EntryRow + 1;
          EntryRow = RspnMaxRows + 1;
        }
      }
    }
            
    // If we do not detect any value in Round Column, we reached the end of the list and skip
    if (EntryRow <= RspnMaxRows && EntryData[EntryRow][colArrRspnRound] == ''){
      EntryRow = RspnMaxRows + 1;
    }
  }
  return DuplicateRow;
}


// **********************************************
// function fcnFindMatchingData()
//
// This function searches for the other match entry 
// in the Response Tab. The functions returns the Row number
// where the matching data was found. 
// 
// If no matching data is found, it returns 0;
//
// **********************************************

function fcnFindMatchingData(ss, cfgEvntParam, cfgColRspSht, cfgColMatchRep, cfgExecData, shtRspn, ResponseData, RspnRow, RspnMaxRows) {
  
  Logger.log("Routine: fcnFindMatchingData");

  // Code Execution Options
  var exeDualSubmission = cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead

  // Event Parameters
  var evntFormat =   cfgEvntParam[9][0];
  
  // Column Values for Data in Response Sheet
  var RspnDataInputs =    cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var colDataConflict =   cfgColRspSht[3][0];
  
  var colArrRspnMatchID = cfgColRspSht[1][0]-1;
  var colArrRspnPrcsd =   cfgColRspSht[2][0]-1;
  
  // Column Values for Data in Response Sheet
  var colArrRspnPwd =     cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =   cfgColMatchRep[ 2][0]-1;
  var colArrRspnPlyr1 =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnTeam1 =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnPts1 =    cfgColMatchRep[ 5][0]-1;
  var colArrRspnPlyr2 =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnTeam2 =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnPts2 =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =     cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =     cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub = cfgColMatchRep[19][0]-1;
  
  // Response Data
  var RspnDataPlyrSubmit = ResponseData[0][colArrRspnPlyrSub]; // Player Submitting
  var RspnDataRound =      ResponseData[0][colArrRspnRound];

  // Winning and Losing Player/Team
  // If Single Event
  if(evntFormat == "Single"){
    var RspnDataPT1 = ResponseData[0][colArrRspnPlyr1];  // Winning Player
    var RspnDataPT2 = ResponseData[0][colArrRspnPlyr2];  // Losing Player
  }
  // If Team Event
  if(evntFormat == "Team"){
    var RspnDataPT1 = ResponseData[0][colArrRspnTeam1];  // Winning Team
    var RspnDataPT2 = ResponseData[0][colArrRspnTeam2];  // Losing Team
  }

  var EntryData;
  var EntryPlyrSubmit;
  var EntryRound;
  var EntryPT1;
  var EntryPT2;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DataMatchingRow = 0;
  
  var DataConflict = -1;
  
  // Gets Entry Data to analyze
  EntryData = shtRspn.getRange(1, 1, RspnMaxRows, RspnDataInputs).getValues();
  
  // Loop to find if another entry has the same data
  for (var EntryRow = 1; EntryRow <= RspnMaxRows; EntryRow++){
    
    // Filters only entries of the same Round the response was posted
    if (EntryData[EntryRow][colArrRspnRound] == RspnDataRound){
      
      EntryRound = EntryData[EntryRow][colArrRspnRound]; // Round
      
      // Winning and Losing Player/Team
      // If Single Event
      if(evntFormat == "Single"){
        EntryPT1 = EntryData[EntryRow][colArrRspnPlyr1];  // Winning Player
        EntryPT2 = EntryData[EntryRow][colArrRspnPlyr2];  // Losing Player
      }
      // If Team Event
      if(evntFormat == "Team"){
        EntryPT1 = EntryData[EntryRow][colArrRspnTeam1];  // Winning Team
        EntryPT2 = EntryData[EntryRow][colArrRspnTeam2];  // Losing Team
      }
      
      EntryMatchID = EntryData[EntryRow][colArrRspnMatchID]; // Match ID
      EntryPrcssd =  EntryData[EntryRow][colArrRspnPrcsd];   // Entry Processed Flag
      
      // If both rows are different, Round Number, Player A and Player B are matching, we found the other match to compare data to
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID == '' && RspnDataRound == EntryRound && RspnDataPT1 == EntryPT1 && RspnDataPT2 == EntryPT2){
        
        // If Dual Submission is Enabled, look for Player Submitting, if they are different, continue          
        if ((exeDualSubmission == 'Enabled' && RspnDataPlyrSubmit != EntryPlyrSubmit) || exeDualSubmission == 'Disabled'){ 
          
          // Compare New Response Data and Entry Data. If Data is not equal to the other, the conflicting Data ID is returned
          DataConflict = subCheckDataConflict(ResponseData, EntryData, colArrRspnRound, colArrRspnMatchID-1);
          
          // 
          if (DataConflict == 0){
            // Sets Conflict Flag to 'No Conflict'
            shtRspn.getRange(RspnRow, colDataConflict).setValue('No Conflict');
            shtRspn.getRange(EntryRow, colDataConflict).setValue('No Conflict');
            DataMatchingRow = EntryRow;
          }
          
          // If Data Conflict was detected, sends email to notify Data Conflict
          if (DataConflict != 0 && DataConflict != -1){
            
            // Sets the Conflict Value to the Data ID value where the conflict was found
            shtRspn.getRange(RspnRow, colDataConflict).setValue(DataConflict);
            shtRspn.getRange(EntryRow, colDataConflict).setValue(DataConflict);
          }
        }
      }
      
      // If Dual Submission is Enabled, look for Player Submitting, if they are the same, set negative value of Entry Row as Duplicate          
      if (exeDualSubmission == 'Enabled' && RspnDataPlyrSubmit == EntryPlyrSubmit){
        DataMatchingRow = 0 - EntryRow;
      }
      
      // Loop reached the end of responses entered or found matching data
      if(EntryRound == '' || DataMatchingRow != 0) {
        EntryRow = RspnMaxRows + 1;
      }
    }
  }
  return DataMatchingRow;
}

