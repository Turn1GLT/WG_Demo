// **********************************************
// function fcnPostMatchResultsWG()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Round X tab
//
// **********************************************

function fcnPostMatchResultsWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData, cfgColMatchRep, ResponseData, MatchingRspnData, MatchID, MatchData) {
  
  Logger.log("Routine: fcnPostMatchResultsWG");
  
  var cfgColMatchRslt = ss.getSheetByName("Config").getRange(21,18,32,1).getValues();
  
  // Code Execution Options
  var exeDualSubmission =      cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var exePlyrMatchValidation = cfgExecData[2][0];
  var exePostRoundResult =     cfgExecData[8][0];
  
  // Event Parameters
  var evntGameType =       cfgEvntParam[4][0];
  var evntFormat =         cfgEvntParam[9][0];
  var evntLocationBonus =  cfgEvntParam[23][0];
  var evntPtsGainedMatch = cfgEvntParam[32][0];
  var evntTiePossible =    cfgEvntParam[34][0];
  
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
  var MatchValidWinr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var MatchValidLosr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  
  // Column Values for Data in Response Sheet
  var colArrRspnPwd =       cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =     cfgColMatchRep[ 2][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinPts =    cfgColMatchRep[ 4][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosPts =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =       cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =       cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub =   cfgColMatchRep[19][0]-1;

  // Column Values for Data in Match Result Sheet
  var colArrRsltResultID = cfgColMatchRslt[ 0][0]-1;
  var colArrRsltMatchCnt = cfgColMatchRslt[ 1][0]-1;
  var colArrRsltMatchID =  cfgColMatchRslt[ 2][0]-1;
  var colArrRsltRound =    cfgColMatchRslt[ 3][0]-1;
  var colArrRsltWinPT =    cfgColMatchRslt[ 4][0]-1;
  var colArrRsltWinPts =   cfgColMatchRslt[ 5][0]-1;
  var colArrRsltLosPT =    cfgColMatchRslt[ 6][0]-1;
  var colArrRsltLosPts =   cfgColMatchRslt[ 7][0]-1;
  var colArrRsltTie =      cfgColMatchRslt[ 8][0]-1;
  var colArrRsltLoc =      cfgColMatchRslt[ 9][0]-1;
  var colArrRsltBal =      cfgColMatchRslt[10][0]-1;  
  
  // Column Values for Rounds Sheets
  var colRndBalBonus = cfgColRndSht[ 9][0];
  
  var MatchPostedStatus = 0;
  var Winr;
  var Losr;
  
  // Sets which Data set is Player A and Player B. Data[0][1] = Player who posted the data
  if (exeDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][2]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchingRspnData;
  }
  
  if (exeDualSubmission == 'Enabled' && ResponseData[0][1] == ResponseData[0][3]) {
    RsltPlyrDataA = ResponseData;
    RsltPlyrDataB = MatchingRspnData;
  }
  
  // Copy Players Data
  ResultData[0][colArrRsltRound] = ResponseData[0][colArrRspnRound];  // Round Number
  
  // If Single Event
  if(evntFormat == "Single"){
    ResultData[0][colArrRsltWinPT] = ResponseData[0][colArrRspnWinPlyr];  // Winning Player
    ResultData[0][colArrRsltLosPT] = ResponseData[0][colArrRspnLosPlyr];  // Losing Player
  }
  // If Team Event
  if(evntFormat == "Team"){
    ResultData[0][colArrRsltWinPT] = ResponseData[0][colArrRspnWinTeam];  // Winning Team
    ResultData[0][colArrRsltLosPT] = ResponseData[0][colArrRspnLosTeam];  // Losing Team
  }
  
  // Points
  if(evntPtsGainedMatch == "Enabled"){
    ResultData[0][colArrRsltWinPts] = ResponseData[0][colArrRspnWinPts]; // Winning Player Points
    ResultData[0][colArrRsltLosPts] = ResponseData[0][colArrRspnLosPts]; // Losing Player Points
  }
  else{
    ResultData[0][colArrRsltWinPts] = "-"; // Winning Player Points
    ResultData[0][colArrRsltLosPts] = "-"; // Losing Player Points
  }
  
  // Winner and Loser Names
  Winr = ResultData[0][colArrRsltWinPT];
  Losr = ResultData[0][colArrRsltLosPT];
  
  ResultData[0][colArrRsltTie] =    ResponseData[0][colArrRspnTie];  // Game is Tie
  if(evntLocationBonus == "Enabled") ResultData[0][colArrRsltLoc]   = ResponseData[0][colArrRspnLoc];  // Location
  
  // If option is enabled, Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  if (exePlyrMatchValidation == 'Enabled'){
    // Call subroutine to check if players match are valid
    MatchValidWinr = subPlayerMatchValidation(ss, Winr, MatchValidWinr);
    Logger.log('%s Match Validation: %s',Winr, MatchValidWinr[0]);
    
    MatchValidLosr = subPlayerMatchValidation(ss, Losr, MatchValidLosr);
    Logger.log('%s Match Validation: %s',Losr, MatchValidLosr[0]);
  }

  // If option is disabled, Consider Matches are valid
  if (exePlyrMatchValidation == 'Disabled'){
    MatchValidWinr[0] = 1;
    MatchValidLosr[0] = 1;
  }
  
  // If both players have played a valid match
  if (MatchValidWinr[0] == 1 && MatchValidLosr[0] == 1){
    
    // Copies Match ID
    ResultData[0][colArrRsltMatchID] = MatchID; // Match ID
    
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
    // Reserved for TCG
        
    
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
  if (MatchValidWinr[0] == -1 && MatchValidLosr[0] == 1)  MatchPostedStatus = -11;
  
  // returns Error that Winning Player has played too many matches
  if (MatchValidWinr[0] == -2 && MatchValidLosr[0] == 1)  MatchPostedStatus = -12;  
  
  // returns Error that Losing Player is Eliminated from the League
  if (MatchValidLosr[0] == -1 && MatchValidWinr[0] == 1)  MatchPostedStatus = -21;
  
  // returns Error that Losing Player has played too many matches
  if (MatchValidLosr[0] == -2 && MatchValidWinr[0] == 1)  MatchPostedStatus = -22;
  
  // returns Error that Both Players are Eliminated from the League
  if (MatchValidWinr[0] == -1 && MatchValidLosr[0] == -1) MatchPostedStatus = -31;
  
  // returns Error that Winning Player is Eliminated from the League and Losing Player has played too many matches
  if (MatchValidWinr[0] == -1 && MatchValidLosr[0] == -2) MatchPostedStatus = -32;

  // returns Error that Winning Player has player too many matches and Losing Player is Eliminated from the League
  if (MatchValidWinr[0] == -2 && MatchValidLosr[0] == -1) MatchPostedStatus = -33;
  
  // returns Error that Both Players have played too many matches
  if (MatchValidWinr[0] == -2 && MatchValidLosr[0] == -2) MatchPostedStatus = -34;
  
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
  var colRndTeam =     cfgColRndSht[ 1][0];
  var colRndMP =       cfgColRndSht[ 2][0];
  var colRndLocation = cfgColRndSht[ 8][0];
  var colRndBalBonus = cfgColRndSht[ 9][0];
  var colRndMatchup =  cfgColRndSht[11][0];
  
  // Column Values for Data in Match Result Sheet
  var colArrRsltRound =    cfgColMatchRslt[ 3][0]-1;
  var colArrRsltWinPT =    cfgColMatchRslt[ 4][0]-1;
  var colArrRsltWinPts =   cfgColMatchRslt[ 5][0]-1;
  var colArrRsltLosPT =    cfgColMatchRslt[ 6][0]-1;
  var colArrRsltLosPts =   cfgColMatchRslt[ 7][0]-1;
  var colArrRsltTie =      cfgColMatchRslt[ 8][0]-1;
  var colArrRsltLoc =      cfgColMatchRslt[ 9][0]-1;  
  
  // League Parameters
  var evntGameType =       cfgEvntParam[ 4][0];
  var evntBalBonus =       cfgEvntParam[21][0];
  var evntBalBonusVal =    cfgEvntParam[22][0];
  var evntLocationBonus =  cfgEvntParam[23][0];  
  
  // function variables
  var shtRndRslt;
  var shtRndMaxCol;
  var RndRecPlyr1;
  var RndLocPlyr1
  var RndRecPlyr2;
  var RndLocPlyr2;
  var RndPackData;
  var RndPlyr1Matchup;
  var RndPlyr2Matchup;
  
  var RndPlyr1Row = 0;
  var RndPlyr2Row = 0;
  var RndMatchTie = 0; // Match is not a Tie by default
  
  var Plyr2BalBonusVal;
  
  // Match Values
  var MatchRound =      ResultData[0][colArrRsltRound];
  var MatchDataPT1 =    ResultData[0][colArrRsltWinPT];
  var MatchDataPT1Pts = ResultData[0][colArrRsltWinPts];
  var MatchDataPT2 =    ResultData[0][colArrRsltLosPT];
  var MatchDataPT2Pts = ResultData[0][colArrRsltLosPts];
  var MatchDataTie  =   ResultData[0][colArrRsltTie];
  var MatchLoc =        ResultData[0][colArrRsltLoc];
  
  var Round = 'Round'+MatchRound;
  shtRndRslt = ss.getSheetByName(Round);
  
  shtRndMaxCol = shtRndRslt.getMaxColumns();

  // Find Player Rows : subFindPlayerRow(sheet, rowStart, colPlyr, length, PlayerName)
  RndPlyr1Row = subFindPlayerRow(shtRndRslt, 5, colRndPlyr, 32, MatchDataPT1);
  RndPlyr2Row = subFindPlayerRow(shtRndRslt, 5, colRndPlyr, 32, MatchDataPT2);
  
  // Get Winner and Loser Records when both rows have been found
  // Get Winner (Player 1) and Loser (Player 2) Match Record, 6 values: Matches Played, Wins, Loss, Ties, Points, Win%
  RndRecPlyr1 = shtRndRslt.getRange(RndPlyr1Row,colRndMP,1,6).getValues();
  RndRecPlyr2 = shtRndRslt.getRange(RndPlyr2Row,colRndMP,1,6).getValues();
  
  // Update Winner and Loser Location Bonus if Applicable
  if(evntLocationBonus == "Enabled" && (MatchLoc == 'Yes' || MatchLoc == 'Oui')){
    RndLocPlyr1 = shtRndRslt.getRange(RndPlyr1Row,colRndLocation).getValue() + 1;
    RndLocPlyr2 = shtRndRslt.getRange(RndPlyr2Row,colRndLocation).getValue() + 1;
    
    shtRndRslt.getRange(RndPlyr1Row,colRndLocation).setValue(RndLocPlyr1);
    shtRndRslt.getRange(RndPlyr2Row,colRndLocation).setValue(RndLocPlyr2);
  }
  
  // Match Tie Result
  if((MatchDataPT1Pts == MatchDataPT2Pts) || (ResultData[0][colArrRsltTie] == 'Yes' || ResultData[0][colArrRsltTie] == 'Oui')){
    RndMatchTie = 1;  
    // Update Player 1
    RndRecPlyr1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPlyr1, "Tie", MatchDataPT1Pts);
    // Update Player 2
    RndRecPlyr2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPlyr2, "Tie", MatchDataPT2Pts);
  }
  if(RndMatchTie == 0) {
    // Update Player 1
    RndRecPlyr1 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPlyr1, "Win", MatchDataPT1Pts);
    // Update Player 2
    RndRecPlyr2 = subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPlyr2, "Loss", MatchDataPT2Pts);
  }

  // Update Round Matchups
  // Winning Player
  RndPlyr1Matchup = shtRndRslt.getRange(RndPlyr1Row,colRndMatchup).getValue();
  if(RndPlyr1Matchup == '') RndPlyr1Matchup = MatchDataPT2;
  else RndPlyr1Matchup += ', ' + MatchDataPT2;
  
  // Losing Player
  RndPlyr2Matchup = shtRndRslt.getRange(RndPlyr2Row,colRndMatchup).getValue();
  if(RndPlyr2Matchup == '') RndPlyr2Matchup = MatchDataPT1;
  else RndPlyr2Matchup += ', ' + MatchDataPT1;
  
  // Update the Round Results Sheet
  shtRndRslt.getRange(RndPlyr1Row,colRndMP,1,6).setValues(RndRecPlyr1);
  shtRndRslt.getRange(RndPlyr2Row,colRndMP,1,6).setValues(RndRecPlyr2);
  shtRndRslt.getRange(RndPlyr1Row,colRndMatchup).setValue(RndPlyr1Matchup);
  shtRndRslt.getRange(RndPlyr2Row,colRndMatchup).setValue(RndPlyr2Matchup);

  // If Match is not a Tie and Balance Bonus is Enabled
  if (RndMatchTie == 0 && evntBalBonus == 'Enabled'){
    // Get Loser Amount of Balance Bonus Points and Increase by value from Config file
    Plyr2BalBonusVal = shtRndRslt.getRange(RndPlyr2Row,colRndBalBonus).getValue() + evntBalBonusVal;
    shtRndRslt.getRange(RndPlyr2Row,colRndBalBonus).setValue(Plyr2BalBonusVal);
  }
  
  // Update Match Data
  // Player/Team 1
  MatchData[3][1]= MatchDataPT1Pts;  // Points
  MatchData[3][2]= RndRecPlyr1[0][0];  // Matches Played this Event
  
  // Player/Team 2
  MatchData[4][1]= MatchDataPT2Pts;  // Points
  MatchData[4][2]= RndRecPlyr2[0][0];  // Matches Played this Event
  MatchData[4][3]= Plyr2BalBonusVal;
  
  
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
  var colArrRspnPwd =       cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =     cfgColMatchRep[ 2][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinPts =    cfgColMatchRep[ 4][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosPts =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =       cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =       cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub =   cfgColMatchRep[19][0]-1;
   
  // Values from Response Data
  var RspnDataRound =     ResponseData[0][colArrRspnRound];   // Round Number
  
  // Winning and Losing Player/Team
  // If Single Event
  if(evntFormat == "Single"){
    var RspnDataWinPlyr = ResponseData[0][colArrRspnWinPlyr];  // Winning Player
    var RspnDataLosPlyr = ResponseData[0][colArrRspnLosPlyr];  // Losing Player
  }
  // If Team Event
  if(evntFormat == "Team"){
    var RspnDataWinPlyr = ResponseData[0][colArrRspnWinTeam];  // Winning Team
    var RspnDataLosPlyr = ResponseData[0][colArrRspnLosTeam];  // Losing Team
  }
  
  // Entry Data
  var EntryRound;
  var EntryWinr;
  var EntryLosr;
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
        EntryWinr = EntryData[EntryRow][colArrRspnWinPlyr];  // Winning Player
        EntryLosr = EntryData[EntryRow][colArrRspnLosPlyr];  // Losing Player
      }
      // If Team Event
      if(evntFormat == "Team"){
        EntryWinr = EntryData[EntryRow][colArrRspnWinTeam];  // Winning Team
        EntryLosr = EntryData[EntryRow][colArrRspnLosTeam];  // Losing Team
      }
      
      EntryMatchID = EntryData[EntryRow][colArrRspnMatchID]; // Match ID
      EntryPrcssd =  EntryData[EntryRow][colArrRspnPrcsd];   // Entry Processed Flag
            
      // If both rows are different, the Data Entry was processed and was compiled in the Match Results (Match as a Match ID), Look for player entry combination
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID != ''){
        // If combination of players are the same between the entry data and the new response data, duplicate entry was found. Save Row index
        if ((RspnDataWinPlyr == EntryWinr && RspnDataLosPlyr == EntryLosr) || (RspnDataWinPlyr == EntryLosr && RspnDataLosPlyr == EntryWinr)){
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
  var colArrRspnPwd =       cfgColMatchRep[ 1][0]-1;
  var colArrRspnRound =     cfgColMatchRep[ 2][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinPts =    cfgColMatchRep[ 4][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 6][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosPts =    cfgColMatchRep[ 8][0]-1;
  var colArrRspnTie =       cfgColMatchRep[ 9][0]-1;
  var colArrRspnLoc =       cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub =   cfgColMatchRep[19][0]-1;
  
  // Response Data
  var RspnDataPlyrSubmit = ResponseData[0][colArrRspnPlyrSub]; // Player Submitting
  var RspnDataRound =      ResponseData[0][colArrRspnRound];

  // Winning and Losing Player/Team
  // If Single Event
  if(evntFormat == "Single"){
    var RspnDataWinr = ResponseData[0][colArrRspnWinPlyr];  // Winning Player
    var RspnDataLosr = ResponseData[0][colArrRspnLosPlyr];  // Losing Player
  }
  // If Team Event
  if(evntFormat == "Team"){
    var RspnDataWinr = ResponseData[0][colArrRspnWinTeam];  // Winning Team
    var RspnDataLosr = ResponseData[0][colArrRspnLosTeam];  // Losing Team
  }

  var EntryData;
  var EntryPlyrSubmit;
  var EntryRound;
  var EntryWinr;
  var EntryLosr;
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
        EntryWinr = EntryData[EntryRow][colArrRspnWinPlyr];  // Winning Player
        EntryLosr = EntryData[EntryRow][colArrRspnLosPlyr];  // Losing Player
      }
      // If Team Event
      if(evntFormat == "Team"){
        EntryWinr = EntryData[EntryRow][colArrRspnWinTeam];  // Winning Team
        EntryLosr = EntryData[EntryRow][colArrRspnLosTeam];  // Losing Team
      }
      
      EntryMatchID = EntryData[EntryRow][colArrRspnMatchID]; // Match ID
      EntryPrcssd =  EntryData[EntryRow][colArrRspnPrcsd];   // Entry Processed Flag
      
      // If both rows are different, Round Number, Player A and Player B are matching, we found the other match to compare data to
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID == '' && RspnDataRound == EntryRound && RspnDataWinr == EntryWinr && RspnDataLosr == EntryLosr){
        
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

