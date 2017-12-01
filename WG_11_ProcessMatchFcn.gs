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
  
  var cfgColMatchRslt = ss.getSheetByName("Config").getRange(20,18,32,1).getValues();
  
  // Code Execution Options
  var exeDualSubmission = cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var exePlyrMatchValidation = cfgExecData[2][0];
  
  // Event Parameters
  var evntGameType = cfgEvntParam[4][0];
  var evntFormat =   cfgEvntParam[9][0];
  
  // Cumulative Results sheet variables
  var shtCumul;
  var BalanceBonusLosr;
  
  // Match Results Sheet Values
  var shtRslt = ss.getSheetByName('Match Results');
  var shtRsltMaxRows = shtRslt.getMaxRows();
  var shtRsltMaxCol = shtRslt.getMaxColumns();
  var RsltLastResultRowRng = shtRslt.getRange(3, 4);
  var RsltLastResultRow = RsltLastResultRowRng.getValue() + 1;
  var RsltRng = shtRslt.getRange(RsltLastResultRow, 1, 1, shtRsltMaxCol);
  var ResultData = RsltRng.getValues();
  
  // Routine Variables
  var MatchValidWinr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var MatchValidLosr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  var DataPostedLosr;
  
  // Column Values for Data in Response Sheet
  var colArrRspnPwd =       cfgColMatchRep[ 1][0]-1;
  var colArrRspnLoc =       cfgColMatchRep[ 2][0]-1;
  var colArrRspnRound =     cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 5][0]-1;
  var colArrRspnWinPts =    cfgColMatchRep[ 6][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 8][0]-1;
  var colArrRspnLosPts =    cfgColMatchRep[ 9][0]-1;
  var colArrRspnTie =       cfgColMatchRep[10][0]-1;
  var colArrRspnPlyrSub =   cfgColMatchRep[19][0]-1;

  // Column Values for Data in Match Result Sheet
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
  
  // Copies Players Data
  ResultData[0][colArrRsltLoc]   = ResponseData[0][colArrRspnLoc];  // Location
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
  // Winner and Losser Names
  Winr = ResultData[0][colArrRsltWinPT];
  Losr = ResultData[0][colArrRsltLosPT];
  
  ResultData[0][colArrRsltTie] =    ResponseData[0][colArrRspnTie];  // Game is Tie
  ResultData[0][colArrRsltWinPts] = ResponseData[0][colArrRspnWinPts]; // Winning Player Points
  ResultData[0][colArrRsltLosPts] = ResponseData[0][colArrRspnLosPts]; // Losing Player Points
  

  
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
    
//    // Post Results in Appropriate Round Number for Both Players
//    // DataPostedLosr is an Array with [0]=Post Status (1=Success) [1]=Loser Row [2]=Power Level Column
//    DataPostedLosr = fcnPostResultRoundWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRslt, ResultData);
//    
//    // Gets New Power Level / Points Bonus for Loser from Cumulative Results Sheet
//    shtCumul = ss.getSheetByName('Cumulative Results');
//    BalanceBonusLosr = shtCumul.getRange(DataPostedLosr[1],DataPostedLosr[2]).getValue();
//    Logger.log('Cumulative Power Level: %s',BalanceBonusLosr);
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
  // Time Stamp
  MatchData[0][0] = Utilities.formatDate (ResponseData[0][0], Session.getScriptTimeZone(), 'YYYY-MM-dd HH:mm:ss');
  
  MatchData[1][0] = ResultData[0][colArrRsltMatchID]; // MatchID
  MatchData[2][0] = ResultData[0][colArrRsltRound];   // Round Number
  
  // Winning Side
  MatchData[3][0] = ResultData[0][colArrRsltWinPT];  // Winning Player/Team
  MatchData[3][1] = ResultData[0][colArrRsltWinPts]; // Winning Points
  MatchData[3][2] = MatchValidWinr[1];               // Winning Player/Team Matches Played
  
  // Losing Side
  MatchData[4][0] = ResultData[0][colArrRsltLosPT];  // Losing Player/Team
  MatchData[4][1] = ResultData[0][colArrRsltLosPts]; // Losing Points
  MatchData[4][2] = MatchValidLosr[1];               // Losing Player/Team Matches Played
  MatchData[4][3] = BalanceBonusLosr;                // Losing Player/Team Balance Bonus Value

  MatchData[5][0] = ResultData[0][colArrRsltTie];    // Game is Tie
  MatchData[6][0] = ResultData[0][colArrRsltLoc];    // Location (Store Y/N)
  MatchData[7][0] = "";
  MatchData[25][0] = MatchPostedStatus;
  
  return MatchData;
}


// **********************************************
// function fcnPostResultRoundWG()
//
// Once the Match Data has been posted in the 
// Match Results Tab, the Round X results are updated
// for each player
//
// **********************************************

function fcnPostResultRoundWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRslt, ResultData) {
  
  Logger.log("Routine: fcnPostResultRoundWG");

  // Column Values for Rounds Sheets
  var colRndPlyr =     cfgColRndSht[ 0][0];
  var colRndTeam =     cfgColRndSht[ 1][0];
  var colRndWin =      cfgColRndSht[ 3][0];
  var colRndLos =      cfgColRndSht[ 4][0];
  var colRndTie =      cfgColRndSht[ 5][0];
  var colRndPoints =   cfgColRndSht[ 6][0];
  var colRndWinPerc =  cfgColRndSht[ 7][0];
  var colRndLocation = cfgColRndSht[ 8][0];
  var colRndBalBonus = cfgColRndSht[ 9][0];
  var colRndPenLoss =  cfgColRndSht[10][0];
  var colRndMatchup =  cfgColRndSht[11][0];
  
  // Column Values for Data in Match Result Sheet
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
  
  // League Parameters
  var evntGameType =       cfgEvntParam[ 4][0];
  var evntBalBonus =       cfgEvntParam[21][0];
  var evntBalBonusVal =    cfgEvntParam[22][0];
  var evntLocationBonus =  cfgEvntParam[23][0];
  var evntPtsPerWin =      cfgEvntParam[29][0];
  var evntPtsPerLoss =     cfgEvntParam[30][0];
  var evntPtsPerTie =      cfgEvntParam[31][0];
  var evntPtsGainedMatch = cfgEvntParam[32][0];
  
  
  // function variables
  var shtRoundRslt;
  var shtRoundMaxCol;
  var RoundPlyrList;
  var RoundWinrRec;
  var RoundWinrLoc
  var RoundLosrRec;
  var RoundLosrLoc;
  var RoundPackData;
  var RoundWinrMatchup;
  var RoundLosrMatchup;
  
  var RoundWinrRow = 0;
  var RoundLosrRow = 0;
  var RoundMatchTie = 0; // Match is not a Tie by default
  
  var BalanceDataLosr = new Array(3);
  var LosrBalBonusVal;
  
  // Match Values
  var MatchRound =      ResultData[0][colArrRsltRound];
  var MatchDataWinPT =  ResultData[0][colArrRsltWinPT];
  var MatchDataWinPts = ResultData[0][colArrRsltWinPts];
  var MatchDataLosPT =  ResultData[0][colArrRsltLosPT];
  var MatchDataLosPts = ResultData[0][colArrRsltLosPts];
  var MatchDataTie  =   ResultData[0][colArrRsltTie];
  var MatchLoc =        ResultData[0][colArrRsltLoc];
  
  var Round = 'Round'+MatchRound;
  shtRoundRslt = ss.getSheetByName(Round);
  
  shtRoundMaxCol = shtRoundRslt.getMaxColumns();

  // Gets All Players Names
  RoundPlyrList = shtRoundRslt.getRange(5,colRndPlyr,32,1).getValues();
  
  // Find the Winning and Losing Player in the Round Result Tab
  for (var RsltRow = 5; RsltRow <= 36; RsltRow ++){
    
    // Get Rows for Winner and Loser
    if (RoundPlyrList[RsltRow - 5][0] == MatchDataWinPT) RoundWinrRow = RsltRow;
    if (RoundPlyrList[RsltRow - 5][0] == MatchDataLosPT) RoundLosrRow = RsltRow;
    
    // Get Winner and Loser Records when both rows have been found
    if (RoundWinrRow != '' && RoundLosrRow != '') {
      // Get Winner and Loser Match Record, 4 values: Win, Loss, Ties, Points
      RoundWinrRec = shtRoundRslt.getRange(RoundWinrRow,colRndWin,1,4).getValues();
      RoundLosrRec = shtRoundRslt.getRange(RoundLosrRow,colRndWin,1,4).getValues();
      
      // Get Winner and Loser Location Bonus if Applicable
      if(evntLocationBonus == "Enabled"){
        RoundWinrLoc = shtRoundRslt.getRange(RoundWinrRow,colRndLocation).getValue();
        RoundLosrLoc = shtRoundRslt.getRange(RoundLosrRow,colRndLocation).getValue();
      }
      // Exit For Loop
      RsltRow = 37;
    }
  }
  
  // Initializes Empty Cells for both Winner and Loser
  if (RoundWinrRec[0][0] == '') RoundWinrRec[0][0] = 0; // Win
  if (RoundWinrRec[0][1] == '') RoundWinrRec[0][1] = 0; // Loss
  if (RoundWinrRec[0][2] == '') RoundWinrRec[0][2] = 0; // Tie
  if (RoundWinrRec[0][3] == '') RoundWinrRec[0][3] = 0; // Points
  
  if (RoundLosrRec[0][0] == '') RoundLosrRec[0][0] = 0; // Win
  if (RoundLosrRec[0][1] == '') RoundLosrRec[0][1] = 0; // Loss
  if (RoundLosrRec[0][2] == '') RoundLosrRec[0][2] = 0; // Tie
  if (RoundWinrRec[0][3] == '') RoundWinrRec[0][3] = 0; // Points
  
  // Match Tie Result
  if(ResultData[0][colArrRsltTie] == 'Yes' || ResultData[0][colArrRsltTie] == 'Oui'){
    RoundMatchTie = 1;  
  }
  
  // If match is not a Tie
  if(RoundMatchTie == 0){
    // Update Winning Player Results
    RoundWinrRec[0][0] = RoundWinrRec[0][0] + 1;
    // Update Losing Player Results
    RoundLosrRec[0][1] = RoundLosrRec[0][1] + 1;
  }

  // If match is a Tie
  if(RoundMatchTie == 1){
    // Update "Winning" Player Results
    RoundWinrRec[0][2] = RoundWinrRec[0][2] + 1;
    // Update "Losing" Player Results
    RoundLosrRec[0][2] = RoundLosrRec[0][2] + 1;
  }
  
  // Update Points if Disabled
  // Points are equal to the sum of Wins*PtsPerWin + Loss*PtsPerLoss + Ties*PtsPerTie
  if(evntPtsGainedMatch == "Disabled"){
    RoundWinrRec[0][3] = (RoundWinrRec[0][0] * evntPtsPerWin) + (RoundWinrRec[0][1] * evntPtsPerLoss) + (RoundWinrRec[0][2] * evntPtsPerTie);
    RoundLosrRec[0][3] = (RoundLosrRec[0][0] * evntPtsPerWin) + (RoundLosrRec[0][1] * evntPtsPerLoss) + (RoundLosrRec[0][2] * evntPtsPerTie);
  }
  
  // Update Points if Enabled
  if(evntPtsGainedMatch == "Enabled"){
    RoundWinrRec[0][3] = RoundWinrRec[0][3] + MatchDataWinPts;
    RoundLosrRec[0][3] = RoundLosrRec[0][3] + MatchDataLosPts;
  }
  
  // Updates Match Location
  if (evntLocationBonus == "Enabled" && (MatchLoc == 'Yes' || MatchLoc == 'Oui')) {
    RoundWinrLoc = RoundWinrLoc + 1;
    RoundLosrLoc = RoundLosrLoc + 1;
  }
  
  // Update Round Matchups
  // Winning Player
  RoundWinrMatchup = shtRoundRslt.getRange(RoundWinrRow,colRndMatchup).getValue();
  if(RoundWinrMatchup == '') RoundWinrMatchup = MatchDataLosPT;
  else RoundWinrMatchup += ', ' + MatchDataLosPT;
  
  // Losing Player
  RoundLosrMatchup = shtRoundRslt.getRange(RoundLosrRow,colRndMatchup).getValue();
  if(RoundLosrMatchup == '') RoundLosrMatchup = MatchDataWinPT;
  else RoundLosrMatchup += ', ' + MatchDataWinPT;
  
  // Update the Round Results Sheet
  shtRoundRslt.getRange(RoundWinrRow,colRndWin,1,4).setValues(RoundWinrRec);
  shtRoundRslt.getRange(RoundWinrRow,colRndMatchup).setValue(RoundWinrMatchup);
  shtRoundRslt.getRange(RoundLosrRow,colRndWin,1,4).setValues(RoundLosrRec);
  shtRoundRslt.getRange(RoundLosrRow,colRndMatchup).setValue(RoundLosrMatchup);
  
  // If Location Bonus is Enabled
  if (evntLocationBonus == "Enabled"){
    shtRoundRslt.getRange(RoundWinrRow,colRndLocation).setValue(RoundWinrLoc);
    shtRoundRslt.getRange(RoundLosrRow,colRndLocation).setValue(RoundLosrLoc);
  }

  // If Game Type is Wargame and Balance Bonus is Enabled
  if (RoundMatchTie == 0 && evntBalBonus == 'Enabled'){
    // Get Loser Amount of Balance Bonus Points and Increase by value from Config file
    LosrBalBonusVal = shtRoundRslt.getRange(RoundLosrRow,colRndBalBonus).getValue() + evntBalBonusVal;
    shtRoundRslt.getRange(RoundLosrRow,colRndBalBonus).setValue(LosrBalBonusVal);
  }
  
  // Populate Data Posted for Loser
  BalanceDataLosr[0]= 1;
  BalanceDataLosr[1]= RoundLosrRow;
  BalanceDataLosr[2]= colRndBalBonus;
                 
  return BalanceDataLosr;
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
  var colArrRspnDataPrcsd = colArrRspnPrcsd-1;
  var colArrRspnPwd =       cfgColMatchRep[ 1][0]-1;
  var colArrRspnLoc =       cfgColMatchRep[ 2][0]-1;
  var colArrRspnRound =     cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 5][0]-1;
  var colArrRspnWinPts =    cfgColMatchRep[ 6][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 8][0]-1;
  var colArrRspnLosPts =    cfgColMatchRep[ 9][0]-1;
  var colArrRspnTie =       cfgColMatchRep[10][0]-1;
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
  var colArrRspnDataPrcsd = colArrRspnPrcsd-1;
  var colArrRspnRound =     cfgColMatchRep[ 3][0]-1;
  var colArrRspnWinPlyr =   cfgColMatchRep[ 4][0]-1;
  var colArrRspnWinTeam =   cfgColMatchRep[ 5][0]-1;
  var colArrRspnLosPlyr =   cfgColMatchRep[ 7][0]-1;
  var colArrRspnLosTeam =   cfgColMatchRep[ 8][0]-1;
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

