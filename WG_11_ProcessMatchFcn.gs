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

function fcnFindDuplicateData(ss, shtRspn, RspnDataInputs, ResponseData, RspnRow, RspnMaxRows, shtTest) {
  
  Logger.log("Routine: fcnFindDuplicateData");
  
  // Response Data
  var RspnRound = ResponseData[0][3];
  var RspnWinr = ResponseData[0][4];
  var RspnLosr = ResponseData[0][5];

  // Entry Data
  var EntryRound;
  var EntryWinr;
  var EntryLosr;
  var EntryData;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DuplicateRow = 0;
  
  var EntryRoundData = shtRspn.getRange(1, 4, RspnMaxRows-3,1).getValues();
    
  // Loop to find if another entry has the same data
  for (var EntryRow = 1; EntryRow <= RspnMaxRows; EntryRow++){
    
    // Filters only entries of the same Round the response was posted
    if (EntryRoundData[EntryRow][0] == RspnRound){
      
      // Gets Entry Data to analyze
      EntryData = shtRspn.getRange(EntryRow+1, 1, 1, RspnDataInputs).getValues();
      
      EntryRound = EntryData[0][3];
      EntryWinr = EntryData[0][4];
      EntryLosr = EntryData[0][5];
      EntryMatchID = EntryData[0][8];
      EntryPrcssd = EntryData[0][9];
            
      // If both rows are different, the Data Entry was processed and was compiled in the Match Results (Match as a Match ID), Look for player entry combination
      if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID != ''){
        // If combination of players are the same between the entry data and the new response data, duplicate entry was found. Save Row index
        if ((RspnWinr == EntryWinr && RspnLosr == EntryLosr) || (RspnWinr == EntryLosr && RspnLosr == EntryWinr)){
          DuplicateRow = EntryRow + 1;
          EntryRow = RspnMaxRows + 1;
        }
      }
    }
            
    // If we do not detect any value in Round Column, we reached the end of the list and skip
    if (EntryRow <= RspnMaxRows && EntryRoundData[EntryRow][0] == ''){
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

function fcnFindMatchingData(ss, cfgColRspSht, cfgExecData, shtRspn, ResponseData, RspnRow, RspnMaxRows, shtTest) {
  
  Logger.log("Routine: fcnFindMatchingData");

  // Code Execution Options
  var exeDualSubmission = cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  
  // Columns Values and Parameters
  var RspnDataInputs = cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var ColDataConflict = cfgColRspSht[3][0];
  
  var RspnPlyrSubmit = ResponseData[0][1]; // Player Submitting
  var RspnRound = ResponseData[0][3];
  var RspnWinr = ResponseData[0][4];
  var RspnLosr = ResponseData[0][5];

  var EntryData;
  var EntryPlyrSubmit;
  var EntryRound;
  var EntryWinr;
  var EntryLosr;
  var EntryPrcssd;
  var EntryMatchID;
  
  var DataMatchingRow = 0;
  
  var DataConflict = -1;
  
  // Loop to find if the other player posted the game results
      for (var EntryRow = 2; EntryRow <= RspnMaxRows; EntryRow++){
        
        // Gets Entry Data to analyze
        EntryData = shtRspn.getRange(EntryRow, 1, 1, RspnDataInputs).getValues();
        
        EntryPlyrSubmit = EntryData[0][1];
        EntryRound = EntryData[0][3];
        EntryWinr = EntryData[0][4];
        EntryLosr = EntryData[0][5];
        EntryMatchID = EntryData[0][8];
        EntryPrcssd = EntryData[0][9];
        
        // If both rows are different, Round Number, Player A and Player B are matching, we found the other match to compare data to
        if (EntryRow != RspnRow && EntryPrcssd == 1 && EntryMatchID == '' && RspnRound == EntryRound && RspnWinr == EntryWinr && RspnLosr == EntryLosr){

          // If Dual Submission is Enabled, look for Player Submitting, if they are different, continue          
          if ((exeDualSubmission == 'Enabled' && RspnPlyrSubmit != EntryPlyrSubmit) || exeDualSubmission == 'Disabled'){ 
            
            // Compare New Response Data and Entry Data. If Data is not equal to the other, the conflicting Data ID is returned
            DataConflict = subCheckDataConflict(ResponseData, EntryData, 1, RspnDataInputs - 4, shtTest);
            
            // 
            if (DataConflict == 0){
              // Sets Conflict Flag to 'No Conflict'
              shtRspn.getRange(RspnRow, ColDataConflict).setValue('No Conflict');
              shtRspn.getRange(EntryRow, ColDataConflict).setValue('No Conflict');
              DataMatchingRow = EntryRow;
            }
            
            // If Data Conflict was detected, sends email to notify Data Conflict
            if (DataConflict != 0 && DataConflict != -1){
              
              // Sets the Conflict Value to the Data ID value where the conflict was found
              shtRspn.getRange(RspnRow, ColDataConflict).setValue(DataConflict);
              shtRspn.getRange(EntryRow, ColDataConflict).setValue(DataConflict);
            }
          }
        }
        
        // If Dual Submission is Enabled, look for Player Submitting, if they are the same, set negative value of Entry Row as Duplicate          
        if (exeDualSubmission == 'Enabled' && RspnPlyrSubmit == EntryPlyrSubmit){
          DataMatchingRow = 0 - EntryRow;
        }

        // Loop reached the end of responses entered or found matching data
        if(EntryRound == '' || DataMatchingRow != 0) {
          EntryRow = RspnMaxRows + 1;
        }
      }

  return DataMatchingRow;
}


// **********************************************
// function fcnPostMatchResults()
//
// Once both players have submitted their forms 
// the data in the Game Results tab are copied into
// the Round X tab
//
// **********************************************

function fcnPostMatchResultsWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData, shtRspn, ResponseData, MatchingRspnData, MatchID, MatchData, shtTest) {
  
  Logger.log("Routine: fcnPostMatchResultsWG");
  
  // Code Execution Options
  var exeDualSubmission = cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var exePlyrMatchValidation = cfgExecData[2][0];
  
  // League Parameters
  var evntGameType = cfgEvntParam[4][0];
  
  // Cumulative Results sheet variables
  var shtCumul;
  var PwrLvlBonusLosr;
  
  // Match Results Sheet Variables
  var shtRslt = ss.getSheetByName('Match Results');
  var shtRsltMaxRows = shtRslt.getMaxRows();
  var shtRsltMaxCol = shtRslt.getMaxColumns();
  var RsltLastResultRowRng = shtRslt.getRange(3, 4);
  var RsltLastResultRow = RsltLastResultRowRng.getValue() + 1;
  var RsltRng = shtRslt.getRange(RsltLastResultRow, 1, 1, shtRsltMaxCol);
  var ResultData = RsltRng.getValues();
  var MatchValidWinr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var MatchValidLosr = new Array(2); // [0] = Status, [1] = Matches Played by Player used for Error Validation
  var RsltPlyrDataA;
  var RsltPlyrDataB;
  var DataPostedLosr;
  
  var MatchPostedStatus = 0;
  
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
  ResultData[0][2] = ResponseData[0][2];  // Location
  ResultData[0][3] = ResponseData[0][3];  // Round Number
  ResultData[0][4] = ResponseData[0][4];  // Winning Player
  ResultData[0][5] = ResponseData[0][5];  // Losing Player
  ResultData[0][6] = ResponseData[0][6];  // Game is Tie
  
  // If option is enabled, Validate if players are allowed to post results (look for number of games played versus total amount of games allowed
  if (exePlyrMatchValidation == 'Enabled'){
    // Call subroutine to check if players match are valid
    MatchValidWinr = subPlayerMatchValidation(ss, ResultData[0][4], MatchValidWinr, shtTest);
    Logger.log('%s Match Validation: %s',ResultData[0][4], MatchValidWinr[0]);
    
    MatchValidLosr = subPlayerMatchValidation(ss, ResultData[0][5], MatchValidLosr,shtTest);
    Logger.log('%s Match Validation: %s',ResultData[0][5], MatchValidLosr[0]);
  }

  // If option is disabled, Consider Matches are valid
  if (exePlyrMatchValidation == 'Disabled'){
    MatchValidWinr[0] = 1;
    MatchValidLosr[0] = 1;
  }
  
  // If both players have played a valid match
  if (MatchValidWinr[0] == 1 && MatchValidLosr[0] == 1){
    
    // Copies Match ID
    ResultData[0][1] = MatchID; // Match ID
    
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
    ResultData[0][shtRsltMaxCol-1] = '= if(INDIRECT("R[0]C[-6]",FALSE)<>"",1,"")';
    RsltRng.setValues(ResultData);
    
    // Update the Match Posted Status
    MatchPostedStatus = 1;
    
    // Post Results in Appropriate Round Number for Both Players
    // DataPostedLosr is an Array with [0]=Post Status (1=Success) [1]=Loser Row [2]=Power Level Column
    DataPostedLosr = fcnPostResultRoundWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, ResultData, shtTest);
    
    // Gets New Power Level / Points Bonus for Loser from Cumulative Results Sheet
    shtCumul = ss.getSheetByName('Cumulative Results');
    PwrLvlBonusLosr = shtCumul.getRange(DataPostedLosr[1],DataPostedLosr[2]).getValue();
    Logger.log('Cumulative Power Level: %s',PwrLvlBonusLosr);
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
  MatchData[0][0] = ResponseData[0][0]; // TimeStamp
  MatchData[0][0] = Utilities.formatDate (MatchData[0][0], Session.getScriptTimeZone(), 'YYYY-MM-dd HH:mm:ss');
  
  MatchData[1][0] = ResponseData[0][2];  // Location (Store Y/N)
  MatchData[2][0] = MatchID;             // MatchID
  MatchData[3][0] = ResponseData[0][3];  // Round Number
  MatchData[4][0] = ResponseData[0][4];  // Winning Player
  MatchData[4][1] = MatchValidWinr[1];   // Winning Player Matches Played
  MatchData[5][0] = ResponseData[0][5];  // Losing Player
  MatchData[5][1] = MatchValidLosr[1];   // Losing Player Matches Played
  MatchData[5][2] = PwrLvlBonusLosr;
  MatchData[6][0] = ResponseData[0][6];  // Game is Tie
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

function fcnPostResultRoundWG(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, ResultData, shtTest) {
  
  Logger.log("Routine: fcnPostResultRoundWG");

  // Column Values
  var colPlyr = cfgColRndSht[0][0];
  var colTeam = cfgColRndSht[1][0];
  var colWin = cfgColRndSht[3][0];
  var colLos = cfgColRndSht[4][0];
  var colTie = cfgColRndSht[5][0];
  var colPoints = cfgColRndSht[6][0];
  var colWinPerc = cfgColRndSht[7][0];
  var colLocation = cfgColRndSht[8][0];
  var colBalanceBonus = cfgColRndSht[10][0];
  
  // League Parameters
  var evntGameType = cfgEvntParam[4][0];
  var evntBalance = cfgEvntParam[21][0];
  var evntBalanceBonus = cfgEvntParam[22][0];
  
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
  
  var LosrPowerLevel;
  
  var RoundWinrRow = 0;
  var RoundLosrRow = 0;
  var RoundMatchTie = 0; // Match is not a Tie by default
  var DataPostedLosr = new Array(3);
  
  var MatchLoc = ResultData[0][2];
  var MatchRound = ResultData[0][3];
  var MatchDataWinr = ResultData[0][4];
  var MatchDataLosr = ResultData[0][5];
  var MatchDataTie  = ResultData[0][6];
  
  var Round = 'Round'+MatchRound;
  shtRoundRslt = ss.getSheetByName(Round);
  
  shtRoundMaxCol = shtRoundRslt.getMaxColumns();

  // Gets All Players Names
  RoundPlyrList = shtRoundRslt.getRange(5,colPlyr,32,1).getValues();
  
  // Find the Winning and Losing Player in the Round Result Tab
  for (var RsltRow = 5; RsltRow <= 36; RsltRow ++){
    
    if (RoundPlyrList[RsltRow - 5][0] == MatchDataWinr) RoundWinrRow = RsltRow;
    if (RoundPlyrList[RsltRow - 5][0] == MatchDataLosr) RoundLosrRow = RsltRow;
    
    if (RoundWinrRow != '' && RoundLosrRow != '') {
      // Get Winner and Loser Match Record, 3 values, Win, Loss, Ties
      RoundWinrRec = shtRoundRslt.getRange(RoundWinrRow,colWin,1,3).getValues();
      RoundWinrLoc = shtRoundRslt.getRange(RoundWinrRow,colLocation).getValue();
      RoundLosrRec = shtRoundRslt.getRange(RoundLosrRow,colWin,1,3).getValues();
      RoundLosrLoc = shtRoundRslt.getRange(RoundLosrRow,colLocation).getValue();
            
      RsltRow = 37;
    }
  }
  
  // Fill Empty Cells for both Winner and Loser
  if (RoundWinrRec[0][0] == '') RoundWinrRec[0][0] = 0; 
  if (RoundWinrRec[0][1] == '') RoundWinrRec[0][1] = 0; 
  if (RoundWinrRec[0][2] == '') RoundWinrRec[0][2] = 0; 
  if (RoundLosrRec[0][0] == '') RoundLosrRec[0][0] = 0; 
  if (RoundLosrRec[0][1] == '') RoundLosrRec[0][1] = 0; 
  if (RoundLosrRec[0][2] == '') RoundLosrRec[0][2] = 0; 
  
  // Match Tie Result
  if(ResultData[0][6] == 'Yes' || ResultData[0][6] == 'Oui'){
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
  
  // Updates Match Location
  if (MatchLoc == 'Yes' || MatchLoc == 'Oui') {
    RoundWinrLoc = RoundWinrLoc + 1;
    RoundLosrLoc = RoundLosrLoc + 1;
  }
  
  // Update the Round Results Sheet
  shtRoundRslt.getRange(RoundWinrRow,colWin,1,3).setValues(RoundWinrRec);
  shtRoundRslt.getRange(RoundWinrRow,colLocation).setValue(RoundWinrLoc);
  shtRoundRslt.getRange(RoundLosrRow,colWin,1,3).setValues(RoundLosrRec);
  shtRoundRslt.getRange(RoundLosrRow,colLocation).setValue(RoundLosrLoc);
  

  // If Game Type is Wargame
  if (RoundMatchTie == 0 && evntBalance == 'Enabled'){
    // Get Loser Amount of Power Level Bonus and Increase by value from Config file
    LosrPowerLevel = shtRoundRslt.getRange(RoundLosrRow,colBalanceBonus).getValue() + evntBalanceBonus;
    shtRoundRslt.getRange(RoundLosrRow,colBalanceBonus).setValue(LosrPowerLevel);
  }
  
  // Populate Data Posted for Loser
  DataPostedLosr[0]= 1;
  DataPostedLosr[1]= RoundLosrRow;
  DataPostedLosr[2]= colBalanceBonus;
                 
  return DataPostedLosr;
}




