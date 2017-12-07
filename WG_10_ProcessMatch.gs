// **********************************************
// function fcnProcessMatchWG()
//
// This function copies the new Match Report response
// to the Match Report Process Queue and executes
// the AnalyzeResult function until Process Queue is 
// Empty
//
// **********************************************

function fcnProcessMatchWG() {
  
  Logger.log("Routine: fcnProcessMatchWG");
  // Opens Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  //  var shtTest = ss.getSheetByName('Test');
  
  // Config Sheet to get options
  var shtConfig = ss.getSheetByName('Config');
  var cfgEvntParam =    shtConfig.getRange( 4, 4,48,1).getValues();
  var cfgColRspSht =    shtConfig.getRange( 4,18,16,1).getValues();
  var cfgColRndSht =    shtConfig.getRange( 4,21,16,1).getValues();
  var cfgExecData  =    shtConfig.getRange( 4,24,16,1).getValues();
  var cfgColMatchRep =  shtConfig.getRange( 4,31,20,1).getValues();
  var cfgColMatchRslt = shtConfig.getRange(21,18,32,1).getValues();
  
  // Execution Options
  var exeSendEmail = cfgExecData[5][0];
  var exeTrigReport = cfgExecData[4][0];
  
  // Column Values and Parameters
  var RspnDataInputs =      cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var colMatchID =          cfgColRspSht[1][0];
  var colDataPrcsd =        cfgColRspSht[2][0];
  var colMatchIDLastVal =   cfgColRspSht[6][0];
  var colNextEmptyRow =     cfgColRspSht[7][0];
  var colNbUnprcsdEntries = cfgColRspSht[8][0];
  
  // Column Values for Data in Response Sheet
  var colArrRspnDataPrcsd = colDataPrcsd-1;
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
  
  // Event Parameters
  var evntFormat =        cfgEvntParam[ 9][0];
  var evntPassword =      cfgEvntParam[12][0];
  var evntRoundDuration = cfgEvntParam[13][0];
  
  // Get Sheet IDs
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Get Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');

  // Get Number of Players and Players Email
  var shtPlayers = ss.getSheetByName('Players');
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayersEmail = shtPlayers.getRange(3,3,NbPlayers,1).getValues();
  
  // Open Responses sheets
  var shtRspn = ss.getSheetByName('Responses');
  var shtRspnEN = ss.getSheetByName('MatchResp EN');
  var shtRspnFR = ss.getSheetByName('MatchResp FR');

  var RspnMaxRowsEN = shtRspnEN.getMaxRows();
  var RspnMaxRowsFR = shtRspnFR.getMaxRows();
  
  // Function Variables
  var ResponseData;
  var DataCopiedStatus = 0;
  var TimeStamp;
  var Password;
  var PasswordValid = 0;
  var RspnRow;
  var RoundNum;
  var PT1; // Player/Team 1
  var PT2; // Player/Team 2
  
  var EmailAddresses = subCreateArray(3,2);
  
  // Data Processing Flags
  var Status = new Array(2); // Status[0] = Status Value, Status[1] = Status Message
  
  // Function Polled Values
  var RspnNextRow = shtRspn.getRange(1, colNextEmptyRow).getValue();
  var EntriesProcessing;
  var RspnNextRowEN = shtRspnEN.getRange(1, colNextEmptyRow).getValue();
  var RspnNextRowFR = shtRspnFR.getRange(1, colNextEmptyRow).getValue();
    
  // Execute if Trigger is Enabled
  if(exeTrigReport == 'Enabled'){

    Logger.log('------- New Match Report Received -------');
    var MatchProcessLog = 'TCG Match Process Log - ' + shtConfig.getRange(4,2).getValue() + ' - Entry Row EN: ' + RspnNextRowEN + ' - Entry Row FR: ' + RspnNextRowFR;
    Logger.log(MatchProcessLog);
    
    EntriesProcessing = shtRspn.getRange(1, colNbUnprcsdEntries).getValue();
    Logger.log('Nb of Entries Before Copying: %s',EntriesProcessing);
    
    // Look for Unprocessed Data in MatchResp EN
    for (RspnRow = RspnNextRowEN; RspnRow <= RspnMaxRowsEN; RspnRow++){
      
      Logger.log('Row: %s',RspnRow)
      // Copy the new response data (from Time Stamp to Data Copied Field)
      ResponseData = shtRspnEN.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
      TimeStamp = ResponseData[0][0];
      Password = ResponseData[0][colArrRspnPwd];
      RoundNum = ResponseData[0][colArrRspnRound];
      DataCopiedStatus = ResponseData[0][colArrRspnDataPrcsd];
      
      if(evntFormat == "Single"){
        PT1 = ResponseData[0][colArrRspnPlyr1];
        PT2 = ResponseData[0][colArrRspnPlyr2];
      }
      if(evntFormat == "Team"){
        PT1 = ResponseData[0][colArrRspnTeam1];
        PT2 = ResponseData[0][colArrRspnTeam2];
      }
      
      // If TimeStamp is null, Delete Row and start over
      if(TimeStamp == '' && RspnRow < RspnMaxRowsEN) {
        shtRspnEN.deleteRow(RspnRow);
        RspnRow = RspnNextRowEN - 1;
        RspnMaxRowsEN = shtRspnEN.getMaxRows();
      }

      // If Timestamp is not null, analyze response data
      if(TimeStamp != ''){
        // Look if Password is valid
        Logger.log('Password Entered: %s', Password);
        Logger.log('Event Password: %s', evntPassword);
        if(Password == evntPassword) PasswordValid = 1; 
        
        // Check if DataCopied Field is null and Password is Valid, we found new data to copy
        if (DataCopiedStatus == '' && PasswordValid == 1){
          Logger.log('Password Valid, Data Copied to Responses');
          DataCopiedStatus = 'Data Copied';
          shtRspnEN.getRange(RspnRow, colDataPrcsd).setValue(DataCopiedStatus);
          // Creates formula to update Last Entry Processed
          shtRspnEN.getRange(RspnRow, colNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-'+ colMatchIDLastVal +']",FALSE)<>"",1,"")');
        }
        
        // If Password is not Valid, update Data Copied and Next Empty Row Cells
        if (TimeStamp != '' && PasswordValid == 0){
          Logger.log('Password Not Valid');
          DataCopiedStatus = 'Password Not Valid';
          shtRspnEN.getRange(RspnRow, colDataPrcsd).setValue(DataCopiedStatus);
          // Creates formula to update Last Entry Processed
          shtRspnEN.getRange(RspnRow, colNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-'+ colMatchIDLastVal +']",FALSE)<>"",1,"")');
        }
        // If Data is copied or Password is not Valid or TimeStamp is null, Exit loop MatchResp EN and Loop through MatchResp FR
        if (DataCopiedStatus == 'Data Copied' || DataCopiedStatus == 'Password Not Valid' || (TimeStamp == '' && RspnRow >= RspnMaxRowsEN)) {
          RspnRow = RspnMaxRowsEN + 1;
          if(TimeStamp == '' && RspnRow >= RspnMaxRowsEN) DataCopiedStatus = 0;
        }
      }
    }
    
    // Executes MatchResp FR loop only if MatchResp EN did not find anything
    if (DataCopiedStatus == 0){
      
      // Look for Unprocessed Data in MatchResp FR
      for (RspnRow = RspnNextRowFR; RspnRow <= RspnMaxRowsFR; RspnRow++){
        
      // Copy the new response data (from Time Stamp to Data Copied Field)
        ResponseData = shtRspnFR.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
        TimeStamp = ResponseData[0][0];
        Password = ResponseData[0][colArrRspnPwd];
        RoundNum = ResponseData[0][colArrRspnRound];
        DataCopiedStatus = ResponseData[0][colArrRspnDataPrcsd];

        // If TimeStamp is null, Delete Row and start over
        if (TimeStamp == '' && RspnRow < RspnMaxRowsFR) {
          shtRspnFR.deleteRow(RspnRow);
          RspnRow = RspnNextRowFR - 1;
          RspnMaxRowsFR = shtRspnFR.getMaxRows();
        }
        
        // If Timestamp is not null, analyze response data
        if(TimeStamp != ''){
          // Look if Password is valid
          Logger.log('Password Entered: %s', Password);
          if(Password == evntPassword) PasswordValid = 1;
          
          // Check if DataCopied Field is null and Password is Valid, we found new data to copy
          if (TimeStamp != '' && DataCopiedStatus == '' && PasswordValid == 1){
            Logger.log('Password Valid, Data Copied to Responses');
            DataCopiedStatus = 'Data Copied';
            shtRspnFR.getRange(RspnRow, colDataPrcsd).setValue(DataCopiedStatus);
            // Creates formula to update Last Entry Processed
            shtRspnFR.getRange(RspnRow, colNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-'+ colMatchIDLastVal +']",FALSE)<>"",1,"")');
          }
          
          // If Password is not Valid, update Data Copied and Next Empty Row Cells
          if (TimeStamp != '' && PasswordValid == 0){
            Logger.log('Password Not Valid');
            DataCopiedStatus = 'Password Not Valid';
            shtRspnFR.getRange(RspnRow, colDataPrcsd).setValue(DataCopiedStatus);
            // Creates formula to update Last Entry Processed
            shtRspnFR.getRange(RspnRow, colNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-'+ colMatchIDLastVal +']",FALSE)<>"",1,"")');
          }
          // If Data is copied or Password is not Valid or TimeStamp is null, Exit loop MatchResp FR to process data
          if (DataCopiedStatus == 'Data Copied' || DataCopiedStatus == 'Password Not Valid' || (TimeStamp == '' && RspnRow >= RspnMaxRowsFR)) {
            RspnRow = RspnMaxRowsFR + 1;
          }
        }
      }
    }
    
    // If Data is copied, put it in Responses Sheet
    if (DataCopiedStatus == 'Data Copied'){
      
      // Copy New Entry Data to Main Responses Sheet
      shtRspn.getRange(RspnNextRow, 1, 1, RspnDataInputs).setValues(ResponseData);
      
      Logger.log('Match Data Copied for Players: %s, %s',PT1,PT2);
      
      // Copy Formula to detect if an entry is currently processing
      shtRspn.getRange(RspnNextRow, colNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-'+ colMatchIDLastVal +']",FALSE)<>"",1,"")');
      shtRspn.getRange(RspnNextRow, colNbUnprcsdEntries).setValue('=IF(AND(INDIRECT("R[0]C[-'+ colNextEmptyRow +']",FALSE)<>"",INDIRECT("R[0]C[-4]",FALSE)<>2),1,"")');
      
      // Troubleshoot
      EntriesProcessing = shtRspn.getRange(1, colNbUnprcsdEntries).getValue();
      Logger.log('Nb of Entries Pending After Copying Match Data: %s',EntriesProcessing)
      
      // Make sure that we only execute this loop on the first instance call
      if (EntriesProcessing == 1){
        // Execute Game Results Analysis for as long as there are unprocessed entries
        while (EntriesProcessing >= 1) {
          Status = fcnAnalyzeResultsWG(ss, shtConfig, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRep, cfgExecData, shtRspn);
          EntriesProcessing = shtRspn.getRange(1, colNbUnprcsdEntries).getValue();
          Logger.log('Nb of Entries Pending After Processing: %s',EntriesProcessing)
        }
      }
      // If the Match was successfully Posted, Update League Standings
      if (Status[0] == 10){
        Logger.log('--------- Updating Standings ---------');
        Logger.log('Update Standings');
        // Execute Ranking function in Standing tab
        fcnUpdateStandings(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData);
        Logger.log('Copy to League Spreadsheets');
        // Copy all data to League Spreadsheet
        fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, Status[2], 0);
        Logger.log('------------ Standings Updated ------------');
      }
    }
  }
  
  // Send Error Email to sender if Email is not valid
//  if(PasswordValid == 0){
//    Logger.log('Password Not Valid : %s', Password);
//    // Get Emails from both players
//    EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, ResponseData[0][4], ResponseData[0][5]);
//    fcnMatchReportPwdError(shtConfig, EmailAddresses);
//  }
  
  // Post Log to Log Sheet
  subPostLog(shtLog,Logger.getLog());
  
}


// **********************************************
// function fcnAnalyzeResultsWG()
//
// This function analyzes the Match Results 
// once a player submitted his Form to populate the 
// Event Sheets
//
// **********************************************

function fcnAnalyzeResultsWG(ss, shtConfig, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgColMatchRep, cfgExecData, shtRspn) {
  
  Logger.log("Routine: fcnAnalyzeResultsWG");
  
  // Data from Configuration File
    
  // Code Execution Options
  var exeDualSubmission =      cfgExecData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var exePostMatchResult =     cfgExecData[1][0];
  var exePlyrMatchValidation = cfgExecData[2][0];
  var exeSendEmail =           cfgExecData[5][0];
  var exeUpdatePlyrDB =        cfgExecData[6][0];
  var exeMemberProfileLink =   cfgExecData[7][0];
  
  // Columns Values and Parameters
  var RspnDataInputs =      cfgColRspSht[0][0]; // from Time Stamp to Data Processed
  var colMatchID =          cfgColRspSht[1][0];
  var colDataPrcsd =        cfgColRspSht[2][0];
  var colDataConflict =     cfgColRspSht[3][0];
  var colStatus =           cfgColRspSht[4][0];
  var colStatusMsg =        cfgColRspSht[5][0];
  var colMatchIDLastVal =   cfgColRspSht[6][0];
  var colNextEmptyRow =     cfgColRspSht[7][0];
  var colNbUnprcsdEntries = cfgColRspSht[8][0];
  
  // Column Values for Data in Response Sheet
  var colArrRspnDataPrcsd = colDataPrcsd-1;
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
  
  // Event Parameters
  var evntNameEN =         cfgEvntParam[ 0][0] + ' ' + cfgEvntParam[7][0];
  var evntNameFR =         cfgEvntParam[ 8][0] + ' ' + cfgEvntParam[0][0];
  var evntGameType =       cfgEvntParam[ 4][0];
  var evntGameSystem =     cfgEvntParam[ 5][0];
  var evntFormat =         cfgEvntParam[ 9][0];
  var evntRoundDuration =  cfgEvntParam[13][0];
  var evntLocationBonus =  cfgEvntParam[23][0];
  var evntBalanceBonus =   cfgEvntParam[21][0];
  var evntNbCardPack =     cfgEvntParam[25][0];
  var evntPtsGainedMatch = cfgEvntParam[27][0];
  var evntTiePossible =    cfgEvntParam[31][0];
    
  // Form Responses Sheet Variables
  var RspnMaxRows = shtRspn.getMaxRows();
  var RspnMaxCols = shtRspn.getMaxColumns();
  var RspnNextRowPrcss = shtRspn.getRange(1, colNextEmptyRow).getValue() - shtRspn.getRange(1, colNbUnprcsdEntries).getValue();
  var ResponseData;
  var MatchingRspnData;
  
  // Match Data Variables
  var MatchID; 
  var MatchData = subCreateArray(26,4);
  // [0][0]= TimeStamp, [0][1]= Event Name EN, [0][2]= Event Name FR
  // [1][0]= MatchID,   [1][1]= Game System
  // [2][0]= Round Number
  // [3][0]= Winning Player / Team, [3][1]= Points, [3][2]= Matches Played 
  // [4][0]= Losing Player / Team,  [4][1]= Points, [4][2]= Matches Played, [4][3]= Total Balance Bonus Value
  // [5][0]= Game Tie (Yes or No)
  // [6][0]= Location Bonus
  // [7][0]= Balance Bonus
  // [8-23] = Not Used
  // [25] = MatchPostStatus
      
  // Email Addresses Array
  var EmailAddresses = subCreateArray(3,2);
  // [0][0]= Administrator Language Preference
  // [1][0]= Winning Player Language Preference
  // [2][0]= Losing Player Language Preference
  
  // [0][1]= Administrator Email Address
  // [1][1]= Winning Player Email Address 
  // [2][1]= Losing Player Email Address

  EmailAddresses[0][0] = 'English';
  EmailAddresses[0][1] = 'turn1glt@gmail.com';
  EmailAddresses[1][1] = '';
  EmailAddresses[2][1] = '';

  // Data Processing Flags
  var Status = new Array(2); // [0]= Status Value, [1]= Status Message
  Status[0] = 0;
  
  // Log Status for Players/Teams
  var logStatusPT1 = new Array(3); // [0]= Status Value, [1]= Status Message, [2]= Player
  logStatusPT1[0] = 0;
  logStatusPT1[1] = '';
  logStatusPT1[2] = '';
  var logStatusPT2 = new Array(3); // [0]= Status Value, [1]= Status Message, [2]= Player
  logStatusPT2[0] = 0;
  logStatusPT2[1] = '';
  logStatusPT2[2] = '';
  
  // Routine Variables
  var PT1; // Player/Team 1
  var PT2; // Player/Team 2
  
  var DuplicateRspn = -99;
  var MatchingRspn = -98;
  var MatchPostStatus = -97;
  var CardDBUpdated = -96;
  
  var shtTest = ss.getSheetByName("Test");
  
  Logger.log('--------- Posting Match ---------'); 
  Logger.log('--------- Options ---------');
  Logger.log('Dual Submission Option: %s',exeDualSubmission);
  Logger.log('Post Results Option: %s',exePostMatchResult);
  Logger.log('Player Match Validation Option: %s',exePlyrMatchValidation);
  Logger.log('Game Type: %s',evntGameType);
  Logger.log('Send Email Option: %s',exeSendEmail);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = shtRspn.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    // Values from Response Data
    var RspnDataPwd      = ResponseData[0][colArrRspnPwd];    // Password
    var RspnDataRoundNum = ResponseData[0][colArrRspnRound];  // Round Number
    var RspnDataPlyr1    = ResponseData[0][colArrRspnPlyr1];  // Winning Player
    var RspnDataTeam1    = ResponseData[0][colArrRspnTeam1];  // Winning Team
    var RspnDataPts1     = ResponseData[0][colArrRspnPts1];   // Winning Points
    var RspnDataPlyr2    = ResponseData[0][colArrRspnPlyr2];  // Losing Player
    var RspnDataTeam2    = ResponseData[0][colArrRspnTeam2];  // Losing Team
    var RspnDataPts2     = ResponseData[0][colArrRspnPts2];   // Losing Points
    var RspnDataTie      = ResponseData[0][colArrRspnTie];    // Tie
    var RspnDataLocation = ResponseData[0][colArrRspnLoc];    // Match Location (Store Yes or No)
    
    var RspnDataPrcssd     = ResponseData[0][colArrRspnDataPrcsd];// Data Processed Status
    var RspnDataPlyrSubmit = ResponseData[0][colArrRspnPlyrSub];  // Player Submitting Match Report
    
    // Select Player or Team Depending on Event Format (Single or Team)
    switch(evntFormat){
      case "Single": {
        PT1 = RspnDataPlyr1;
        PT2 = RspnDataPlyr2;
        break;
      }
      case "Team": {
        PT1 = RspnDataTeam1;
        PT2 = RspnDataTeam2;
        break;
      }
    }
    
    Logger.log('Players: %s, %s',PT1,PT2);
    
    // If Round number is not empty and Processed is empty, Response Data needs to be processed
    if (RspnDataRoundNum != '' && RspnDataPrcssd == ''){
      
      // If both Players in the response are different, continue
      if (PT1 != PT2){
        
        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 1;
          Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
        }
                
        // Generates the Match ID in advance if data analysis is successful
        MatchID = shtRspn.getRange(1, colMatchIDLastVal).getValue() + 1;
        
        Logger.log('New Data Found at Row: %s',RspnRow);

        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 2;
          Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
        }
        
        // Look for Duplicate Entry (looks in all entries with MatchID and combination of Round Number, Winner and Loser) 
        // Real code will look at Player Posting Data as well
        DuplicateRspn = fcnFindDuplicateData(ss, shtRspn, cfgEvntParam, cfgColRspSht, cfgColMatchRep, RspnDataInputs, ResponseData, RspnRow, RspnMaxRows);  
        if(DuplicateRspn == 0) Logger.log('No Duplicate Found');
        if(DuplicateRspn > 0 ) Logger.log('Duplicate Found at Row: %s', DuplicateRspn);
        
        Logger.log("Routine: fcnAnalyzeResultsWG");
        
        // FindDuplicateEntry function was executed properly and didn't find any Duplicate entry, continue analyzing the response data
        if (DuplicateRspn == 0){
          
          // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
          if (exeDualSubmission == 'Enabled'){
            
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 3; 
              Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
            }
            // function returns row where the matching data was found
            MatchingRspn = fcnFindMatchingData(ss, cfgEvntParam, cfgColRspSht, cfgColMatchRep, cfgExecData, shtRspn, ResponseData, RspnRow, RspnMaxRows);
            Logger.log("Routine: fcnAnalyzeResultsWG");
            if (MatchingRspn < 0) DuplicateRspn = 0 - MatchingRspn;
          }
          
          // Search if the other Entry matching this response has been submitted
          if (exeDualSubmission == 'Disabled'){
            MatchingRspn = RspnRow;
          }      
          
          Logger.log('Matching Result: %s', MatchingRspn);
          
          // If the result of the fcnFindMatchingEntry function returns something greater than 0, we found a matching entry, continue analyzing the response data
          if (MatchingRspn > 0){
            
            if (exePostMatchResult == 'Enabled'){
              
              // Get the Entry Data found at row MatchingRspn
              MatchingRspnData = shtRspn.getRange(MatchingRspn, 1, 1, RspnDataInputs).getValues();
              
              // Updates the Status while processing
              if(Status[0] >= 0){
                Status[0] = 4; 
                Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
              }
              // Pre-Populate Match Data
              MatchData[0][0] = Utilities.formatDate (ResponseData[0][0], Session.getScriptTimeZone(), 'YYYY-MM-dd HH:mm:ss'); // TimeStamp
              MatchData[0][1] = evntNameEN;          // Event Name English 
              MatchData[0][2] = evntNameFR;          // Event Name French
              
              MatchData[1][0] = MatchID;             // MatchID
              MatchData[1][1] = evntGameSystem;      // Game System (ie Magic, Warhammer 40k etc)
              MatchData[2][0] = ResponseData[0][colArrRspnRound];  // Round Number
              
              MatchData[3][0] = PT1;  // Player/Team 1
              MatchData[4][0] = PT2;  // Player/Team 2

                // If Event uses Points Gained per Match
              if(evntPtsGainedMatch == "Enabled"){
                MatchData[3][1] = ResponseData[0][colArrRspnPts1];  // Player/Team 1 Points
                MatchData[4][1] = ResponseData[0][colArrRspnPts2];  // Player/Team 2 Points
              }
              // If Event doesn't use Points Gained per Match
              if(evntPtsGainedMatch == "Disabled"){
                MatchData[3][1] = "-";  // Player/Team 1 Points
                MatchData[4][1] = "-";  // Player/Team 2 Points
              }
              
              if(evntTiePossible == "Enabled")   MatchData[5][0] = ResponseData[0][colArrRspnTie];  // Game is a Tie
              if(evntLocationBonus == "Enabled") MatchData[6][0] = ResponseData[0][colArrRspnLoc];  // Location (Store Y/N)
              
              // Execute function to populate Match Result Sheet from Response Sheet
              MatchData = fcnPostMatchResultsWG(ss, shtConfig, ResponseData, MatchingRspnData, MatchID, MatchData);
              MatchPostStatus = MatchData[25][0];
              
              shtTest.getRange(2, 2, 26, 4).setValues(MatchData);
              
              Logger.log("Routine: fcnAnalyzeResultsWG");
              
              Logger.log('Match Post Status: %s',MatchPostStatus);
              
              // If Match was populated in Match Results Tab
              if (MatchPostStatus == 1){
                // Match ID doesn't change because we assumed it was already OK
                Logger.log('Match Posted ID: %s',MatchID);
                
                // Log Players Match Data
                logStatusPT1[2] = PT1;
                logStatusPT1 = fcnLogEventMatch(ss, shtConfig, cfgEvntParam, logStatusPT1, MatchData);
                if(exeMemberProfileLink == "Enabled") fcnLogMemberMatch(ss, shtConfig, logStatusPT1, MatchData);
                Logger.log('Event Player Record Status for %s : %s',logStatusPT1[2],logStatusPT1[1]);
                
                logStatusPT2[2] = PT2;
                logStatusPT2 = fcnLogEventMatch(ss, shtConfig, cfgEvntParam, logStatusPT2, MatchData);
                if(exeMemberProfileLink == "Enabled") fcnLogMemberMatch(ss, shtConfig, logStatusPT2, MatchData);
                Logger.log('Event Player Record Status for %s : %s',logStatusPT2[2],logStatusPT2[1]);
                
                // If Event Game Type is Wargame
                if(evntGameType == 'Wargame'){
                  // Updates the Status while processing
                  if(Status[0] >= 0){
                    Status[0] = 5; 
                    Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
                  }
                  // Update Player Army DB and Army List
                  if(exeUpdatePlyrDB == 'Enabled' && evntBalanceBonus == 'Enabled') fcnUpdateArmyDB(ss, shtConfig, cfgColRndSht, PT2);
                  Logger.log("Routine: fcnAnalyzeResultsWG");
                }
              }
              
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
              
              // If MatchPostSuccess = 0, function was executed but was not able to post in the Match Result Tab
              if (MatchPostStatus < 0){
                // Updates the Match ID to an empty value 
                MatchID = '';
                // Generate the Status Message
                Status = subGenErrorMsg(Status, MatchPostStatus,0);
              }
            }
            // If Posting is disabled, generate Match ID for testing        
            if (exePostMatchResult == 'Disabled'){
              // Match ID doesn't change because we assumed it was already OK
              
            }
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 6; 
              Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
            }
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          }
          
          // If MatchingEntry = 0, fcnFindMatchingEntry did not find a matching entry, it might be the first response entry
          if (exeDualSubmission == 'Enabled' && MatchingRspn == 0){
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 0;
              Status[1] = 'Waiting for Other Response Submission';
            }
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          } 
          
          // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, send email to notify
          if (exeDualSubmission == 'Enabled' && MatchingRspn == -1){
            // Set the Status Message
            Status = subGenErrorMsg(Status, MatchingRspn,0);
          }
        }
        
        // If Duplicate is found, send email to notify, set Response Data Processed to -1 to represent the Duplicate Entry
        if (DuplicateRspn > 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;        
          
          // Sets the Status Message
          Status = subGenErrorMsg(Status, -10,DuplicateRspn);
        }
        
        // If FindDuplicateEntry was not executed properly, send email to notify, set Response Data Processed to -2 to represent processing error
        if (DuplicateRspn < 0){
          
          // Updates the Match ID to an empty value 
          MatchID = '';
          
          // Set the Data Processed Flag
          RspnDataPrcssd = 1;  
          
          // Set the Status Message
          Status = subGenErrorMsg(Status, DuplicateRspn,0);
        }
      } 
      
      // If Both Players are the same, report error
      if (PT1 == PT2){
        
        // Updates the Match ID to an empty value 
        MatchID = '';
        
        // Set the Data Processed Flag
        RspnDataPrcssd = 1;  
        
        // Set the Status Message
        Status = subGenErrorMsg(Status, -50,0);
      }
      
      Logger.log('Match Post Status: %s - %s',Status[0], Status[1])
      
      // Call the Email Function, sends Match Data if Send Email Option is Enabled
      if(Status[0] >= 0 && exeSendEmail == 'Enabled') {
        
        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 7; 
          Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
        }
        // Get Email addresses from Config File
        EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, PT1, PT2);
        Logger.log("Routine: fcnAnalyzeResultsWG");
        
        // Send email to players. Each function analyzes language preferences
        fcnSendConfirmEmail(shtConfig, EmailAddresses, MatchData);
        Logger.log('Confirmation Emails Sent');
      }
      
      // If an Error has been detected that prevented to process the Match Data, send available data and Error Message
      if(Status[0] < 0 && exeSendEmail == 'Enabled') {
      
        // Get Email addresses from Config File
        EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, PT1, PT2);
        
        // Send Error Message, each function analyzes language preferences
        fcnSendErrorEmail(shtConfig, EmailAddresses, MatchData, MatchID, Status);
        Logger.log('Error Emails Sent');
      }
      
      // Updates the Status while processing
      if(Status[0] >= 0){
        Status[0] = 9; 
        Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
      }
      // Set the Match ID (for both Response and Matching Entry), and Updates the Last Match ID generated, 
      if (MatchPostStatus == 1 || exePostMatchResult == 'Disabled'){
        shtRspn.getRange(RspnRow, colMatchID).setValue(MatchID);
        shtRspn.getRange(1, colMatchIDLastVal).setValue(MatchID);
      }
      
      // Updates the Status while processing
      if(Status[0] >= 0){
        Status[0] = 10; 
        Status[1] = subUpdateStatus(shtRspn, RspnRow, colStatus, colStatusMsg, Status[0]);
        Status[2] = MatchData[2][0]; // Round Processed
      }
      // Updating Match Process Data
      shtRspn.getRange(RspnRow, colDataPrcsd).setValue(RspnDataPrcssd);
      shtRspn.getRange(RspnRow, colNbUnprcsdEntries).setValue(0);
      
      // If Process Error has been detected, update the Response Process Data
      if(Status[0] < 0){
        shtRspn.getRange(RspnRow, colStatus).setValue(Status[0]);
        shtRspn.getRange(RspnRow, colStatusMsg).setValue(Status[1]);
      }
      
      // Set the Matching Response Match ID if Matching Response found
      if (MatchingRspn > 0) shtRspn.getRange(MatchingRspn, colMatchID).setValue(MatchID);	  
            
    }
    // When Round Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
    if(RspnDataRoundNum == '' || RspnDataPrcssd == 1) {
      Logger.log('Response Loop exit at Row: %s',RspnRow)
      RspnRow = RspnMaxRows + 1;
    }
  }
  
  Logger.log('------------ Game Posted ------------');
    
  return Status;
}


