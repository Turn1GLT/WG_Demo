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
  var ConfigData = shtConfig.getRange(56,2,32,1).getValues();
  var OptSendEmail = ConfigData[6][0];
  var cfgTrigReport = ConfigData[9][0];
  
  // Columns Values and Parameters
  var ColDataCopied = ConfigData[21][0];
  var ColNextEmptyRow = ConfigData[26][0];
  var ColNbUnprcsdEntries = ConfigData[27][0];
  var RspnDataInputs = ConfigData[28][0]; // from Time Stamp to Data Processed
    
  // League Duration
  var LgParamDuration = shtConfig.getRange(10,9).getValue();
  
  // Get Log Sheet
  var shtIDs = shtConfig.getRange(17,7,20,1).getValues();
  var shtLog = SpreadsheetApp.openById(shtIDs[14][0]).getSheetByName('Log');

  // Get Number of Players and Players Email
  var shtPlayers = ss.getSheetByName('Players');
  var NbPlayers = shtPlayers.getRange('F2').getValue();
  var PlayersEmail = shtPlayers.getRange(3,3,NbPlayers,1).getValues();
  
  // Open Responses sheets
  var shtRspn = ss.getSheetByName('Responses');
  var shtRspnEN = ss.getSheetByName('Responses EN');
  var shtRspnFR = ss.getSheetByName('Responses FR');

  var RspnMaxRowsEN = shtRspnEN.getMaxRows();
  var RspnMaxRowsFR = shtRspnFR.getMaxRows();
  
  // Function Variables
  var ResponseData;
  var DataCopiedStatus = 0;
  var TimeStamp;
  var Password;
  var LeaguePassword = shtConfig.getRange(8,9).getValue();
  var PasswordValid = 0;
  var RspnRow;
  var WeekNum;
  
  var EmailAddresses = subCreateArray(3,2);
  
  // Data Processing Flags
  var Status = new Array(2); // Status[0] = Status Value, Status[1] = Status Message
  
  // Function Polled Values
  var RspnNextRow = shtRspn.getRange(1, ColNextEmptyRow).getValue();
  var EntriesProcessing;
  var RspnNextRowEN = shtRspnEN.getRange(1, ColNextEmptyRow).getValue();
  var RspnNextRowFR = shtRspnFR.getRange(1, ColNextEmptyRow).getValue();
    
  // Execute if Trigger is Enabled
  if(cfgTrigReport == 'Enabled'){

    Logger.log('------- New Match Report Received -------');
    var MatchProcessLog = 'TCG Match Process Log - ' + shtConfig.getRange(11,2).getValue() + ' ' + shtConfig.getRange(13,2).getValue() + '- Entry Row ' + RspnRow;
    Logger.log(MatchProcessLog);
    
    EntriesProcessing = shtRspn.getRange(1, ColNbUnprcsdEntries).getValue();
    Logger.log('Nb of Entries Before Copying: %s',EntriesProcessing);
    
    // Look for Unprocessed Data in Responses EN
    for (RspnRow = RspnNextRowEN; RspnRow <= RspnMaxRowsEN; RspnRow++){
      
      // Copy the new response data (from Time Stamp to Data Copied Field)
      ResponseData = shtRspnEN.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
      TimeStamp = ResponseData[0][0];
      Password = ResponseData[0][1];
      WeekNum = ResponseData[0][3];
      DataCopiedStatus = ResponseData[0][9];
      
      // Look if Password is valid
      Logger.log('Password Entered: %s', Password);
      if(Password == LeaguePassword) PasswordValid = 1; 
      
      // Check if DataCopied Field is null and Email is Valid, we found new data to copy
      if (DataCopiedStatus == '' && PasswordValid == 1){
        Logger.log('Password Valid, Data Copied to Responses');
        DataCopiedStatus = 'Data Copied';
        shtRspnEN.getRange(RspnRow, ColDataCopied).setValue(DataCopiedStatus);
        // Creates formula to update Last Entry Processed
        shtRspnEN.getRange(RspnRow, ColNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-30]",FALSE)<>"",1,"")');
      }
      // If TimeStamp is null, Delete Row and start over
      if (TimeStamp == '' && RspnRow < RspnMaxRowsEN) {
        shtRspnEN.deleteRow(RspnRow);
        RspnRow = RspnNextRowEN - 1;
      }
      // If Email is not Valid, update Data Copied and Next Empty Row Cells
      if (PasswordValid == 0){
        Logger.log('Password Not Valid');
        DataCopiedStatus = 'Password Not Valid';
        shtRspnEN.getRange(RspnRow, ColDataCopied).setValue(DataCopiedStatus);
        // Creates formula to update Last Entry Processed
        shtRspnEN.getRange(RspnRow, ColNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-30]",FALSE)<>"",1,"")');
      }
      // If Data is copied or Email is not Valid or TimeStamp is null, Exit loop Responses EN to process data
      if (DataCopiedStatus == 'Data Copied' || DataCopiedStatus == 'Password Not Valid' || (TimeStamp == '' && RspnRow >= RspnMaxRowsEN)) {
        RspnRow = RspnMaxRowsEN + 1;
      }
    }
    
    // Executes Responses FR loop only if Responses EN did not find anything
    if (DataCopiedStatus == 0){
      
      // Look for Unprocessed Data in Responses FR
      for (RspnRow = RspnNextRowFR; RspnRow <= RspnMaxRowsFR; RspnRow++){
        
        // Copy the new response data (from Time Stamp to Data Copied Field)
        ResponseData = shtRspnFR.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
        TimeStamp = ResponseData[0][0];
        Password = ResponseData[0][1];
        WeekNum = ResponseData[0][3];
        DataCopiedStatus = ResponseData[0][9];
        
        // Look if Password is valid
        Logger.log('Password Entered: %s', Password);
        if(Password == LeaguePassword) PasswordValid = 1;
        
        // Check if DataCopied Field is null and Email is Valid, we found new data to copy
        if (DataCopiedStatus == '' && PasswordValid == 1){
          Logger.log('Password Valid, Data Copied to Responses');
          DataCopiedStatus = 'Data Copied';
          shtRspnFR.getRange(RspnRow, ColDataCopied).setValue(DataCopiedStatus);
          // Creates formula to update Last Entry Processed
          shtRspnFR.getRange(RspnRow, ColNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-30]",FALSE)<>"",1,"")');
        }
        // If TimeStamp is null, Delete Row and start over
        if (TimeStamp == '' && RspnRow < RspnMaxRowsFR) {
          shtRspnFR.deleteRow(RspnRow);
          RspnRow = RspnNextRowFR - 1;
        }
        // If Email is not Valid, update Data Copied and Next Empty Row Cells
        if (PasswordValid == 0){
          Logger.log('Password Not Valid');
          DataCopiedStatus = 'Password Not Valid';
          shtRspnFR.getRange(RspnRow, ColDataCopied).setValue(DataCopiedStatus);
          // Creates formula to update Last Entry Processed
          shtRspnFR.getRange(RspnRow, ColNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-30]",FALSE)<>"",1,"")');
        }
        // If Data is copied, Exit loop Responses FR to process data
        if (DataCopiedStatus == 'Data Copied' || DataCopiedStatus == 'Password Not Valid' || (TimeStamp == '' && RspnRow >= RspnMaxRowsFR)) {
          RspnRow = RspnMaxRowsFR + 1;
        }
      }
    }
    
    // If Data is copied, put it in Responses Sheet
    if (DataCopiedStatus == 'Data Copied'){
      
      // Copy New Entry Data to Main Responses Sheet
      shtRspn.getRange(RspnNextRow, 1, 1, RspnDataInputs).setValues(ResponseData);
      
      Logger.log('Match Data Copied for Players: %s, %s',ResponseData[0][4],ResponseData[0][5]);
      
      // Copy Formula to detect if an entry is currently processing
      shtRspn.getRange(RspnNextRow, ColNextEmptyRow).setValue('=IF(INDIRECT("R[0]C[-30]",FALSE)<>"",1,"")');
      shtRspn.getRange(RspnNextRow, ColNbUnprcsdEntries).setValue('=IF(AND(INDIRECT("R[0]C[-31]",FALSE)<>"",INDIRECT("R[0]C[-4]",FALSE)<>2),1,"")');
      
      // Troubleshoot
      EntriesProcessing = shtRspn.getRange(1, ColNbUnprcsdEntries).getValue();
      Logger.log('Nb of Entries Pending After Copying Match Data: %s',EntriesProcessing)
      
      // Make sure that we only execute this loop on the first instance call
      if (EntriesProcessing == 1){
        // Execute Game Results Analysis for as long as there are unprocessed entries
        while (EntriesProcessing >= 1) {
          Status = fcnAnalyzeResultsWG(ss, shtConfig, ConfigData, shtRspn);
          EntriesProcessing = shtRspn.getRange(1, ColNbUnprcsdEntries).getValue();
          Logger.log('Nb of Entries Pending After Processing: %s',EntriesProcessing)
        }
      }
      // If the Match was successfully Posted, Update League Standings
      if (Status[0] == 10){
        Logger.log('--------- Updating Standings ---------');
        Logger.log('Update Standings');
        // Execute Ranking function in Standing tab
        fcnUpdateStandings(ss, shtConfig);
        
        Logger.log('Copy to League Spreadsheets');
        // Copy all data to League Spreadsheet
        fcnCopyStandingsSheets(ss, shtConfig, Status[2], 0);
        Logger.log('------------ Standings Updated ------------');
      }
    }
  }
  
  // Send Error Email to sender if Email is not valid
  if(PasswordValid == 0){
    Logger.log('Submission Email Not Valid : %s',Email);
    // Get Emails from both players
    EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, ResponseData[0][4], ResponseData[0][5]);
    fcnMatchReportPwdError(shtConfig, EmailAddresses);
  }
  
  // Post Log to Log Sheet
  subPostLog(shtLog);
  
}


// **********************************************
// function fcnGameResults()
//
// This function populates the Game Results tab 
// once a player submitted his Form
//
// **********************************************

function fcnAnalyzeResultsWG(ss, shtConfig, ConfigData, shtRspn) {
  
  // Data from Configuration File
  // Code Execution Options
  var OptDualSubmission = ConfigData[0][0]; // If Dual Submission is disabled, look for duplicate instead
  var OptPostResult = ConfigData[1][0];
  var OptPlyrMatchValidation = ConfigData[2][0];
  var OptWargame = ConfigData[4][0];
  var OptSendEmail = ConfigData[6][0];
  var cfgNbCards = ConfigData[13][0];
  
  // Columns Values and Parameters
  var ColMatchID = ConfigData[20][0];
  var ColPrcsd = ConfigData[21][0];
  var ColDataConflict = ConfigData[22][0];
  var ColStatus = ConfigData[23][0];
  var ColStatusMsg = ConfigData[24][0];
  var ColMatchIDLastVal = ConfigData[25][0];
  var ColNextEmptyRow = ConfigData[26][0];
  var ColNbUnprcsdEntries = ConfigData[27][0];
  var RspnDataInputs = ConfigData[28][0]; // from Time Stamp to Data Processed

  // League Duration
  var LgParamDuration = shtConfig.getRange(10,9).getValue();
    
  // Test Sheet (for Debug)
  var shtTest = ss.getSheetByName('Test') ;  
  
  // Form Responses Sheet Variables
  var RspnMaxRows = shtRspn.getMaxRows();
  var RspnMaxCols = shtRspn.getMaxColumns();
  var RspnNextRowPrcss = shtRspn.getRange(1, ColNextEmptyRow).getValue() - shtRspn.getRange(1, ColNbUnprcsdEntries).getValue();
  var RspnPlyrSubmit;
  var RspnLocation;
  var RspnWeekNum;
  var RspnDataWeek;
  var RspnDataWinr;
  var RspnDataLosr;
  var RspnDataTie;
  var RspnDataPrcssd = 0;
  var ResponseData;
  var MatchingRspnData;

  // Match Data Variables
  var MatchID; 
  var MatchData = subCreateArray(26,4);
  // 0 = TimeStamp
  // 1 = Location
  // 2 = MatchID
  // 3 = Week #
  // 4 = Winning Player
  // 5 = Losing Player
  // 6 = Game Tie (Yes or No)
  // 7-23 = Not Used
  // 24 = MatchPostStatus
    
   
  
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
  var Status = new Array(2); // Status[0] = Status Value, Status[1] = Status Message
  Status[0] = 0;
  
  var DuplicateRspn = -99;
  var MatchingRspn = -98;
  var MatchPostStatus = -97;
  var CardDBUpdated = -96;
  
  Logger.log('--------- Posting Match ---------'); 
  Logger.log('--------- Options ---------');
  Logger.log('Dual Submission Option: %s',OptDualSubmission);
  Logger.log('Post Results Option: %s',OptPostResult);
  Logger.log('Player Match Validation Option: %s',OptPlyrMatchValidation);
  Logger.log('Wargame Option: %s',OptWargame);
  Logger.log('Send Email Option: %s',OptSendEmail);
  
  // Find a Row that is not processed in the Response Sheet (added data)
  for (var RspnRow = RspnNextRowPrcss; RspnRow <= RspnMaxRows; RspnRow++){
    
    // Copy the new response data (from Time Stamp to Data Processed Field
    ResponseData = shtRspn.getRange(RspnRow, 1, 1, RspnDataInputs).getValues();
    
    // Values from Response Data
    RspnDataPrcssd = ResponseData[0][9];
    RspnPlyrSubmit = ResponseData[0][1]; // Player Submitting
    RspnLocation   = ResponseData[0][2]; // Match Location (Store Yes or No)
    RspnWeekNum    = ResponseData[0][3]; // Week / Round Number
    RspnDataWinr   = ResponseData[0][4]; // Winning Player
    RspnDataLosr   = ResponseData[0][5]; // Losing Player
    RspnDataTie    = ResponseData[0][6]; // Tie
    
    Logger.log('Players: %s, %s',ResponseData[0][4],ResponseData[0][5]);
    
    // If week number is not empty and Processed is empty, Response Data needs to be processed
    if (RspnWeekNum != '' && RspnDataPrcssd == ''){
      
      // If both Players in the response are different, continue
      if (RspnDataWinr != RspnDataLosr){
        
        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 1;
          Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
        }
                
        // Generates the Match ID in advance if data analysis is successful
        MatchID = shtRspn.getRange(1, ColMatchIDLastVal).getValue() + 1;
        
        Logger.log('New Data Found at Row: %s',RspnRow);

        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 2;
          Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
        }
        
        // Look for Duplicate Entry (looks in all entries with MatchID and combination of Week Number, Winner and Loser) 
        // Real code will look at Player Posting Data as well
        DuplicateRspn = fcnFindDuplicateData(ss, ConfigData, shtRspn, ResponseData, RspnRow, RspnMaxRows, shtTest);  
        
        if(DuplicateRspn == 0) Logger.log('No Duplicate Found');
        if(DuplicateRspn > 0 ) Logger.log('Duplicate Found at Row: %s', DuplicateRspn);
        
        // FindDuplicateEntry function was executed properly and didn't find any Duplicate entry, continue analyzing the response data
        if (DuplicateRspn == 0){
          
          // If Dual Submission is enabled, Search if the other Entry matching this response has been submitted (must be enabled)
          if (OptDualSubmission == 'Enabled'){
            
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 3; 
              Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
            }
            // function returns row where the matching data was found
            MatchingRspn = fcnFindMatchingData(ss, ConfigData, shtRspn, ResponseData, RspnRow, RspnMaxRows, shtTest);
            if (MatchingRspn < 0) DuplicateRspn = 0 - MatchingRspn;
          }
          
          // Search if the other Entry matching this response has been submitted
          if (OptDualSubmission == 'Disabled'){
            MatchingRspn = RspnRow;
          }      
          
          Logger.log('Matching Result: %s', MatchingRspn);
          
          // If the result of the fcnFindMatchingEntry function returns something greater than 0, we found a matching entry, continue analyzing the response data
          if (MatchingRspn > 0){
            
            if (OptPostResult == 'Enabled'){
              
              // Get the Entry Data found at row MatchingRspn
              MatchingRspnData = shtRspn.getRange(MatchingRspn, 1, 1, RspnDataInputs).getValues();
              
              // Updates the Status while processing
              if(Status[0] >= 0){
                Status[0] = 4; 
                Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
              }
              // Execute function to populate Match Result Sheet from processed data
              MatchData = fcnPostMatchResultsWG(ss, ConfigData, shtRspn, ResponseData, MatchingRspnData, MatchID, MatchData, shtTest);
              MatchPostStatus = MatchData[25][0];
              
              Logger.log('Match Post Status: %s',MatchPostStatus);
              
              // If Match was populated in Match Results Tab
              if (MatchPostStatus == 1){
                // Match ID doesn't change because we assumed it was already OK
                Logger.log('Match Posted ID: %s',MatchID);
                
                // Reserved for TCG
                // Reserved for TCG
                // Reserved for TCG
                // Reserved for TCG
                
                // If Game Type is Wargame, Update available amount of Power Level Available
                if(OptWargame == 'Enabled'){
                  // Updates the Status while processing
                  if(Status[0] >= 0){
                    Status[0] = 5; 
                    Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
                  }
                  // Update Player Army DB and Army List
                  fcnUpdateArmyDB(shtConfig, RspnDataLosr, MatchData[5][2], shtTest); // MatchData[5][2] = Loser Power Level Bonus
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
            if (OptPostResult == 'Disabled'){
              // Match ID doesn't change because we assumed it was already OK
              
            }
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 6; 
              Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
            }
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          }
          
          // If MatchingEntry = 0, fcnFindMatchingEntry did not find a matching entry, it might be the first response entry
          if (OptDualSubmission == 'Enabled' && MatchingRspn == 0){
            // Updates the Status while processing
            if(Status[0] >= 0){
              Status[0] = 0;
              Status[1] = 'Waiting for Other Response Submission';
            }
            // Set the Data Processed Flag
            RspnDataPrcssd = 1;
          } 
          
          // If MatchingEntry = -1, fcnFindMatchingEntry was not executed properly, send email to notify
          if (OptDualSubmission == 'Enabled' && MatchingRspn == -1){
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
      if (RspnDataWinr == RspnDataLosr){
        
        // Updates the Match ID to an empty value 
        MatchID = '';
        
        // Set the Data Processed Flag
        RspnDataPrcssd = 1;  
        
        // Set the Status Message
        Status = subGenErrorMsg(Status, -50,0);
      }
      
      Logger.log('Match Post Status: %s - %s',Status[0], Status[1])
      
      // Call the Email Function, sends Match Data if Send Email Option is Enabled
      if(Status[0] >= 0 && OptSendEmail == 'Enabled') {
        
        // Updates the Status while processing
        if(Status[0] >= 0){
          Status[0] = 7; 
          Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
        }
        // Get Email addresses from Config File
        EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, RspnDataWinr, RspnDataLosr);
        
        // Send email to players. Each function analyzes language preferences
        fcnSendConfirmEmailEN(shtConfig, EmailAddresses, MatchData);
        fcnSendConfirmEmailFR(shtConfig, EmailAddresses, MatchData);
        Logger.log('Confirmation Emails Sent');
      }
      
      // If an Error has been detected that prevented to process the Match Data, send available data and Error Message
      if(Status[0] < 0 && OptSendEmail == 'Enabled') {
      
        // Populates Match Data
        MatchData[0][0] = ResponseData[0][0]; // TimeStamp
        MatchData[0][0] = Utilities.formatDate (MatchData[0][0], Session.getScriptTimeZone(), 'YYYY-MM-dd HH:mm:ss');
        
        MatchData[1][0] = ResponseData[0][2];  // Location (Store Y/N)
        MatchData[2][0] = MatchID;             // MatchID
        MatchData[3][0] = ResponseData[0][3];  // Week/Round Number
        MatchData[4][0] = ResponseData[0][4];  // Winning Player
        MatchData[5][0] = ResponseData[0][5];  // Losing Player
        MatchData[6][0] = ResponseData[0][6];  // Game is a Tie
        
        // Get Email addresses from Config File
        EmailAddresses = subGetEmailAddressDbl(ss, EmailAddresses, RspnDataWinr, RspnDataLosr);
        
        // Send Error Message, each function analyzes language preferences
        fcnSendErrorEmailEN(shtConfig, EmailAddresses, MatchData, MatchID, Status);
        fcnSendErrorEmailFR(shtConfig, EmailAddresses, MatchData, MatchID, Status);
        Logger.log('Error Emails Sent');
      }
      
      // Updates the Status while processing
      if(Status[0] >= 0){
        Status[0] = 9; 
        Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
      }
      // Set the Match ID (for both Response and Matching Entry), and Updates the Last Match ID generated, 
      if (MatchPostStatus == 1 || OptPostResult == 'Disabled'){
        shtRspn.getRange(RspnRow, ColMatchID).setValue(MatchID);
        shtRspn.getRange(1, ColMatchIDLastVal).setValue(MatchID);
      }
      
      // Updates the Status while processing
      if(Status[0] >= 0){
        Status[0] = 10; 
        Status[1] = subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, Status[0]);
        Status[2] = MatchData[3][0]; // Week Processed
      }
      // Updating Match Process Data
      shtRspn.getRange(RspnRow, ColPrcsd).setValue(RspnDataPrcssd);
      shtRspn.getRange(RspnRow, ColNbUnprcsdEntries).setValue(0);
      
      // If Process Error has been detected, update the Response Process Data
      if(Status[0] < 0){
        shtRspn.getRange(RspnRow, ColStatus).setValue(Status[0]);
        shtRspn.getRange(RspnRow, ColStatusMsg).setValue(Status[1]);
      }
      
      // Set the Matching Response Match ID if Matching Response found
      if (MatchingRspn > 0) shtRspn.getRange(MatchingRspn, ColMatchID).setValue(MatchID);	  
            
    }
    // When Week Number is empty or if the Response Data was processed, we have reached the end of the list, then exit the loop
    if(RspnWeekNum == '' || RspnDataPrcssd == 1) {
      Logger.log('Response Loop exit at Row: %s',RspnRow)
      RspnRow = RspnMaxRows + 1;
    }
  }
  
  Logger.log('------------ Game Posted ------------');
    
  return Status;
}


