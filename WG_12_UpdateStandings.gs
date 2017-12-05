// **********************************************
// function fcnUpdateStandings()
//
// Updates the Standings according to the Win % 
// from the Cumulative Results tab to the Standings Tab
//
// **********************************************

function fcnUpdateStandings(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData){

//  function fcnUpdateStandings(){
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  
//  // Config Sheet to get options
//  var shtConfig = ss.getSheetByName('Config');
//  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
//  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
//  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
//  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
//  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  
  Logger.log("Routine: fcnUpdateStandings");
  
  var shtConfig = ss.getSheetByName('Config');
  // Event Parameters
  var evntRanking = cfgEvntParam[17][0];
  var evntRankMatchLimit = cfgEvntParam[18][0];
  var NbPlayers = shtConfig.getRange('B13').getValue();
    
  // Column Values
  var colPlyr =     cfgColRndSht[0][0]+1;
  var colTeam =     cfgColRndSht[1][0]+1;
  var colMatch =    cfgColRndSht[2][0]+1;
  var colWins =     cfgColRndSht[3][0]+1;
  var colLoss =     cfgColRndSht[4][0]+1;
  var colTie =      cfgColRndSht[5][0]+1;
  var colPts =      cfgColRndSht[6][0]+1;
  var colWinPerc =  cfgColRndSht[7][0]+1;
  var colLocation = cfgColRndSht[8][0]+1;
    
//  var shtTest = ss.getSheetByName('Test');
//  shtTest.getRange(10,2,16,1).setValues(cfgColRndSht);
  
  // Sheets
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var shtStand = ss.getSheetByName('Standings');
  
  var NbColValues = 8;

  // Get Cumulative Results Values
  var ValCumul = shtCumul.getRange(5,1,32,NbColValues).getValues(); // Rows = Players, Columns 0= Player Name, 1= N/A, 2= MP, 3= W, 4= L, 5= Tie, 6= Pts, 7=W%
  
  // Standings Ranges In Limits and Out Limits
  var RngStandInLim;
  var RngStandOutLim;
   
  var InLimit = 0;
  var OutLimit = 0;
  var PlyrInLimArray = subCreateArray(NbPlayers,NbColValues);
  var PlyrOutLimArray = subCreateArray(NbPlayers,NbColValues);
  
  // Find Players with enough matches played
  for(var i=0; i<NbPlayers; i++){
    // If player has played enough matches, put it in InLimit Array
    if(ValCumul[i][2] >= evntRankMatchLimit){
      PlyrInLimArray[InLimit] = ValCumul[i];
      //Logger.log('In Limit - Player: %s - MP: %s',PlyrInLimArray[InLimit][0], PlyrInLimArray[InLimit][2]);
      InLimit++;
    }
    // If player has not played enough matches, put it in OutLimit Array
    if(ValCumul[i][2] < evntRankMatchLimit){
      PlyrOutLimArray[OutLimit] = ValCumul[i];
      //Logger.log('Out Limit - Player: %s - MP: %s',PlyrOutLimArray[OutLimit][0], PlyrOutLimArray[OutLimit][2]);
      OutLimit++;
    }
  }
  // Define new lengths for both arrays
  PlyrInLimArray.length  = InLimit;
  PlyrOutLimArray.length = OutLimit;
  
  
  // Create New Ranges with those Arrays
  // In Limit Array
  if(InLimit > 0){
    RngStandInLim = shtStand.getRange(6, 2, InLimit, NbColValues);
    RngStandInLim.setValues(PlyrInLimArray);
//    shtTest.getRange(10,3).setValue(PlyrInLimArray.length);
  }
  // Out Limit Array
  if(OutLimit > 0){
    RngStandOutLim = shtStand.getRange(6+InLimit, 2, OutLimit, NbColValues);
    RngStandOutLim.setValues(PlyrOutLimArray);
//    shtTest.getRange(11,3).setValue(PlyrOutLimArray.length);
  }
  
  // Points - Sorts the Standings Values by Points and Matches Played
  if(evntRanking == 'Points'){
    // Sort In Limit Range
    if(InLimit > 0)  RngStandInLim.sort([{column: colPts, ascending: false},{column: colWinPerc, ascending: false}]);
    // Sort Out Limit Range
    if(OutLimit > 0) RngStandOutLim.sort([{column: colPts, ascending: false},{column: colWinPerc, ascending: false}]);
  }
  // Wins - Sorts the Standings Values by Wins and Win Percentage
  if(evntRanking == 'Wins'){
    // Sort In Limit Range
    if(InLimit > 0)  RngStandInLim.sort([{column: colWins, ascending: false},{column: colWinPerc, ascending: false}]);
    // Sort Out Limit Range
    if(OutLimit > 0) RngStandOutLim.sort([{column: colWins, ascending: false},{column: colWinPerc, ascending: false}]);
  }
  // Win % - Sorts the Standings Values by Win Percentage and Matches Played
  if(evntRanking == 'Win%'){
    // Sort In Limit Range
    if(InLimit > 0)  RngStandInLim.sort([{column: colWinPerc, ascending: false},{column: colMatch, ascending: false}]);
    // Sort Out Limit Range
    if(OutLimit > 0) RngStandOutLim.sort([{column: colWinPerc, ascending: false},{column: colMatch, ascending: false}]);
  }
}

// **********************************************
// function fcnCopyStandingsSheets()
//
// This function copies all Standings and Results in 
// the spreadsheet that is accessible to players
//
// **********************************************

function fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, cfgColRndSht, RspnRoundNum, AllSheets){
//function fcnCopyStandingsSheets(){
//  
//  var ss = SpreadsheetApp.getActiveSpreadsheet();
//  
//  // Config Sheet to get options
//  var shtConfig = ss.getSheetByName('Config');
//  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
//  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
//  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
//  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
//  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
//  
//  var RspnRoundNum = 0;
//  var AllSheets = 1;
  
  Logger.log("Routine: fcnCopyStandingsSheets");

  var shtIDs = shtConfig.getRange(4, 7,20,1).getValues();
  var shtUrl = shtConfig.getRange(4,11,20,1).getValues();
  
  // Open Player Standings Spreadsheet
  var ssStdngEN = SpreadsheetApp.openById(shtIDs[5][0]);
  var ssStdngFR = SpreadsheetApp.openById(shtIDs[6][0]);
  
  // Match Report Form URL
  var FormUrlEN = shtUrl[7][0];
  var FormUrlFR = shtUrl[8][0];
  
  // Event Parameters
  var evntNameEN = cfgEvntParam[0][0] + ' ' + cfgEvntParam[7][0];
  var evntNameFR = cfgEvntParam[8][0] + ' ' + cfgEvntParam[0][0];
  var evntRoundLimit = cfgEvntParam[13][0];
  var RoundSheet = RspnRoundNum + 1; // Round 1 sheet is sheet[2]
  var RoundNum;
  var NbPlayers = shtConfig.getRange('B13').getValue();
  
  // Sheet Initialization
  var rngSheetInitEN = shtConfig.getRange(9,9);
  var SheetInitEN = rngSheetInitEN.getValue();
  var rngSheetInitFR = shtConfig.getRange(10,9);
  var SheetInitFR = rngSheetInitFR.getValue();
  
  // Player Status and Warning Columns
  var colPlyrStatus =  cfgColRndSht[14][0];
  var colPlyrWarning = cfgColRndSht[15][0];
  
  // Header Ranges Values
  var rngValStandings = 'B2:B3';
  var rngValCumulRslt = 'A2:A4';
  var rngValRound =     'A2:A4';
  
  // Function Variables
  var ssMstrSht;
  var ssMstrShtStartRow;
  var ssMstrShtMaxRows;
  var ssMstrShtNbCol;
  var ssMstrShtData;
  var ssMstrStartDate;
  var ssMstrEndDate;
  var NumValues;
  var ColValues;
  var SheetName;
  
  var ssLgShtEN;
  var ssLgShtFR;
  var RoundGame;
  
  // Loops through tabs 0-9 (0= Standings, 1= Cumulative Results, 2-9= Round 1-8)
  for (var sht = 0; sht <= 9; sht++){
    ssMstrSht = ss.getSheets()[sht];
    SheetName = ssMstrSht.getSheetName();
    
    if(sht == 0 || sht == 1 || sht == RoundSheet || AllSheets == 1){
      ssMstrShtMaxRows = ssMstrSht.getMaxRows();
      
      // Get Sheets
      ssLgShtEN = ssStdngEN.getSheets()[sht];
      ssLgShtFR = ssStdngFR.getSheets()[sht];
      
      // If sheet is Standings
      if (sht == 0) {
        ssMstrShtStartRow = 6;
        ssMstrShtNbCol = ssMstrSht.getMaxColumns();
      }
      
      // If sheet is Cumulative Results
      if (sht >= 1) {
        ssMstrShtStartRow = 5;
        ssMstrShtNbCol = ssMstrSht.getMaxColumns();
      }
      
      // Update Header
      // Standings Sheet 
      if(sht == 0){
        ssMstrShtData = ssMstrSht.getRange(rngValStandings).getValues();
        ssLgShtEN.getRange(rngValStandings).setValues(ssMstrShtData);
        ssLgShtFR.getRange(rngValStandings).setValues(ssMstrShtData);
      }      
      
      // Cumulative Results Sheet 
      if(sht == 1){
        ssMstrShtData = ssMstrSht.getRange(rngValCumulRslt).getValues();
        ssLgShtEN.getRange(rngValCumulRslt).setValues(ssMstrShtData);
        ssLgShtFR.getRange(rngValCumulRslt).setValues(ssMstrShtData);
      }
            
      // Set the number of values to fetch
      NumValues = ssMstrShtMaxRows - ssMstrShtStartRow + 1;
      
      // Get Range and Data from Master
      ssMstrShtData = ssMstrSht.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).getValues();
      
      // And copy to Standings
      ssLgShtEN.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).setValues(ssMstrShtData);
      ssLgShtFR.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).setValues(ssMstrShtData);
      Logger.log("Copied to Sheet: %s",ssLgShtEN.getName());
      
      // Hide Unused Rows
      if(NbPlayers > 0){
        ssLgShtEN.hideRows(ssMstrShtStartRow, ssMstrShtMaxRows - ssMstrShtStartRow + 1);
        ssLgShtEN.showRows(ssMstrShtStartRow, NbPlayers);
        ssLgShtFR.hideRows(ssMstrShtStartRow, ssMstrShtMaxRows - ssMstrShtStartRow + 1);
        ssLgShtFR.showRows(ssMstrShtStartRow, NbPlayers);
      }
    }
    
    // If Sheet Titles are not initialized, initialize them
    if(SheetInitEN != "Init"){
      // Standings Sheet
      if (sht == 0){
        Logger.log('Standings Sheet Updated');
        // Update League Name
        ssLgShtEN.getRange(4,2).setValue(evntNameEN + ' Standings')
        ssLgShtFR.getRange(4,2).setValue('Classement ' + evntNameFR)
        // Update Form Link
        ssLgShtEN.getRange(2,5).setValue('=HYPERLINK("' + FormUrlEN + '","Send Match Results")');      
        ssLgShtFR.getRange(2,5).setValue('=HYPERLINK("' + FormUrlFR + '","Envoyer Résultats de Match")'); 
      }
      
      // Cumulative Results Sheet
      if (sht == 1){
        Logger.log('Cumulative Results Sheet Updated');
        RoundGame = ssMstrSht.getRange(rngValRound).getValues();
        ssLgShtEN.getRange(rngValRound).setValues(RoundGame);
        ssLgShtFR.getRange(rngValRound).setValues(RoundGame);
        
        // Loop through Values in Player Status to translate each value
        ColValues = ssLgShtFR.getRange(ssMstrShtStartRow, colPlyrStatus, NumValues, 1).getValues();
        for (var row = 0 ; row < NumValues; row++){
          if (ColValues[row][0] == 'Active') ColValues[row][0] = 'Actif';
          if (ColValues[row][0] == 'Eliminated') ColValues[row][0] = 'Éliminé';
        }
        ssLgShtFR.getRange(ssMstrShtStartRow, colPlyrStatus, NumValues, 1).setValues(ColValues);
        
        // Loop through Values in Player Warning to translate each value
        ColValues = ssLgShtFR.getRange(ssMstrShtStartRow, colPlyrWarning, NumValues, 1).getValues();
        for (var row = 0 ; row < NumValues; row++){
          if (ColValues[row][0] == 'Yes') ColValues[row][0] = 'Oui';
          if (ColValues[row][0] == 'No')  ColValues[row][0] = 'Non';
        }
        ssLgShtFR.getRange(ssMstrShtStartRow, colPlyrWarning, NumValues, 1).setValues(ColValues);
      }
      
      // Round Sheet 
      if(sht >= 2 && sht <= 9){
        Logger.log('Round %s Sheet Updated',sht-1);
        ssMstrStartDate = ssMstrSht.getRange(3,1).getValue();
        ssMstrEndDate   = ssMstrSht.getRange(4,1).getValue();
        ssLgShtEN.getRange(3,1).setValue('Start: ' + ssMstrStartDate);
        ssLgShtEN.getRange(4,1).setValue('End: ' + ssMstrEndDate);
        ssLgShtFR.getRange(3,1).setValue('Début: ' + ssMstrStartDate);
        ssLgShtFR.getRange(4,1).setValue('Fin: ' + ssMstrEndDate);
      }
      
      // If the current sheet is greater than League Round Limit, hide sheet
      if(sht > evntRoundLimit + 1){
        ssLgShtEN.hideSheet();
        ssLgShtFR.hideSheet();
      }
      // If the current sheet is less than League Round Limit, show sheet
      if(sht <= evntRoundLimit + 1){
        ssLgShtEN.showSheet();
        ssLgShtFR.showSheet();
      }
      
      // Set Initialized Value to Config Sheet to skip this part
      if(sht == 9) {
        rngSheetInitEN.setValue("Init");
        rngSheetInitFR.setValue("Init");
      }
    }
  }
}