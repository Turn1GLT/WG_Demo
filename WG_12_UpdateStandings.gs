// **********************************************
// function fcnUpdateStandings()
//
// Updates the Standings according to the Win % 
// from the Cumulative Results tab to the Standings Tab
//
// **********************************************

function fcnUpdateStandings(ss, cfgEvntParam, cfgColRspSht, cfgColRndSht, cfgExecData){
  
  Logger.log("Routine: fcnPostResultRoundWG");
  
  // League / Tournament Parameters
  var evntRanking = cfgEvntParam[17][0];
  var evntRankMatchLimit = cfgEvntParam[18][0];
  var evntNbPlayers = cfgEvntParam[31][0];
    
  // Column Values
  var colPlyr = cfgColRndSht[0][0];
  var colTeam = cfgColRndSht[1][0];
  var colMatchPlayed = cfgColRndSht[2][0];
  var colWins = cfgColRndSht[3][0];
  var colLoss = cfgColRndSht[4][0];
  var colPts = cfgColRndSht[6][0];
  var colWinPerc = cfgColRndSht[7][0];
  var colLocation = cfgColRndSht[8][0];
  
  // Sheets
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var shtStand = ss.getSheetByName('Standings');
    
  // Get Cumulative Results Values
  var ValCumul = shtCumul.getRange(5,2,32,6).getValues(); // Rows = Players, Columns 0= Player Name, 1= N/A, 2= MP, 3= W, 4= L, 5= W%
  
  // Standings Ranges In Limits and Out Limits
  var RngStandInLim;
  var RngStandOutLim;
   
  var InLimit = 0;
  var OutLimit = 0;
  var PlyrInLimArray = subCreateArray(EvntNbPlayers,6);
  var PlyrOutLimArray = subCreateArray(EvntNbPlayers,6);
  
  // Find Players with enough matches played
  for(var i=0; i<evntNbPlayers; i++){
    // If player has played enough matches, put it in InLimit Array
    if(ValCumul[i][2] >= evntRankMatchLimit){
      PlyrInLimArray[InLimit] = ValCumul[i];
      Logger.log('In Limit - Player: %s - MP: %s',PlyrInLimArray[InLimit][0], PlyrInLimArray[InLimit][2]);
      InLimit++;
    }
    // If player has not played enough matches, put it in OutLimit Array
    if(ValCumul[i][2] < evntRankMatchLimit){
      PlyrOutLimArray[OutLimit] = ValCumul[i];
      Logger.log('Out Limit - Player: %s - MP: %s',PlyrOutLimArray[OutLimit][0], PlyrOutLimArray[OutLimit][2]);
      OutLimit++;
    }
  }
  // Define new lengths for both arrays
  PlyrInLimArray.length  = InLimit;
  PlyrOutLimArray.length = OutLimit;
  
  // Create New Ranges with those Arrays
  // In Limit Array
  if(InLimit > 0){
    RngStandInLim = shtStand.getRange(6, 2, InLimit, 6);
    RngStandInLim.setValues(PlyrInLimArray);
  }
  // Out Limit Array
  if(OutLimit > 0){
    RngStandOutLim = shtStand.getRange(6+InLimit, 2, OutLimit, 6);
    RngStandOutLim.setValues(PlyrOutLimArray);
  }
  
  // Points - Sorts the Standings Values by Points and Matches Played
  if(EvntRanking == 'Points'){
    // Sort In Limit Range
    RngStandInLim.sort([{column: colPts, ascending: false},{column: colMatchPlayed, ascending: false}]);
    // Sort Out Limit Range
    RngStandOutLim.sort([{column: colPts, ascending: false},{column: colMatchPlayed, ascending: false}]);
  }
  // Wins - Sorts the Standings Values by Wins and Win Percentage
  if(EvntRanking == 'Wins'){
    // Sort In Limit Range
    RngStandInLim.sort([{column: colWins, ascending: false},{column: colWinPerc, ascending: false}]);
    // Sort Out Limit Range
    RngStandOutLim.sort([{column: colWins, ascending: false},{column: colWinPerc, ascending: false}]);
  }
  // Win % - Sorts the Standings Values by Win Percentage and Matches Played
  if(EvntRanking == 'Win%'){
    // Sort In Limit Range
    RngStandInLim.sort([{column: colWinPerc, ascending: false},{column: colMatchPlayed, ascending: false}]);
    // Sort Out Limit Range
    RngStandOutLim.sort([{column: colWinPerc, ascending: false},{column: colMatchPlayed, ascending: false}]);
  }
}

// **********************************************
// function fcnCopyStandingsSheets()
//
// This function copies all Standings and Results in 
// the spreadsheet that is accessible to players
//
// **********************************************

function fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, RspnRoundNum, AllSheets){
  
  Logger.log("Routine: fcnCopyStandingsSheets");

  var shtIDs = shtConfig.getRange(4, 7,20,1).getValues();
  var shtUrl = shtConfig.getRange(4,11,20,1).getValues();
  
  // Open Player Standings Spreadsheet
  var ssStdngEN = SpreadsheetApp.openById(shtIDs[5][0]);
  var ssStdngFR = SpreadsheetApp.openById(shtIDs[6][0]);
  
  // Match Report Form URL
  var FormUrlEN = shtUrl[7][0];
  var FormUrlFR = shtUrl[8][0];
  
  // League Name
  var evntNameEN = cfgEvntParam[0][0] + ' ' + cfgEvntParam[7][0];
  var evntNameFR = cfgEvntParam[8][0] + ' ' + cfgEvntParam[0][0];
  var evntNbPlayers = cfgEvntParam[31][0];
  var evntRoundLimit = cfgEvntParam[13][0];
  var RoundSheet = RspnRoundNum + 1;
  
  // Sheet Initialization
  var rngSheetInitEN = shtConfig.getRange(9,9);
  var SheetInitEN = rngSheetInitEN.getValue();
  var rngSheetInitFR = shtConfig.getRange(10,9);
  var SheetInitFR = rngSheetInitFR.getValue();
  
  // Player Status and Warning Columns
  var colPlyrStatus = 13;
  var colPlyrWarning = 15;
  
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
  
  var ssLgShtEn;
  var ssLgShtFr;
  var RoundGame;
  
  // Loops through tabs 0-9 (Standings, Cumulative Results, Round 1-8)
  for (var sht = 0; sht <= 9; sht++){
    ssMstrSht = ss.getSheets()[sht];
    SheetName = ssMstrSht.getSheetName();
    
    if(sht == 0 || sht == 1 || sht == RoundSheet || AllSheets == 1){
      ssMstrShtMaxRows = ssMstrSht.getMaxRows();
      
      // Get Sheets
      ssLgShtEn = ssStdngEN.getSheets()[sht];
      ssLgShtFr = ssStdngFR.getSheets()[sht];
      
      // If sheet is Standings
      if (sht == 0) {
        ssMstrShtStartRow = 6;
        ssMstrShtNbCol = 7;
      }
      
      // If sheet is Cumulative Results or Round Results
      if (sht == 1) {
        ssMstrShtStartRow = 5;
        ssMstrShtNbCol = 13;
      }
            
      // If sheet is Cumulative Results or Round Results
      if (sht > 1 && sht <= 9) {
        ssMstrShtStartRow = 5;
        ssMstrShtNbCol = 11;
      }
      
      // Set the number of values to fetch
      NumValues = ssMstrShtMaxRows - ssMstrShtStartRow + 1;
      
      // Get Range and Data from Master
      ssMstrShtData = ssMstrSht.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).getValues();
      
      // And copy to Standings
      ssLgShtEn.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).setValues(ssMstrShtData);
      ssLgShtFr.getRange(ssMstrShtStartRow,1,NumValues,ssMstrShtNbCol).setValues(ssMstrShtData);
      
      // Hide Unused Rows
      if(EvntNbPlayers > 0){
        ssLgShtEn.hideRows(ssMstrShtStartRow, ssMstrShtMaxRows - ssMstrShtStartRow + 1);
        ssLgShtEn.showRows(ssMstrShtStartRow, EvntNbPlayers);
        ssLgShtFr.hideRows(ssMstrShtStartRow, ssMstrShtMaxRows - ssMstrShtStartRow + 1);
        ssLgShtFr.showRows(ssMstrShtStartRow, EvntNbPlayers);
      }
       
      // Round Sheet 
      if (sht == RoundSheet){
        Logger.log('Round %s Sheet Updated',sht-1);
        ssMstrStartDate = ssMstrSht.getRange(3,2).getValue();
        ssMstrEndDate   = ssMstrSht.getRange(4,2).getValue();
        ssLgShtEn.getRange(3,2).setValue('Start: ' + ssMstrStartDate);
        ssLgShtEn.getRange(4,2).setValue('End: ' + ssMstrEndDate);
        ssLgShtFr.getRange(3,2).setValue('Début: ' + ssMstrStartDate);
        ssLgShtFr.getRange(4,2).setValue('Fin: ' + ssMstrEndDate);
      }
      
      // If the current sheet is greater than League Round Limit, hide sheet
      if(sht > EvntRoundLimit + 1){
        ssLgShtEn.hideSheet();
        ssLgShtFr.hideSheet();
      }
    }
    
    // If Sheet Titles are not initialized, initialize them
    if(SheetInitEN != "Initialized"){
      // Standings Sheet
      if (sht == 0){
        Logger.log('Standings Sheet Updated');
        // Update League Name
        ssLgShtEn.getRange(4,2).setValue(EvntNameEN + ' Standings')
        ssLgShtFr.getRange(4,2).setValue('Classement ' + EvntNameFR)
        // Update Form Link
        ssLgShtEn.getRange(2,5).setValue('=HYPERLINK("' + FormUrlEN + '","Send Match Results")');      
        ssLgShtFr.getRange(2,5).setValue('=HYPERLINK("' + FormUrlFR + '","Envoyer Résultats de Match")'); 
      }
      
      // Cumulative Results Sheet
      if (sht == 1){
        Logger.log('Cumulative Results Sheet Updated');
        RoundGame = ssMstrSht.getRange(2,3,3,1).getValues();
        ssLgShtEn.getRange(2,3,3,1).setValues(RoundGame);
        ssLgShtFr.getRange(2,3,3,1).setValues(RoundGame);
        
        // Loop through Values in Player Status to translate each value
        ColValues = ssLgShtFr.getRange(ssMstrShtStartRow, colPlyrStatus, NumValues, 1).getValues();
        for (var row = 0 ; row < NumValues; row++){
          if (ColValues[row][0] == 'Active') ColValues[row][0] = 'Actif';
          if (ColValues[row][0] == 'Eliminated') ColValues[row][0] = 'Éliminé';
        }
        ssLgShtFr.getRange(ssMstrShtStartRow, colPlyrStatus, NumValues, 1).setValues(ColValues);
        
        // Loop through Values in Player Warning to translate each value
        ColValues = ssLgShtFr.getRange(ssMstrShtStartRow, colPlyrWarning, NumValues, 1).getValues();
        for (var row = 0 ; row < NumValues; row++){
          if (ColValues[row][0] == 'Yes') ColValues[row][0] = 'Oui';
          if (ColValues[row][0] == 'No')  ColValues[row][0] = 'Non';
        }
        ssLgShtFr.getRange(ssMstrShtStartRow, colPlyrWarning, NumValues, 1).setValues(ColValues);
      }
      
     // Set Initialized Value to Config Sheet to skip this part
      if(sht == 9) {
        rngSheetInitEN.setValue("Initialized");
        rngSheetInitFR.setValue("Initialized");
      }
    }
  }
}