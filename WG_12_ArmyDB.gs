// **********************************************
// function fcnUpdateCardDB()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateArmyDB(shtConfig, Player, AvailValue, shtTest){
  
  // Config Spreadsheet
  var ssArmyDBID = shtConfig.getRange(31,2).getValue();
  var ssArmyListEnID = shtConfig.getRange(32,2).getValue();
  var ssArmyListFrID = shtConfig.getRange(33,2).getValue();

  var cfgRatingMode = shtConfig.getRange(6,7).getValue();
  var cfgCurrRoundValue = shtConfig.getRange(10,7).getValue();
  
  // Player Card DB Spreadsheet
  var shtArmyDB = SpreadsheetApp.openById(ssArmyDBID).getSheetByName(Player);
  var rngArmyDBCurrRoundPwrLvl = shtArmyDB.getRange(5,9);
  var rngArmyDBAvailPwrLvl = shtArmyDB.getRange(5,10);
  var rngArmyDBCurrRoundPoints = shtArmyDB.getRange(5,11);
  var rngArmyDBAvailPoints = shtArmyDB.getRange(5,12);

  // Army List Spreadsheet
  var shtArmyListEn = SpreadsheetApp.openById(ssArmyListEnID).getSheetByName(Player);
  var rngArmyListEnCurrRoundPwrLvl = shtArmyListEn.getRange(5,9);
  var rngArmyListEnAvailPwrLvl = shtArmyListEn.getRange(5,10);
  var rngArmyListEnCurrRoundPoints = shtArmyListEn.getRange(5,11);
  var rngArmyListEnAvailPoints = shtArmyListEn.getRange(5,12);
  
  var shtArmyListFr = SpreadsheetApp.openById(ssArmyListFrID).getSheetByName(Player);
  var rngArmyListFrCurrRoundPwrLvl = shtArmyListFr.getRange(5,9);
  var rngArmyListFrAvailPwrLvl = shtArmyListFr.getRange(5,10);
  var rngArmyListFrCurrRoundPoints = shtArmyListFr.getRange(5,11);
  var rngArmyListFrAvailPoints = shtArmyListFr.getRange(5,12);
  
  // Get Cells to Update according to the Army Rating Mode (Power Level or Points)
  if(cfgRatingMode == 'Power Level'){
    // Update the Army DB
    rngArmyDBCurrRoundPwrLvl.setValue(cfgCurrRoundValue);
    rngArmyDBAvailPwrLvl.setValue(AvailValue);
    
    // Update the Losing Player Army List
    // English File
    rngArmyListEnCurrRoundPwrLvl.setValue(cfgCurrRoundValue);
    rngArmyListEnAvailPwrLvl.setValue(AvailValue);
    // French File
    rngArmyListFrCurrRoundPwrLvl.setValue(cfgCurrRoundValue);
    rngArmyListFrAvailPwrLvl.setValue(AvailValue);
    
    // Hide Points Columns (6-7-8, 11-12)
    shtArmyListEn.hideColumns(6, 3);
    shtArmyListEn.hideColumns(11, 2);    
    shtArmyListFr.hideColumns(6, 3);
    shtArmyListFr.hideColumns(11, 2);
  }

  if(cfgRatingMode == 'Points'){
    // Update the Army DB
    rngArmyDBCurrRoundPoints.setValue(cfgCurrRoundValue);
    rngArmyDBAvailPoints.setValue(AvailValue);
    
    // Update the Losing Player Army List
    // English File
    rngArmyListEnCurrRoundPoints.setValue(cfgCurrRoundValue);
    rngArmyListEnAvailPoints.setValue(AvailValue);
    // French File
    rngArmyListFrCurrRoundPoints.setValue(cfgCurrRoundValue);
    rngArmyListFrAvailPoints.setValue(AvailValue);
  
    // Hide Power Level Columns (5, 9-10)
    shtArmyListEn.hideColumn(5);
    shtArmyListEn.hideColumns(9, 2);    
    shtArmyListFr.hideColumn(5);
    shtArmyListFr.hideColumns(9, 2);
  }
}

