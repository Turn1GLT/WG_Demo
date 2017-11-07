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
  var cfgCurrWeekValue = shtConfig.getRange(10,7).getValue();
  
  // Player Card DB Spreadsheet
  var shtArmyDB = SpreadsheetApp.openById(ssArmyDBID).getSheetByName(Player);
  var rngArmyDBCurrWeekPwrLvl = shtArmyDB.getRange(5,9);
  var rngArmyDBAvailPwrLvl = shtArmyDB.getRange(5,10);
  var rngArmyDBCurrWeekPoints = shtArmyDB.getRange(5,11);
  var rngArmyDBAvailPoints = shtArmyDB.getRange(5,12);

  // Army List Spreadsheet
  var shtArmyListEn = SpreadsheetApp.openById(ssArmyListEnID).getSheetByName(Player);
  var rngArmyListEnCurrWeekPwrLvl = shtArmyListEn.getRange(5,9);
  var rngArmyListEnAvailPwrLvl = shtArmyListEn.getRange(5,10);
  var rngArmyListEnCurrWeekPoints = shtArmyListEn.getRange(5,11);
  var rngArmyListEnAvailPoints = shtArmyListEn.getRange(5,12);
  
  var shtArmyListFr = SpreadsheetApp.openById(ssArmyListFrID).getSheetByName(Player);
  var rngArmyListFrCurrWeekPwrLvl = shtArmyListFr.getRange(5,9);
  var rngArmyListFrAvailPwrLvl = shtArmyListFr.getRange(5,10);
  var rngArmyListFrCurrWeekPoints = shtArmyListFr.getRange(5,11);
  var rngArmyListFrAvailPoints = shtArmyListFr.getRange(5,12);
  
  // Get Cells to Update according to the Army Rating Mode (Power Level or Points)
  if(cfgRatingMode == 'Power Level'){
    // Update the Army DB
    rngArmyDBCurrWeekPwrLvl.setValue(cfgCurrWeekValue);
    rngArmyDBAvailPwrLvl.setValue(AvailValue);
    
    // Update the Losing Player Army List
    // English File
    rngArmyListEnCurrWeekPwrLvl.setValue(cfgCurrWeekValue);
    rngArmyListEnAvailPwrLvl.setValue(AvailValue);
    // French File
    rngArmyListFrCurrWeekPwrLvl.setValue(cfgCurrWeekValue);
    rngArmyListFrAvailPwrLvl.setValue(AvailValue);
    
    // Hide Points Columns (6-7-8, 11-12)
    shtArmyListEn.hideColumns(6, 3);
    shtArmyListEn.hideColumns(11, 2);    
    shtArmyListFr.hideColumns(6, 3);
    shtArmyListFr.hideColumns(11, 2);
  }

  if(cfgRatingMode == 'Points'){
    // Update the Army DB
    rngArmyDBCurrWeekPoints.setValue(cfgCurrWeekValue);
    rngArmyDBAvailPoints.setValue(AvailValue);
    
    // Update the Losing Player Army List
    // English File
    rngArmyListEnCurrWeekPoints.setValue(cfgCurrWeekValue);
    rngArmyListEnAvailPoints.setValue(AvailValue);
    // French File
    rngArmyListFrCurrWeekPoints.setValue(cfgCurrWeekValue);
    rngArmyListFrAvailPoints.setValue(AvailValue);
  
    // Hide Power Level Columns (5, 9-10)
    shtArmyListEn.hideColumn(5);
    shtArmyListEn.hideColumns(9, 2);    
    shtArmyListFr.hideColumn(5);
    shtArmyListFr.hideColumns(9, 2);
  }
}

