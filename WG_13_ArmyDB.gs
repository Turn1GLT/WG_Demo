// **********************************************
// function fcnUpdateArmyDB()
//
// This function updates the Player Army List  
//
// **********************************************

function fcnUpdateArmyDB(shtConfig, Player, NewValue, shtTest){
  
  Logger.log("Routine: fcnUpdateArmyDB");
  
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Config Spreadsheet
  var ssArmyDBID = shtIDs[2][0];
  var ssArmyListEnID = shtIDs[3][0];
  var ssArmyListFrID = shtIDs[4][0];

  var armyBldRatingMode = cfgArmyBuild[0][0];
  var armyBldCurrRoundValue = cfgArmyBuild[4][0];
  
  // Player Card DB Spreadsheet
  Logger.log("Player Army DB: %s",Player);
  var shtArmyDB = SpreadsheetApp.openById(ssArmyDBID).getSheetByName(Player);
  var rngArmyDBCurrRoundPwrLvl = shtArmyDB.getRange(5,9);
  var rngArmyDBAvailPwrLvl = shtArmyDB.getRange(5,10);
  var rngArmyDBCurrRoundPoints = shtArmyDB.getRange(5,11);
  var rngArmyDBAvailPoints = shtArmyDB.getRange(5,12);

  // Army List Spreadsheet
  var shtArmyListEn = SpreadsheetApp.openById(ssArmyListEnID).getSheetByName(Player);
  var rngArmyListEnCurrRoundPwrLvl = shtArmyListEn.getRange(5,9);
  var rngArmyListEnBonusPwrLvl = shtArmyListEn.getRange(5,10);
  var rngArmyListEnCurrRoundPoints = shtArmyListEn.getRange(5,11);
  var rngArmyListEnBonusPoints = shtArmyListEn.getRange(5,12);
  
  var shtArmyListFr = SpreadsheetApp.openById(ssArmyListFrID).getSheetByName(Player);
  var rngArmyListFrCurrRoundPwrLvl = shtArmyListFr.getRange(5,9);
  var rngArmyListFrBonusPwrLvl = shtArmyListFr.getRange(5,10);
  var rngArmyListFrCurrRoundPoints = shtArmyListFr.getRange(5,11);
  var rngArmyListFrBonusPoints = shtArmyListFr.getRange(5,12);
  
  // Get Cells to Update according to the Army Rating Mode (Power Level or Points)
  if(armyBldRatingMode == 'Power Level'){
    // Update the Army DB
    rngArmyDBCurrRoundPwrLvl.setValue(armyBldCurrRoundValue);
    rngArmyDBAvailPwrLvl.setValue(NewValue);
    
    // Update the Player Army List
    // English File
    rngArmyListEnCurrRoundPwrLvl.setValue(armyBldCurrRoundValue);
    rngArmyListEnBonusPwrLvl.setValue(NewValue);
    // French File
    rngArmyListFrCurrRoundPwrLvl.setValue(armyBldCurrRoundValue);
    rngArmyListFrBonusPwrLvl.setValue(NewValue);
    
    // Hide Points Columns (6-7-8, 11-12)
    shtArmyListEn.hideColumns(6, 3);
    shtArmyListEn.hideColumns(11, 2);    
    shtArmyListFr.hideColumns(6, 3);
    shtArmyListFr.hideColumns(11, 2);
  }

  if(armyBldRatingMode == 'Points'){
    // Update the Army DB
    rngArmyDBCurrRoundPoints.setValue(armyBldCurrRoundValue);
    rngArmyDBAvailPoints.setValue(NewValue);
    
    // Update the Player Army List
    // English File
    rngArmyListEnCurrRoundPoints.setValue(armyBldCurrRoundValue);
    rngArmyListEnBonusPoints.setValue(NewValue);
    // French File
    rngArmyListFrCurrRoundPoints.setValue(armyBldCurrRoundValue);
    rngArmyListFrBonusPoints.setValue(NewValue);
  
    // Hide Power Level Columns (5, 9-10)
    shtArmyListEn.hideColumns(5, 1);
    shtArmyListEn.hideColumns(9, 2);    
    shtArmyListFr.hideColumns(5, 1);
    shtArmyListFr.hideColumns(9, 2);
  }
}

