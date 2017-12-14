// **********************************************
// function fcnUpdateArmyDB()
//
// This function updates the Player Army List  
//
// **********************************************

function fcnUpdateArmyDB(ss, shtConfig, cfgColRndSht, Player){
  
  Logger.log("Routine: fcnUpdateArmyDB");
  
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Column Values for Rounds Sheets
  var colRndPlyr =     cfgColRndSht[ 0][0];
  var colRndBalBonus = cfgColRndSht[10][0];
  
  // Config Spreadsheet
  var ssArmyDBID =     shtIDs[2][0];
  var ssArmyListEnID = shtIDs[3][0];
  var ssArmyListFrID = shtIDs[4][0];

  var armyBldRating =       cfgArmyBuild[0][0];
  var armyBldCurrRoundVal = cfgArmyBuild[4][0];
  
  // Gets New Balance Bonus for Player from Cumulative Results Sheet
  var shtCumul = ss.getSheetByName('Cumulative Results');
  
  // Find Player Rows : subFindPlayerRow(sheet, rowStart, colPlyr, length, PlayerName)
  var RndPlyr2Row = subFindPlayerRow(shtCumul, 5, colRndPlyr, 32, Player);  
  
  var BalanceBonusVal = shtCumul.getRange(RndPlyr2Row,colRndBalBonus).getValue();
  Logger.log('Total Balance Bonus: %s',BalanceBonusVal);
  
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
  if(armyBldRating == 'Power Level'){
    // Update the Army DB
    rngArmyDBCurrRoundPwrLvl.setValue(armyBldCurrRoundVal);
    rngArmyDBAvailPwrLvl.setValue(BalanceBonusVal);
    
    // Update the Player Army List
    // English File
    rngArmyListEnCurrRoundPwrLvl.setValue(armyBldCurrRoundVal);
    rngArmyListEnBonusPwrLvl.setValue(BalanceBonusVal);
    // French File
    rngArmyListFrCurrRoundPwrLvl.setValue(armyBldCurrRoundVal);
    rngArmyListFrBonusPwrLvl.setValue(BalanceBonusVal);
    
    // Hide Points Columns (6-7-8, 11-12)
    shtArmyListEn.hideColumns(6, 3);
    shtArmyListEn.hideColumns(11, 2);    
    shtArmyListFr.hideColumns(6, 3);
    shtArmyListFr.hideColumns(11, 2);
  }

  if(armyBldRating == 'Points'){
    // Update the Army DB
    rngArmyDBCurrRoundPoints.setValue(armyBldCurrRoundVal);
    rngArmyDBAvailPoints.setValue(BalanceBonusVal);
    
    // Update the Player Army List
    // English File
    rngArmyListEnCurrRoundPoints.setValue(armyBldCurrRoundVal);
    rngArmyListEnBonusPoints.setValue(BalanceBonusVal);
    // French File
    rngArmyListFrCurrRoundPoints.setValue(armyBldCurrRoundVal);
    rngArmyListFrBonusPoints.setValue(BalanceBonusVal);
  
    // Hide Power Level Columns (5, 9-10)
    shtArmyListEn.hideColumns(5, 1);
    shtArmyListEn.hideColumns(9, 2);    
    shtArmyListFr.hideColumns(5, 1);
    shtArmyListFr.hideColumns(9, 2);
  }
}

