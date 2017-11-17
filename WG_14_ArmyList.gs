// **********************************************
// function fcnUpdateArmyList()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateArmyList(shtConfig, shtArmyDB, Player, shtTest){
  
  Logger.log("Routine: fcnUpdateArmyList");
  
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  
  // Army Lists
  var ssArmyListEnID = shtIDs[3][0];
  var ssArmyListFrID = shtIDs[4][0];
  
  // Army DB Spreadsheet
  var ArmyDBMaxRow = shtArmyDB.getMaxRows();
  var ArmyDBMaxCol = shtArmyDB.getMaxColumns();
  // Get Values from Army DB
  var ArmyValues = shtArmyDB.getRange(1, 1, ArmyDBMaxRow, ArmyDBMaxCol).getValues();  
    
  // Army List Spreadsheet
  var shtArmyListsEn = SpreadsheetApp.openById(ssArmyListEnID).getSheetByName(Player);
  var shtArmyListsFr = SpreadsheetApp.openById(ssArmyListFrID).getSheetByName(Player);

  // Paste Values from Army DB
  shtArmyListsEn.getRange(1, 1, ArmyDBMaxRow, ArmyDBMaxCol).setValues(ArmyValues);;  
  shtArmyListsFr.getRange(1, 1, ArmyDBMaxRow, ArmyDBMaxCol).setValues(ArmyValues);;  
}