// **********************************************
// function fcnUpdateArmyList()
//
// This function updates the Player card database  
// with the list of cards sent in arguments
//
// **********************************************

function fcnUpdateArmyList(shtConfig, shtArmyDB, Player, shtTest){
  
  // Config Spreadsheet
  var ssArmyListsEnID = shtConfig.getRange(32,2).getValue();
  var ssArmyListsFrID = shtConfig.getRange(33,2).getValue();
  
  // Army DB Spreadsheet
  var ArmyValues = shtArmyDB.getRange(1, 1, 66, 12).getValues();  
    
  // Army List Spreadsheet
  var shtArmyListsEn = SpreadsheetApp.openById(ssArmyListsEnID).getSheetByName(Player);
  var shtArmyListsFr = SpreadsheetApp.openById(ssArmyListsFrID).getSheetByName(Player);
  var rngArmyListsEn = shtArmyListsEn.getRange(1, 1, 66, 12);  
  var rngArmyListsFr = shtArmyListsFr.getRange(1, 1, 66, 12);  

 

}