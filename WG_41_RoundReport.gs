// **********************************************
// function fcnGenerateRoundReport()
//
// This function analyzes all players records
// and adds a loss to a player who has not played
// the minimum amount of games. This also 
//
// **********************************************

function fcnGenerateRoundReport(){


}





// **********************************************
// function fcnModifyRoundMatchReport()
//
// This function modifies the Round Number in 
// the Match Report Form
//
// **********************************************

function fcnModifyRoundMatchReport(ss, shtConfig){
  
  Logger.log("Routine: fcnModifyRoundMatchReport");

  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var MatchFormEN = FormApp.openById(shtIDs[7][0]);
  var FormItemEN = MatchFormEN.getItems();
  var NbFormItem = FormItemEN.length;
  
  var MatchFormFR = FormApp.openById(shtIDs[8][0]);
  var FormItemFR = MatchFormFR.getItems();
  
  var Round = shtConfig.getRange(7,2).getValue();

  // Function Variables
  var ItemTitle;
  var ItemListEN;
  var ItemListFR;
  var ItemChoice;
  var RoundChoice = [];
  
  // Loops to Find Players List
  for(var item = 0; item < NbFormItem; item++){
    ItemTitle = FormItemEN[item].getTitle();
    if(ItemTitle == 'Round'){
      
      // Get the List Item from the Match Report Form
      ItemListEN = FormItemEN[item].asListItem();
      ItemListFR = FormItemFR[item].asListItem();
      
      // Set the New Choice for Item
      RoundChoice[0] = Round;
      
      // Set the Item Choices in the Match Report Forms
      ItemListEN.setChoiceValues(RoundChoice);
      ItemListFR.setChoiceValues(RoundChoice);
      
      // Exit For
      item = NbFormItem;
    }
  }
}

// **********************************************
// function fcnAnalyzeLossPenalty()
//
// This function analyzes all players records
// and adds a loss to a player who has not played
// the minimum amount of games. This also 
//
// **********************************************

function fcnAnalyzeLossPenalty(ss, Round, PlayerData){
  
  Logger.log("Routine: fcnAnalyzeLossPenalty");
  
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var CumulMaxCol = shtCumul.getMaxColumns();
  var RoundShtName = 'Round'+Round;
  var shtRound = ss.getSheetByName(RoundShtName);
  var MissingMatch;
  var Loss;
  var PlayerDataPntr = 0;
  
  var colCumulLoss = 6;
  var colCumulMiss = 14;
  
  var shtTest = ss.getSheetByName('Test');
  
  // Get Player Record Range
  var RngCumul = shtCumul.getRange(5,1,32,CumulMaxCol);
  var ValCumul = RngCumul.getValues(); 
  // 0= Player ID, 1= Player Name, 2= Team Name, 3= MP, 4= Win, 5= Loss, 6= Tie, 7= Points, 8= Win%, 9= Matches in Store, 10= Penalty Losses, 11= Balance Bonus (Packs, Bonus Pts), 12= Status, 13= Matches Missing, 14= Warning 
  
  for (var plyr = 0; plyr < 32; plyr++){
    // If Player Exists
    if (ValCumul[plyr][1] != ''){      
      // Check if player has matches missing
      if (ValCumul[plyr][13] > 0){
        // Saves Missing Match and Losses
        MissingMatch = ValCumul[plyr][13];
        Loss = ValCumul[plyr][5];
        // Updates Losses
        Loss = Loss + MissingMatch;
        
        // Updates Round Results Sheet 
        shtRound.getRange(plyr+5,colCumulLoss).setValue(Loss);
        shtRound.getRange(plyr+5,colCumulMiss).setValue(MissingMatch);
        
        // Saves Player and Missing Matches for Roundly Report
        PlayerData[PlayerDataPntr][0] = ValCumul[plyr][0];
        PlayerData[PlayerDataPntr][1] = MissingMatch;
        PlayerDataPntr++;
      }
    }
    // Exit when the loop reaches the end of the list 
    if (ValCumul[plyr][0] == '') plyr = 32;
  }
  return PlayerData;
}
