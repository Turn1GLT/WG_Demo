// **********************************************
// function fcnCreateReportForm_WG_S()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateReportForm_WG_S() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var ssID = shtIDs[0][0]; 
  var OptGenerateResp = shtConfig.getRange(64,2).getValue();
  var OptLocation = shtConfig.getRange(67, 2).getValue();
  
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  var shtRespMaxRow;
  var shtRespMaxCol;
  var FirstCellVal;
  
  var ssTexts = SpreadsheetApp.openById('1DkSr5HbGqZ_c38DlHKiBhgcBXw3fr3CK9zDE04187fE');
  var shtTxtReport = ssTexts.getSheetByName('Match Report WG');
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  
  var RoundNum = shtConfig.getRange(5,7).getValue();
  var RoundArray = new Array(1); RoundArray[0] = RoundNum;
  
  var PlayerNum = shtPlayers.getRange(2,1).getValue();
  var Players;
  var PlayerList;
  
  var PlayerWinList;
  var PlayerLosList;
  
  var ConfirmMsgEN;
  var ConfirmMsgFR;
  
  var RowFormUrlEN = 19;
  var RowFormUrlFR = 22;
  var RowFormIdEN = 23;
  var RowFormIdFR = 24;
  
  var ErrorVal = '';
  
  // Gets the Subscription ID from the Config File
  FormIdEN = shtConfig.getRange(RowFormIdEN, 7).getValue();
  FormIdFR = shtConfig.getRange(RowFormIdFR, 7).getValue();

  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdEN already exists. Unlink Response and Delete Form');
  }
  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdFR already exists. Unlink Response and Delete Form');
  }

  if (FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    
    //---------------------------------------------
    // TITLE SECTION
    // English
    FormNameEN = shtConfig.getRange(3, 2).getValue() + " Match Reporter EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    // French    
    FormNameFR = shtConfig.getRange(3, 2).getValue() + " Match Reporter FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Set Match Report Form Description
    formEN.setDescription("Please enter the following information to submit your match result");
    //formEN.setCollectEmail(true);
    
    formFR.setDescription("SVP, entrez les informations suivantes pour soumettre votre rapport de match");
    //formFR.setCollectEmail(true);
    
    //---------------------------------------------
    // PASSWORD SECTION
      
    // English
    formEN.addPageBreakItem().setTitle("League Password")
    formEN.addTextItem()
    .setTitle("League Password")
    .setHelpText("Please enter the league password to send your match report")
    .setRequired(true);
    
    // French
    formFR.addPageBreakItem().setTitle("Mot de passe de Ligue")
    formFR.addTextItem()
    .setTitle("Mot de passe de Ligue")
    .setHelpText("SVP, entrez le mot de passe de la ligue pour envoyer votre rapport de match")
    .setRequired(true);
    
    //---------------------------------------------
    // LOCATION SECTION
    // If Location Bonus is Enabled, add Location Section
    if (OptLocation == 'Enabled'){
      
      // English
      formEN.addPageBreakItem().setTitle("Location")
      formEN.addMultipleChoiceItem()
      .setTitle("Location")
      .setHelpText("Did you play at the store?")
      .setRequired(true)
      .setChoiceValues(["Yes","No"]);
      
      // French
      formFR.addPageBreakItem().setTitle("Localisation")
      formFR.addMultipleChoiceItem()
      .setTitle("Localisation")
      .setHelpText("Avez-vous joué au magasin?")
      .setRequired(true)
      .setChoiceValues(["Oui","Non"]);
    }
    
    //---------------------------------------------
    // ROUND NUMBER & PLAYERS SECTION
    
    // Transfers Players Double Array to Single Array
    if (PlayerNum > 0){
      Players = shtPlayers.getRange(3,2,PlayerNum,1).getValues();
      PlayerList = new Array(PlayerNum);
      for(var i = 0; i < PlayerNum; i++){
        PlayerList[i] = Players[i][0];
      }
    }
    
    // English
    formEN.addPageBreakItem().setTitle("Round Number & Players");
    // Round
    formEN.addListItem()
    .setTitle("Round")
    .setRequired(true)
    .setChoiceValues(RoundArray);
    
    // Winning Players
    PlayerWinList = formEN.addListItem()
    .setTitle("Winning Player")
    .setHelpText("If Game is a Tie, select any player")
    .setRequired(true);
    if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
    
    // Losing Players
    PlayerLosList = formEN.addListItem()
    .setTitle("Losing Player")
    .setHelpText("If Game is a Tie, select any player")
    .setRequired(true);
    if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
    
    // Score
    formEN.addMultipleChoiceItem()
    .setTitle("Game is a Tie?")
    .setRequired(true)
    .setChoiceValues(["No","Yes"]);
    
    // French
    formFR.addPageBreakItem().setTitle("Numéro de Semaine & Joueurs");
    // Semaine
    formFR.addListItem()
    .setTitle("Semaine Numéro")
    .setRequired(true)
    .setChoiceValues(RoundArray);
    
    // Joueurs
    PlayerWinList = formFR.addListItem()
    .setTitle("Joueur Gagnant")
    .setHelpText("Si la partie est nulle, entrez n'importe quel joueur")
    .setRequired(true);
    if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
    
    PlayerLosList = formFR.addListItem()
    .setTitle("Joueur Perdant")
    .setHelpText("Si la partie est nulle, entrez n'importe quel joueur")
    .setRequired(true);
    if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
    
    // Score
    formFR.addMultipleChoiceItem()
    .setTitle("Partie est Nulle?")
    .setRequired(true)
    .setChoiceValues(["Non","Oui"]);

    
    //---------------------------------------------
    // CONFIRMATION MESSAGE
    
    // English
    ConfirmMsgEN = shtTxtReport.getRange(4,2).getValue();
    formEN.setConfirmationMessage(ConfirmMsgEN);
    
    // French
    ConfirmMsgFR = shtTxtReport.getRange(4,3).getValue();
    formFR.setConfirmationMessage(ConfirmMsgFR);
    
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    if(OptGenerateResp == 'Enabled'){
      // English Form
      formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('New Responses EN');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('New Responses EN');
      ss.moveActiveSheet(15);
      shtRespMaxRow = shtResp.getMaxRows();
      shtRespMaxCol = shtResp.getMaxColumns();
      
      // Delete All Empty Rows
      shtResp.deleteRows(3, shtRespMaxRow - 2);
      
      // Delete All Empty Columns
      for(var c = 1;  c <= shtRespMaxCol; c++){
        FirstCellVal = shtResp.getRange(1, c).getValue();
        if(FirstCellVal == '') {
          shtResp.deleteColumns(c,shtRespMaxCol-c+1);
          c = shtRespMaxCol + 1;
        }
      }
      
      // French Form
      formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('New Responses FR');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('New Responses FR');
      ss.moveActiveSheet(16);
      shtRespMaxRow = shtResp.getMaxRows();
      shtRespMaxCol = shtResp.getMaxColumns();
      
      // Delete All Empty Rows
      shtResp.deleteRows(3, shtRespMaxRow - 2);
      
      // Delete All Empty Columns
      for(var c = 1;  c <= shtRespMaxCol; c++){
        FirstCellVal = shtResp.getRange(1, c).getValue();
        if(FirstCellVal == '') {
          shtResp.deleteColumns(c,shtRespMaxCol-c+1);
          c = shtRespMaxCol + 1;
        }
      }
      
      // Set Match Report IDs in Config File
      FormIdEN = formEN.getId();
      shtConfig.getRange(RowFormIdEN, 7).setValue(FormIdEN);
      FormIdFR = formFR.getId();
      shtConfig.getRange(RowFormIdFR, 7).setValue(FormIdFR);
      
      // Create Links to add to Config File  
      shtConfig.getRange(RowFormUrlEN, 2).setValue(formEN.getPublishedUrl()); 
      shtConfig.getRange(RowFormUrlFR, 2).setValue(formFR.getPublishedUrl());
    }
  }
}