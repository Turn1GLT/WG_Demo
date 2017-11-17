// **********************************************
// function fcnCrtMatchReportForm_WG_S()
//
// This function creates the Match Report Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtMatchReportForm_WG_S() {
  
  Logger.log("Routine: fcnCrtMatchReportForm_WG_S");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
    
  // Configuration Data
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,32,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,16,3).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // League Parameters
  var evntFormat = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  var evntLocationBonus = cfgEvntParam[23][0];
    
  var RoundNum = shtConfig.getRange(7,2).getValue();
  var RoundArray = new Array(1); RoundArray[0] = RoundNum;
  
  var PlayerNum = shtConfig.getRange(13,2).getValue();
  
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');

  // Registration ID from the Config File
  var ssID = shtIDs[0][0];
  var FormIdEN = shtIDs[7][0];
  var FormIdFR = shtIDs[8][0];
 
  // Row Column Values to Write Form IDs and URLs
  var rowFormEN  = 11;
  var rowFormFR  = 12;
  var colFormID  = 7;
  var colFormURL = 11
  
  var ssTexts = SpreadsheetApp.openById('1DkSr5HbGqZ_c38DlHKiBhgcBXw3fr3CK9zDE04187fE');
  var shtTxtReport = ssTexts.getSheetByName('Match Report WG');
    
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  var shtRespMaxRow;
  var shtRespMaxCol;
  var FirstCellVal;
  
  var formEN;
  var FormIdEN;
  var FormNameEN;
  var FormItemsEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  
  var Players;
  var PlayerList;
  
  var PlayerWinList;
  var PlayerLosList;
  
  var ConfirmMsgEN;
  var ConfirmMsgFR;
  
  var ErrorVal = '';

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
    if (evntLocationBonus == 'Enabled'){
      
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
    if(exeGnrtResp == 'Enabled'){
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
      shtConfig.getRange(rowFormEN, colFormID).setValue(FormIdEN);
      FormIdFR = formFR.getId();
      shtConfig.getRange(rowFormFR, colFormID).setValue(FormIdFR);
      
      // Create Links to add to Config File  
      shtConfig.getRange(rowFormEN, colFormURL).setValue(formEN.getPublishedUrl()); 
      shtConfig.getRange(rowFormFR, colFormURL).setValue(formFR.getPublishedUrl());
    }
  }
}