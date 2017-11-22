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
  var cfgReportFormCnstrVal = shtConfig.getRange(4,30,20,2).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // Event Properties
  var evntLocation = cfgEvntParam[0][0];
  var evntName = cfgEvntParam[7][0];
  
  var evntFormat = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  var evntLocationBonus = cfgEvntParam[23][0];
  var evntMatchPtsMin = 0;
  var evntMatchPtsMax = cfgEvntParam[29][0];
  
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
  var ConfirmMsgEN = shtTxtReport.getRange(4,2).getValue();
  var ConfirmMsgFR = shtTxtReport.getRange(4,3).getValue();
  
  var QuestionOrder = 2;
    
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
  
  var ErrorVal = '';
  
  // Insert ui to confirm
  var ui = SpreadsheetApp.getUi();
  var title;
  var msg;
  var uiResponse;

  // If Form Exists, Log Error Message
  if(FormIdEN != '' || FormIdFR != ''){
    ErrorVal = 1;
    title = "Match Report Forms Error";
    msg = "The Match Report Forms already exist. Unlink their response sheets then delete the forms and their ID in the configuration file.";
    var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }

  // CREATE UNIT VALIDATIONS
  // Number of Models in Unit
  var PointsValidationEN = FormApp.createTextValidation()
  .setHelpText("Enter a number between " + evntMatchPtsMin + " and " + evntMatchPtsMax)
  .requireNumberBetween(evntMatchPtsMin, evntMatchPtsMax)
  .build();
  
  var PointsValidationFR = FormApp.createTextValidation()
  .setHelpText("Entrez un nombre entre " + evntMatchPtsMin + " et " + evntMatchPtsMax)
  .requireNumberBetween(evntMatchPtsMin, evntMatchPtsMax)
  .build();
  
  // Create Forms
  if (FormIdEN == '' && FormIdFR == ''){
    
    //---------------------------------------------
    // TITLE SECTION
    // English
    FormNameEN = evntLocation + " " + evntName + " Match Reporter EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN)
    .setDescription("Please enter the following information to submit your match result");
    // French    
    FormNameFR = evntLocation + " " + evntName + " Match Reporter FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR)
    .setDescription("SVP, entrez les informations suivantes pour soumettre votre rapport de match");
    
    // Transfers Players Double Array to Single Array
    if (PlayerNum > 0){
      Players = shtPlayers.getRange(3,2,PlayerNum,1).getValues();
      PlayerList = new Array(PlayerNum);
      for(var i = 0; i < PlayerNum; i++){
        PlayerList[i] = Players[i][0];
      }
    }
    
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 1; i < cfgReportFormCnstrVal.length; i++){
      // Look for Col Equal to Question Order
      if(QuestionOrder == cfgReportFormCnstrVal[i][1]){
        switch(cfgReportFormCnstrVal[i][0]){
            
            //---------------------------------------------
            // PASSWORD SECTION
          case 'Password':{ 
            // English
            formEN.addTextItem()
            .setTitle("Event Password")
            .setHelpText("Please enter the Event Password to send your match report")
            .setRequired(true);
            
            // French
            formFR.addTextItem()
            .setTitle("Mot de passe de l'événement")
            .setHelpText("SVP, entrez le mot de passe de l'événement pour envoyer votre rapport de match")
            .setRequired(true);
            
            break;
          }
            
            //---------------------------------------------
            // LOCATION SECTION
          case 'Location':{ 
            // English
            formEN.addPageBreakItem().setTitle("Location")
            formEN.addMultipleChoiceItem()
            .setTitle("Location Bonus")
            .setHelpText("Did you play at the store?")
            .setRequired(true)
            .setChoiceValues(["Yes","No"]);
            
            // French
            formFR.addPageBreakItem().setTitle("Localisation")
            formFR.addMultipleChoiceItem()
            .setTitle("Bonus de Localisation")
            .setHelpText("Avez-vous joué au magasin?")
            .setRequired(true)
            .setChoiceValues(["Oui","Non"]);
            
            break;
          }
            
            //---------------------------------------------
            // ROUND NUMBER
          case 'Round Number':{ 
            // English
            if(evntFormat == 'Single') formEN.addPageBreakItem().setTitle("Round Number & Players");
            if(evntFormat == 'Team')   formEN.addPageBreakItem().setTitle("Round Number & Teams");
            // Round
            formEN.addListItem()
            .setTitle("Round")
            .setRequired(true)
            .setChoiceValues(RoundArray);
                
            // French
            if(evntFormat == 'Single') formFR.addPageBreakItem().setTitle("Numéro de Semaine & Joueurs");
            if(evntFormat == 'Team')   formFR.addPageBreakItem().setTitle("Numéro de Semaine & Équipes");
            
            // Semaine
            formFR.addListItem()
            .setTitle("Ronde")
            .setRequired(true)
            .setChoiceValues(RoundArray);
            
            break;
          }
            
            //---------------------------------------------
            // PLAYERS
            // Winning Player List
          case 'Winning Player':{ 
            // English
            PlayerWinList = formEN.addListItem()
            .setTitle("Winning Player")
            .setHelpText("If Game is a Tie, select your name")
            .setRequired(true);
            if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
            
            // French
            PlayerWinList = formFR.addListItem()
            .setTitle("Joueur Gagnant")
            .setHelpText("Si la partie est nulle, sélectionnez votre nom")
            .setRequired(true);
            if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
            
            break;
          }
            // Losing Player List
          case 'Losing Player':{ 
            // English
            PlayerLosList = formEN.addListItem()
            .setTitle("Losing Player")
            .setHelpText("If Game is a Tie, select your opponent")
            .setRequired(true);
            if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList); 
            
            // French
            PlayerLosList = formFR.addListItem()
            .setTitle("Joueur Perdant")
            .setHelpText("Si la partie est nulle, sélectionnez votre adversaire")
            .setRequired(true);
            if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
            
            break;
          }
            
            //---------------------------------------------
            // TEAMS
            // Winning Player List
          case 'Winning Team':{ 
            // English
            PlayerWinList = formEN.addListItem()
            .setTitle("Winning Team")
            .setHelpText("If Game is a Tie, select your team")
            .setRequired(true);
            if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
            
            // French
            PlayerWinList = formFR.addListItem()
            .setTitle("Équipe Gagnante")
            .setHelpText("Si la partie est nulle, sélectionnez votre équipe")
            .setRequired(true);
            if (PlayerNum > 0) PlayerWinList.setChoiceValues(PlayerList);
            
            break;
          }
            // Losing Player List
          case 'Losing Team':{ 
            // English
            PlayerLosList = formEN.addListItem()
            .setTitle("Losing Team")
            .setHelpText("If Game is a Tie, select the opposing team")
            .setRequired(true);
            if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList); 
            
            // French
            PlayerLosList = formFR.addListItem()
            .setTitle("Équipe Perdante")
            .setHelpText("Si la partie est nulle, sélectionnez l'équipe adverse")
            .setRequired(true);
            if (PlayerNum > 0) PlayerLosList.setChoiceValues(PlayerList);
            
            break;
          }
            
            //---------------------------------------------
            // GAME TIE
          case 'Game is Tie':{
            // English
            formEN.addMultipleChoiceItem()
            .setTitle("Game is a Tie?")
            .setRequired(true)
            .setChoiceValues(["No","Yes"]);
            
            // French
            formFR.addMultipleChoiceItem()
            .setTitle("Partie est Nulle?")
            .setRequired(true)
            .setChoiceValues(["Non","Oui"]);
            break;
          }
            
            //---------------------------------------------
            // WINNING POINTS
          case 'Winning Points':{ 
            // English
            formEN.addTextItem()
            .setTitle("Points: Winner")
            .setHelpText("Enter the points scored by the Winner")
            .setValidation(PointsValidationEN)
            .setRequired(true);
            
            // French
            formFR.addTextItem()
            .setTitle("Points: Gagnant")
            .setHelpText("Entrez les points accumulés par le Gagnant")
            .setValidation(PointsValidationFR)
            .setRequired(true);
            
            break;
          }
            
            //---------------------------------------------
            // LOSING POINTS
          case 'Losing Points':{ 
            // English
            formEN.addTextItem()
            .setTitle("Points: Loser")
            .setHelpText("Enter the points scored by the Loser")
            .setValidation(PointsValidationEN)
            .setRequired(true);
            
            // French
            formFR.addTextItem()
            .setTitle("Points: Perdant")
            .setHelpText("Entrez les points accumulés par le Perdant")
            .setValidation(PointsValidationFR)
            .setRequired(true);
            
            break;
          }
          default : break;
        }
        // Increment to Next Question
        QuestionOrder++;
        // Reset Loop if new question was added
        i = -1;
      }
    }


    
    //---------------------------------------------
    // CONFIRMATION MESSAGE
    
    // English
    formEN.setConfirmationMessage(ConfirmMsgEN);
    
    // French
    formFR.setConfirmationMessage(ConfirmMsgFR);
    
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    if(exeGnrtResp == 'Enabled'){
      Logger.log("Generating Response Sheets and Form Links");
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
      
      Logger.log("Response Sheets and Form Links Generated");
    }
  }

  // Post Log to Log Sheet
  subPostLog(shtLog,Logger.getLog());
  
}