// **********************************************
// function fcnCrtRegstnFormTeam_WG()
//
// This function creates the Team Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtRegstnFormTeam_WG() {
  
  Logger.log("Routine: fcnCreateRegForm_WG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
    
  // Configuration Data
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(24,26,20,3).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // Event Properties
  var evntLocation = cfgEvntParam[0][0];
  var evntName = cfgEvntParam[7][0];
  var evntFormat = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
    
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
  // Registration ID from the Config File
  var ssID = shtIDs[0][0];
  var FormIdEN = shtIDs[11][0];
  var FormIdFR = shtIDs[12][0];
 
  // Row Column Values to Write Form IDs and URLs
  var rowFormEN  = 15;
  var rowFormFR  = 16;
  var colFormID  = 7;
  var colFormURL = 11
  
  var ErrorVal = '';
  var QuestionOrder = 2;
  
  // Routine Variables
  var ssSheets;
  var NbSheets;
  var SheetName;
  var shtResp;
  var shtRespMaxRow;
  var shtRespMaxCol;
  var FirstCellVal;
    
  var formEN;
  var FormNameEN;
  var FormItemsEN;
  var urlFormEN;
  
  var formFR;
  var FormNameFR;
  var FormItemsFR;
  var urlFormFR;

  var TestCol = 1;
  
  // If Form Exists, Log Error Message
  if(FormIdEN != '' || FormIdFR != ''){
    ErrorVal = 1;
    var ui = SpreadsheetApp.getUi();
    var title = "Registration Forms Error";
    var msg = "The Registration Forms already exist. Unlink and delete their response sheets then delete the forms and their ID in the configuration file.";
    var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
  
  // If Form does not exist, create it
  if(FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    FormNameEN = evntLocation + " " + evntName + " Team Registration EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    
    FormNameFR = evntLocation + " " + evntName + " Team Registration FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 1; i < cfgRegFormCnstrVal.length; i++){
      // Check for Question Order in Response Column Value in Configuration File
      if(QuestionOrder == cfgRegFormCnstrVal[i][1]){
        
        switch(cfgRegFormCnstrVal[i][0]){
           
            // EMAIL
          case 'Contact Email': {
            // Set Registration Email collection
            formEN.setCollectEmail(true);
            formFR.setCollectEmail(true);
            break;
          }
            // FULL NAME
          case 'Full Name': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Name")
            .setHelpText("Please, Remove any space at the beginning or end of the name")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Nom")
            .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
            .setRequired(true);
            break;
          }
            // FIRST NAME
          case 'Contact First Name': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Team Contact First Name")
            .setHelpText("Please, Remove any space at the beginning or end of the name")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Prénom du Contact de l'équipe")
            .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
            .setRequired(true);
            break;
          }
            // LAST NAME
          case 'Contact Last Name': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Team Contact Last Name")
            .setHelpText("Please, Remove any space at the beginning or end of the name")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Nom de Famille du Contact de l'équipe")
            .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
            .setRequired(true);
            break;
          }
            // LANGUAGE
          case 'Contact Language': {
            // ENGLISH
            formEN.addMultipleChoiceItem()
            .setTitle("Team Contact Language Preference")
            .setHelpText("Which Language do you prefer to use? The application is available in English and French")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            
            // FRENCH
            formFR.addMultipleChoiceItem()
            .setTitle("Préférence de Langue du Contact de l'équipe")
            .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            break;
          }
          // PHONE NUMBER
          case 'Contact Phone Number': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Team Contact Phone Number")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Numéro de téléphone du Contact de l'équipe")
            .setRequired(true);
            break;
          }
            // TEAM NAME & MEMBERS
          case 'Team Name': {
            if(evntFormat == 'Team'){
              // ENGLISH
              formEN.addPageBreakItem().setTitle("Team");
              formEN.addTextItem()
              .setTitle("Team Name")
              .setRequired(true);
              
              // FRENCH
              formFR.addPageBreakItem().setTitle("Équipe");
              formFR.addTextItem()
              .setTitle("Nom d'équipe")
              .setRequired(true);
              
              for(var member = 1; member <= evntNbPlyrTeam; member++){
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Team Member " + member + " Name")
                .setRequired(true);
                
                // FRENCH
                formFR.addTextItem()
                .setTitle("Nom Membre d'équipe " + member)
                .setRequired(true);
              }
            }
            break;
          }   
        }
        // Increment to Next Question
        QuestionOrder++;
        // Reset Loop if new question was added
        i = -1;
      }
    }
    
    // RESPONSE SHEETS
    // Create Response Sheet in Main File and Rename
    if(exeGnrtResp == 'Enabled'){
      Logger.log("Generating Response Sheets and Form Links");
      var IndexTeams = ss.getSheetByName("Teams").getIndex();
      
      // English Form
      formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('Reg Team EN');
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('Reg Team EN');
      ss.moveActiveSheet(IndexTeams+1);
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
      formFR.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('Reg Team FR');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('Reg Team FR');
      ss.moveActiveSheet(IndexTeams+2);
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
      urlFormEN = formEN.getPublishedUrl();
      shtConfig.getRange(rowFormEN, colFormURL).setValue(urlFormEN); 
      
      urlFormFR = formFR.getPublishedUrl();
      shtConfig.getRange(rowFormFR, colFormURL).setValue(urlFormFR);
      
      Logger.log("Response Sheets and Form Links Generated");
    }
  }
  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}