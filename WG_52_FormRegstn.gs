// **********************************************
// function fcnCrtRegstnForm_WG()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtRegstnForm_WG() {
  
  Logger.log("Routine: fcnCreateRegForm_WG");
  
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
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,20,3).getValues();
  
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
  var FormIdEN = shtIDs[9][0];
  var FormIdFR = shtIDs[10][0];
 
  // Row Column Values to Write Form IDs and URLs
  var rowFormEN  = 13;
  var rowFormFR  = 14;
  var colFormID  = 7;
  var colFormURL = 11
  
  var ErrorVal = '';
  var QuestionOrder = 2;
  
  // Army Building Options
  var ArmyRating = cfgArmyBuild[0][0]; 
  var NbFaction     = cfgArmyBuild[6][0];
  var NbDetachMax   = cfgArmyBuild[7][0];
  var NbUnitDetach1 = cfgArmyBuild[8][0];
  var NbUnitDetach2 = cfgArmyBuild[9][0];
  var NbUnitDetach3 = cfgArmyBuild[10][0];
  var UnitModelMin  = 1;
  var UnitModelMax  = cfgArmyBuild[11][0];
  var UnitRatingMin = 1;
  var UnitRatingMax = cfgArmyBuild[12][0];
  
  var DetachList = shtConfig.getRange(4,34,20,2).getValues();
  var DetachIncr = 0;
  var DetachTypeArray = new Array(12);
  
  var UnitRolesList = shtConfig.getRange(4,37,10,2).getValues();
  var UnitIncr = 0;
  var UnitRoleArray = new Array(9);
  
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
  
  var ChcUnitEN;
  var ChcDetachEN;
  var ChcEndEN;

  var ChcUnitFR;
  var ChcDetachFR;
  var ChcEndFR;

  var ArmyRatingText;
  var NbUnitMax;
  var UnitPageEN = new Array(325);
  var UnitPageFR = new Array(325);
  var UnitIndex;
  var UnitTitle;
  var UnitRole;
  var TestCol = 1;

  var ChcNbDetachArray;
  // Number of Detachment in Army
  if(NbDetachMax == 1) ChcNbDetachArray = ["1"];
  if(NbDetachMax == 2) ChcNbDetachArray = ["1","2"];
  if(NbDetachMax == 3) ChcNbDetachArray = ["1","2","3"];
  
  // If Form Exists, Log Error Message
  if(FormIdEN != '' || FormIdFR != ''){
    ErrorVal = 1;
    title = "Registration Forms Error";
    msg = "The Registration Forms already exist. Unlink and delete their response sheets then delete the forms and their ID in the configuration file.";
    var uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
  
  // If Form does not exist, create it
  if(FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    FormNameEN = evntLocation + " " + evntName + " Registration EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    
    FormNameFR = evntLocation + " " + evntName + " Registration FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 1; i < cfgRegFormCnstrVal.length; i++){
      // Check for Question Order in Response Column Value in Configuration File
      if(QuestionOrder == cfgRegFormCnstrVal[i][1]){
        switch(cfgRegFormCnstrVal[i][0]){
            // EMAIL
          case 'Email': {
            // Set Registration Email collection
            formEN.setCollectEmail(true);
            formFR.setCollectEmail(true);
            break;
          }
            // FULL NAME
          case 'Name': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Name")
            .setHelpText("Please, Remove any space at the end of the name")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Nom")
            .setHelpText("SVP, enlevez les espaces à la fin du nom")
            .setRequired(true);
            break;
          }
            // LANGUAGE
          case 'Language': {
            // ENGLISH
            formEN.addMultipleChoiceItem()
            .setTitle("Language Preference")
            .setHelpText("Which Language do you prefer to use? The application is available in English and French")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            
            // FRENCH
            formFR.addMultipleChoiceItem()
            .setTitle("Préférence de Langue")
            .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            break;
          }
          // PHONE NUMBER
          case 'Phone Number': {
            // ENGLISH
            formEN.addTextItem()
            .setTitle("Phone Number")
            .setRequired(true);
            
            // FRENCH
            formFR.addTextItem()
            .setTitle("Numéro de téléphone")
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
                .setTitle("Teammate " + member)
                .setRequired(true);
                
                // FRENCH
                formFR.addTextItem()
                .setTitle("Équipier " + member)
                .setRequired(true);
              }
            }
            break;
          }   
            // ARMY DEFINITION
          case 'Army Definition': {
            // ENGLISH
            formEN.addPageBreakItem()
            .setTitle("Army Definition")
            .setHelpText("Enter Army's General Information");
            
            // Faction
            if (NbFaction == 1){
              // Faction Keyword 1
              formEN.addTextItem()
              .setTitle("Faction")
              .setHelpText("Enter your Main Faction Keyword")
              .setRequired(true);  
            }
            if (NbFaction == 2){
              // Faction Keyword 1
              formEN.addTextItem()
              .setTitle("Faction 1")
              .setHelpText("Enter your First Faction Keyword")
              .setRequired(true);  
              
              // Faction Keyword 2
              formEN.addTextItem()
              .setTitle("Faction 2")
              .setHelpText("Enter your Second Faction Keyword")
              .setRequired(true);
            }
            
            // Warlord name
            formEN.addTextItem()
            .setTitle("Army Warlord")
            .setHelpText("Enter your Army's Warlord name (or Unit Entry)")
            .setRequired(true); 
            
            // Army name
            formEN.addTextItem()
            .setTitle("Army Name")
            .setHelpText("Enter your Army's Name (optional)")
            .setRequired(false); 
            
            // FRENCH
            formFR.addPageBreakItem()
            .setTitle("Définition d'Armée")
            .setHelpText("Entrez les informations générales de votre armée");
            
            // Faction
            if (NbFaction == 1){
              // Faction Keyword 1
              formFR.addTextItem()
              .setTitle("Faction")
              .setHelpText("Entrez le mot-clé Faction principal de votre armée")
              .setRequired(true);  
            }
            if (NbFaction == 2){
              // Faction Keyword 1
              formFR.addTextItem()
              .setTitle("Faction 1")
              .setHelpText("Entrez le premier mot-clé Faction de votre armée")
              .setRequired(true);  
              
              // Faction Keyword 2
              formFR.addTextItem()
              .setTitle("Faction 2")
              .setHelpText("Entrez le deuxième mot-clé Faction de votre armée")
              .setRequired(true);
            }
            
            // Warlord name
            formFR.addTextItem()
            .setTitle("Seigneur de Guerre de l'armée")
            .setHelpText("Entrez le nom (ou type d'unité) du Seigneur de Guerre de votre armée")
            .setRequired(true); 
            
            // Army name
            formFR.addTextItem()
            .setTitle("Nom d'Armée")
            .setHelpText("Entrez le nom de votre Armée (optionel)")
            .setRequired(false);
            break;
          }
          case 'Army List': {
            Logger.log('%s - %s',QuestionOrder,cfgRegFormCnstrVal[i][0]);  
            // Army List
            // ENGLISH
            formEN.addPageBreakItem()
            .setTitle("Army List")
            .setHelpText("Please, enter your Army List");
            // Number of Detachment in Army
            formEN.addMultipleChoiceItem()
            .setTitle("Number of Detachments in your Army")
            .setRequired(true)
            .setChoiceValues(ChcNbDetachArray);

            // FRENCH
            formFR.addPageBreakItem()
            .setTitle("Liste d'Armée")
            .setHelpText("SVP, entrez votre liste d'armée");
            // Number of Detachment in Army
            formFR.addMultipleChoiceItem()
            .setTitle("Nombre de Détachements dans votre armée")
            .setRequired(true)
            .setChoiceValues(ChcNbDetachArray);
            
            // CREATE DETACHMENT CHOICES
            // Creates the List of Detachments Allowed for League
            for(var detach = 0; detach < 20; detach++) {
              if(DetachList[detach][1] == 'Yes') {
                DetachTypeArray[DetachIncr] = DetachList[detach][0];
                DetachIncr++;
              }
            }
            DetachTypeArray.length = DetachIncr;
            
            // CREATE UNIT ROLES CHOICES
            // Creates the List of Unit Roles Allowed for League
            for(var unit = 1; unit <= 9; unit++) {
              if(UnitRolesList[unit][1] == 'Yes') {
                UnitRoleArray[UnitIncr] = UnitRolesList[unit][0];
                UnitIncr++;
              }
            }
            UnitRoleArray.length = UnitIncr;
            
            // CREATE UNIT VALIDATIONS
            // Number of Models in Unit
            var ModelValidationEN = FormApp.createTextValidation()
            .setHelpText("Enter a number between " + UnitModelMin + " and " + UnitModelMax)
            .requireNumberBetween(UnitModelMin, UnitModelMax)
            .build();

            var ModelValidationFR = FormApp.createTextValidation()
            .setHelpText("Entrez un nombre entre " + UnitModelMin + " et " + UnitModelMax)
            .requireNumberBetween(UnitModelMin, UnitModelMax)
            .build();
            
            // Unit Rating (Points, Power Level etc...)
            var RatingValidationEN = FormApp.createTextValidation()
            .setHelpText("Enter a number between " + UnitRatingMin + " and " + UnitRatingMax)
            .requireNumberBetween(UnitRatingMin, UnitRatingMax)
            .build();
            
            var RatingValidationFR = FormApp.createTextValidation()
            .setHelpText("Entrez un nombre entre " + UnitRatingMin + " et " + UnitRatingMax)
            .requireNumberBetween(UnitRatingMin, UnitRatingMax)
            .build();            
            
            // DETACHMENT 1
            if(DetachTypeArray.length > 0){
              
              // ENGLISH
              var Detach1EN = formEN.addPageBreakItem().setTitle("Detachment 1");
              // Detachment Name
              formEN.addTextItem()
              .setTitle("Detachment 1 Name")
              .setRequired(true);
              // Detachment Type
              formEN.addListItem()
              .setTitle("Detachment 1 Type")
              .setRequired(true)
              .setChoiceValues(DetachTypeArray);
              
              // FRENCH
              var Detach1FR = formFR.addPageBreakItem().setTitle("Détachement 1")
              // Detachment Name
              formFR.addTextItem()
              .setTitle("Nom du Détachement 1")
              .setRequired(true);
              // Detachment Type
              formFR.addListItem()
              .setTitle("Type du Détachment 1")
              .setRequired(true)
              .setChoiceValues(DetachTypeArray);
              
              // DETACHMENT 2
              if(NbDetachMax >= 2){
                
                // ENGLISH
                var Detach2EN = formEN.addPageBreakItem().setTitle("Detachment 2");
                // Detachment Name
                formEN.addTextItem()
                .setTitle("Detachment 2 Name")
                .setRequired(true);
                // Detachment Type
                formEN.addListItem()
                .setTitle("Detachment 2 Type")
                .setRequired(true)
                .setChoiceValues(DetachTypeArray);
                
                // FRENCH
                var Detach2FR = formFR.addPageBreakItem().setTitle("Détachement 2")
                // Detachment Name
                formFR.addTextItem()
                .setTitle("Nom du Détachement 2")
                .setRequired(true);
                // Detachment Type
                formFR.addListItem()
                .setTitle("Type du Détachment 2")
                .setRequired(true)
                .setChoiceValues(DetachTypeArray);
              }
              
              // DETACHMENT 3
              if(NbDetachMax >= 3){
                
                // ENGLISH
                var Detach3EN = formEN.addPageBreakItem().setTitle("Detachment 3");
                // Detachment Name
                formEN.addTextItem()
                .setTitle("Detachment 3 Name")
                .setRequired(true);
                // Detachment Type
                formEN.addListItem()
                .setTitle("Detachment 3 Type")
                .setRequired(true)
                .setChoiceValues(DetachTypeArray);
                
                // FRENCH
                var Detach3FR = formFR.addPageBreakItem().setTitle("Détachement 3")
                // Detachment Name
                formFR.addTextItem()
                .setTitle("Nom du Détachement 3")
                .setRequired(true);
                // Detachment Type
                formFR.addListItem()
                .setTitle("Type du Détachment 3")
                .setRequired(true)
                .setChoiceValues(DetachTypeArray);
              }
            }
            
            // Loop through each potential unit of each detachment
            for(var DetachNb = 1; DetachNb <= NbDetachMax; DetachNb++){
              // Selects the number of Units allowed in each Detachment
              if(DetachNb == 1) NbUnitMax = NbUnitDetach1;
              if(DetachNb == 2) NbUnitMax = NbUnitDetach2;
              if(DetachNb == 3) NbUnitMax = NbUnitDetach3;
              
              Logger.log('Current Detachment:%s',DetachNb);
              Logger.log('Units:%s',NbUnitMax);
              
              for(var UnitNb = 1; UnitNb <= NbUnitMax; UnitNb++){
                
                // UNIT SECTION
                // Set Index (for Form routing)
                UnitIndex = (DetachNb*100) + UnitNb;
                
                // ENGLISH
                // Title
                UnitTitle = "Detachment " + DetachNb + " - Unit " + UnitNb;
                // Set Unit Page
                UnitPageEN[UnitIndex] = formEN.addPageBreakItem().setTitle(UnitTitle);
                
                // FRENCH
                // Title
                UnitTitle = "Détachement " + DetachNb + " - Unité " + UnitNb;
                // Set Unit Page
                UnitPageFR[UnitIndex] = formFR.addPageBreakItem().setTitle(UnitTitle);
                Logger.log(UnitIndex);
                
                
                // UNIT PROFILE
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Profile")
                .setRequired(true);
                
                // FRENCH
                formFR.addTextItem()
                .setTitle("Détachement " + DetachNb + " - Unité " + UnitNb + " - Profil")
                .setRequired(true);
                
                
                // UNIT ROLE
                // ENGLISH
                formEN.addListItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Unit Role")
                .setRequired(true)
                .setChoiceValues(UnitRoleArray);
                
                // FRENCH
                formFR.addListItem()
                .setTitle("Détachement " + DetachNb + " - Unité " + UnitNb + " - Rôle d'Unité")
                .setRequired(true)
                .setChoiceValues(UnitRoleArray);
                
                
                // UNIT COMPOSITION
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Number of Models in Unit")
                .setRequired(true)
                .setValidation(ModelValidationEN);

                // FRENCH
                formFR.addTextItem()
                .setTitle("Détachement " + DetachNb + " - Unité " + UnitNb + " - Nombre de modèles dans l'unité")
                .setRequired(true)
                .setValidation(ModelValidationFR);
                
                
                // POWER LEVEL / POINTS
                // ENGLISH
                if(ArmyRating == 'Power Level') ArmyRatingText = " - Power Level";
                if(ArmyRating == 'Points')      ArmyRatingText = " - Total Points";
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + ArmyRatingText)
                .setRequired(true)
                .setValidation(RatingValidationEN);
                
                // FRENCH
                if(ArmyRating == 'Power Level') ArmyRatingText = " - Niveau Puissance";
                if(ArmyRating == 'Points')      ArmyRatingText = " - Total de Points";
                formFR.addTextItem()
                .setTitle("Détachement " + DetachNb + " - Unité " + UnitNb + ArmyRatingText)
                .setRequired(true)
                .setValidation(RatingValidationFR);                
                
                
                // CONTINUITY
                
                // Add Unit or Detachment
                // ENGLISH
                var AddUnitEN = formEN.addMultipleChoiceItem();
                AddUnitEN.setTitle("Add Unit or New Detachment");
                AddUnitEN.setRequired(true);
                
                // Create the different choices
                ChcUnitEN = AddUnitEN.createChoice("Add Unit",FormApp.PageNavigationType.CONTINUE);
                ChcEndEN  = AddUnitEN.createChoice("My Army List is Complete",FormApp.PageNavigationType.SUBMIT);
                
                // FRENCH
                var AddUnitFR = formFR.addMultipleChoiceItem();
                AddUnitFR.setTitle("Ajouter Unité ou Nouveau Détachement");
                AddUnitFR.setRequired(true);
                
                // Create the different choices
                ChcUnitFR = AddUnitFR.createChoice("Ajouter Unité",FormApp.PageNavigationType.CONTINUE);
                ChcEndFR  = AddUnitFR.createChoice("Ma liste d'armée est complète",FormApp.PageNavigationType.SUBMIT);
                
                
                // If Unit is First Detachment
                if(DetachNb == 1 && NbDetachMax > 1) {
                  ChcDetachEN = AddUnitEN.createChoice("Add New Detachment",Detach2EN);
                  ChcDetachFR = AddUnitFR.createChoice("Ajouter Nouveau Détachement",Detach2FR);
                }
                // If Unit is Second Detachment and there are 3 Detachments
                if(DetachNb == 2 && NbDetachMax > 2) {
                  ChcDetachEN = AddUnitEN.createChoice("Add New Detachment",Detach3EN);
                  ChcDetachFR = AddUnitFR.createChoice("Ajouter Nouveau Détachement",Detach3FR);
                }
                // Sets the Choices depending on the Unit and Detachment
                if(DetachNb < NbDetachMax){
                  if(UnitNb < NbUnitMax) {
                    AddUnitEN.setChoices([ChcUnitEN, ChcDetachEN, ChcEndEN]);
                    AddUnitFR.setChoices([ChcUnitFR, ChcDetachFR, ChcEndFR]);
                  }
                  if(UnitNb == NbUnitMax) {
                    AddUnitEN.setChoices([ChcDetachEN, ChcEndEN]);
                    AddUnitFR.setChoices([ChcDetachFR, ChcEndFR]);
                  }
                }
                
                if(DetachNb == NbDetachMax){
                  if(UnitNb < NbUnitMax) {
                    AddUnitEN.setChoices([ChcUnitEN, ChcEndEN]);
                    AddUnitFR.setChoices([ChcUnitFR, ChcEndFR]);
                  }
                  if(UnitNb == NbUnitMax) {
                    AddUnitEN.setChoices([ChcEndEN]);
                    AddUnitFR.setChoices([ChcEndFR]);
                  }
                }
                
                if (DetachNb == NbDetachMax && UnitNb == NbUnitMax) UnitNb = NbUnitMax + 1; 
              }
            }
            // Sets Go To Detachment 2 Unit 1 Page
            if(NbDetachMax == 2){
              // ENGLISH
              Detach2EN.setGoToPage(UnitPageEN[101]);
              UnitPageEN[101].setGoToPage(UnitPageEN[201]);
              // FRENCH
              Detach2FR.setGoToPage(UnitPageFR[101]);
              UnitPageFR[101].setGoToPage(UnitPageFR[201]);
            }
            
            // Sets Go To Detachment 3 Unit 1 Page   
            if(NbDetachMax == 3){
              // ENGLISH
              Detach2EN.setGoToPage(UnitPageEN[101]);
              Detach3EN.setGoToPage(UnitPageEN[201]);
              UnitPageEN[101].setGoToPage(UnitPageEN[301]);
              // FRENCH
              Detach2FR.setGoToPage(UnitPageFR[101]);
              Detach3FR.setGoToPage(UnitPageFR[201]);
              UnitPageFR[101].setGoToPage(UnitPageFR[301]);
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
      // English Form
      formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
      
      // Find and Rename Response Sheet
      ss = SpreadsheetApp.openById(ssID);
      ssSheets = ss.getSheets();
      ssSheets[0].setName('Registration EN');
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('Registration EN');
      ss.moveActiveSheet(17);
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
      ssSheets[0].setName('Registration FR');
      
      // Move Response Sheet to appropriate spot in file
      shtResp = ss.getSheetByName('Registration FR');
      ss.moveActiveSheet(18);
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