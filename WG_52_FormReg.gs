// **********************************************
// function fcnCreateRegForm_WG()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateRegForm_WG() {
  
  Logger.log("Routine: fcnCreateRegForm_WG");
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var ConfigData = shtConfig.getRange(56,2,32,1).getValues();
  var OptGenerateResp = ConfigData[8][0];
  var cfgTeamFormat = ConfigData[5][0];
  var cfgTeamMembers = ConfigData[13][0];
  
  var shtPlayers = ss.getSheetByName('Players');
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var ssID = shtIDs[0][0];
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
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
  var urlFormEN;
  
  var formFR;
  var FormIdFR;
  var FormNameFR;
  var FormItemsFR;
  var urlFormFR;
  
  var RowFormUrlEN = 23;
  var RowFormUrlFR = 24;
  var RowFormIdEN = 25;
  var RowFormIdFR = 26;
  
  var ErrorVal = '';
  var QuestionOrder = 1;
  
  // Response Columns from Configuration File
  // [x][0] = Response Columns
  var colRegRespValues = shtConfig.getRange(56,6,12,3).getValues();
  
  // Response Columns
  var colRespEmail = colRegRespValues[0][1];
  var colRespName = colRegRespValues[1][1];
  var colRespLanguage = colRegRespValues[2][1];
  var colRespPhone = colRegRespValues[3][1];
  var colRespTeamName = colRegRespValues[4][1];
  var colRespTeamMembers = colRegRespValues[5][1];
  var colRespArmyList = colRegRespValues[6][1];
  
  var ArmyComposition = shtConfig.getRange(7,11,9,1).getValues();
  var NbFaction     = ArmyComposition[1][0]
  var NbDetachMax   = ArmyComposition[2][0];
  var NbUnitDetach1 = ArmyComposition[3][0];
  var NbUnitDetach2 = ArmyComposition[4][0];
  var NbUnitDetach3 = ArmyComposition[5][0];
  
  var NbUnitMax;
  var UnitModelMin  = 1;
  var UnitModelMax  = ArmyComposition[6][0];
  var UnitRatingMin = 1;
  var UnitRatingMax = ArmyComposition[7][0];
  var ArmyRating = shtConfig.getRange(2,11).getValue(); 
  var ArmyRatingText;
  Logger.log('Army Rating: %s',ArmyRating);
  
  var Detachments = shtConfig.getRange(2,12,13,2).getValues();
  var DetachIncr = 0;
  var DetachTypeArray = new Array(12);
  
  var UnitRoles = shtConfig.getRange(2,15,10,2).getValues();
  var UnitIncr = 0;
  var UnitRoleArray = new Array(9);
  
  var ChUnitEN;
  var ChDetachEN;
  var ChEndEN;

  var ChUnitFR;
  var ChDetachFR;
  var ChEndFR;
  
  var UnitPageEN = new Array(325);
  var UnitPageFR = new Array(325);
  var UnitIndex;
  var UnitTitle;
  var UnitRole;
  var TestCol = 1;
  
  // Gets the Registration ID from the Config File
  FormIdEN = shtConfig.getRange(RowFormIdEN, 7).getValue();
  FormIdFR = shtConfig.getRange(RowFormIdFR, 7).getValue();
  
  // If Form Exists, Log Error Message
  if(FormIdEN != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdEN already exists. Unlink Response and Delete Form');
  }
  // If Form Exists, Log Error Message
  if(FormIdFR != ''){
    ErrorVal = 1;
    Logger.log('Error! FormIdFR already exists. Unlink Response and Delete Form');
  }
  
  // If Form does not exist, create it
  if(FormIdEN == '' && FormIdFR == ''){
    // Create Forms
    FormNameEN = shtConfig.getRange(3,2).getValue() + " Registration EN";
    formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
    
    FormNameFR = shtConfig.getRange(3,2).getValue() + " Registration FR";
    formFR = FormApp.create(FormNameFR).setTitle(FormNameFR);
    
    // Loops in Response Columns Values and Create Appropriate Question
    for(var i = 0; i < colRegRespValues.length; i++){
      // Look for Col Equal to Question Order
      if(QuestionOrder == colRegRespValues[i][1]){
        switch(colRegRespValues[i][0]){
          case 'Email': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]);
            // EMAIL
            // Set Registration Email collection
            formEN.setCollectEmail(true);
            formFR.setCollectEmail(true);
            break;
          }
          case 'Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // FULL NAME   
            formEN.addTextItem()
            .setTitle("Name")
            .setHelpText("Please, Remove any space at the end of the name")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Nom")
            .setHelpText("SVP, enlevez les espaces à la fin du nom")
            .setRequired(true);
            break;
          }
          case 'Language': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // LANGUAGE
            formEN.addMultipleChoiceItem()
            .setTitle("Language Preference")
            .setHelpText("Which Language do you prefer to use? The application is available in English and French")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            
            formFR.addMultipleChoiceItem()
            .setTitle("Préférence de Langue")
            .setHelpText("Quelle langue préférez-vous utiliser? L'application est disponible en anglais et en français.")
            .setRequired(true)
            .setChoiceValues(["English","Français"]);
            break;
          }
          case 'Phone Number': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // PHONE NUMBER    
            formEN.addTextItem()
            .setTitle("Phone Number")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Numéro de téléphone")
            .setRequired(true);
            break;
          }

          case 'Team Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // TEAM NAME
            formEN.addPageBreakItem().setTitle("Team");
            formEN.addTextItem()
            .setTitle("Team Name")
            .setRequired(true);
            
            formFR.addPageBreakItem().setTitle("Équipe");
            formFR.addTextItem()
            .setTitle("Nom d'équipe")
            .setRequired(true);

            // TEAM MEMBERS
            for(var member = 1; member <= cfgTeamMembers; member++){
              formEN.addTextItem()
              .setTitle("Teammate " + member)
              .setRequired(true);
              
              formFR.addTextItem()
              .setTitle("Équipier " + member)
              .setRequired(true);
            }
            break;
          }          
          case 'Army List': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // English
            // Army List
            formEN.addPageBreakItem()
            .setTitle("Army List");
            // Faction
            if (NbFaction == 1){
              // Faction Keyword 1
              formEN.addTextItem()
              .setTitle("Faction")
              .setRequired(true);  
            }
            if (NbFaction == 2){
              // Faction Keyword 1
              formEN.addTextItem()
              .setTitle("Faction 1")
              .setRequired(true);  
              
              // Faction Keyword 2
              formEN.addTextItem()
              .setTitle("Faction 2")
              .setRequired(true);
            }
            // Warlord name
            formEN.addTextItem()
            .setTitle("Warlord Name")
            .setRequired(true); 
            
            // Army name
            formEN.addTextItem()
            .setTitle("Army Name")
            .setRequired(false); 
            
            // French
            // Army List
            formFR.addPageBreakItem()
            .setTitle("Liste d'Armée");
            // Faction
            if (NbFaction == 1){
              // Faction Keyword 1
              formFR.addTextItem()
              .setTitle("Faction")
              .setRequired(true);  
            }
            if (NbFaction == 2){
              // Faction Keyword 1
              formFR.addTextItem()
              .setTitle("Faction 1")
              .setRequired(true);  
              
              // Faction Keyword 2
              formFR.addTextItem()
              .setTitle("Faction 2")
              .setRequired(true);
            }
            
            // Warlord name
            formFR.addTextItem()
            .setTitle("Nom du Seigneur de Guerre")
            .setRequired(true); 
            
            // Army name
            formFR.addTextItem()
            .setTitle("Nom d'Armée")
            .setRequired(false);
            
            // CREATE DETACHMENT CHOICES
            // Creates the List of Detachments Allowed for League
            for(var detach = 1; detach <= 12; detach++) {
              if(Detachments[detach][1] == 'Yes') {
                DetachTypeArray[DetachIncr] = Detachments[detach][0];
                DetachIncr++;
              }
            }
            DetachTypeArray.length = DetachIncr;
            
            // CREATE UNIT ROLES CHOICES
            // Creates the List of Unit Roles Allowed for League
            for(var unit = 1; unit <= 9; unit++) {
              if(UnitRoles[unit][1] == 'Yes') {
                UnitRoleArray[UnitIncr] = UnitRoles[unit][0];
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
                ChUnitEN = AddUnitEN.createChoice("Add Unit",FormApp.PageNavigationType.CONTINUE);
                ChEndEN  = AddUnitEN.createChoice("My Army List is Complete",FormApp.PageNavigationType.SUBMIT);
                
                // FRENCH
                var AddUnitFR = formFR.addMultipleChoiceItem();
                AddUnitFR.setTitle("Ajouter Unité ou Nouveau Détachement");
                AddUnitFR.setRequired(true);
                
                // Create the different choices
                ChUnitFR = AddUnitFR.createChoice("Ajouter Unité",FormApp.PageNavigationType.CONTINUE);
                ChEndFR  = AddUnitFR.createChoice("Ma liste d'armée est complète",FormApp.PageNavigationType.SUBMIT);
                
                
                // If Unit is First Detachment
                if(DetachNb == 1 && NbDetachMax > 1) {
                  ChDetachEN = AddUnitEN.createChoice("Add New Detachment",Detach2EN);
                  ChDetachFR = AddUnitFR.createChoice("Ajouter Nouveau Détachement",Detach2FR);
                }
                // If Unit is Second Detachment and there are 3 Detachments
                if(DetachNb == 2 && NbDetachMax > 2) {
                  ChDetachEN = AddUnitEN.createChoice("Add New Detachment",Detach3EN);
                  ChDetachFR = AddUnitFR.createChoice("Ajouter Nouveau Détachement",Detach3FR);
                }
                // Sets the Choices depending on the Unit and Detachment
                if(DetachNb < NbDetachMax){
                  if(UnitNb < NbUnitMax) {
                    AddUnitEN.setChoices([ChUnitEN, ChDetachEN, ChEndEN]);
                    AddUnitFR.setChoices([ChUnitFR, ChDetachFR, ChEndFR]);
                  }
                  if(UnitNb == NbUnitMax) {
                    AddUnitEN.setChoices([ChDetachEN, ChEndEN]);
                    AddUnitFR.setChoices([ChDetachFR, ChEndFR]);
                  }
                }
                
                if(DetachNb == NbDetachMax){
                  if(UnitNb < NbUnitMax) {
                    AddUnitEN.setChoices([ChUnitEN, ChEndEN]);
                    AddUnitFR.setChoices([ChUnitFR, ChEndFR]);
                  }
                  if(UnitNb == NbUnitMax) {
                    AddUnitEN.setChoices([ChEndEN]);
                    AddUnitFR.setChoices([ChEndFR]);
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
    if(OptGenerateResp == 'Enabled'){
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
      shtConfig.getRange(RowFormIdEN, 7).setValue(FormIdEN);
      FormIdFR = formFR.getId();
      shtConfig.getRange(RowFormIdFR, 7).setValue(FormIdFR);
      
      // Create Links to add to Config File  
      urlFormEN = formEN.getPublishedUrl();
      shtConfig.getRange(RowFormUrlEN, 2).setValue(urlFormEN); 
      
      urlFormFR = formFR.getPublishedUrl();
      shtConfig.getRange(RowFormUrlFR, 2).setValue(urlFormFR);
    }
  }
  // Post Log to Log Sheet
  subPostLog(shtLog);
}