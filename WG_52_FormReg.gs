// **********************************************
// function fcnCreateRegForm()
//
// This function creates the Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCreateRegForm() {
  
  var ss = SpreadsheetApp.getActive();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  var shtIDs = shtConfig.getRange(17,7,20,1).getValues();
  var ssID = shtIDs[0][0]; 
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
  var colRegRespValues = shtConfig.getRange(56,6,12,2).getValues();
  
  // Response Columns
  var colRespEmail = colRegRespValues[0][1];
  var colRespName = colRegRespValues[1][1];
  var colRespFirstName = colRegRespValues[2][1];
  var colRespLastName = colRegRespValues[3][1];
  var colRespPhone = colRegRespValues[4][1];
  var colRespLanguage = colRegRespValues[5][1];
  var colRespTeamName = colRegRespValues[6][1];
  var colRespDCI = colRegRespValues[7][1];
  
  var NbDetachMax   = shtConfig.getRange( 8,11).getValue();
  var NbUnitDetach1 = shtConfig.getRange( 9,11).getValue();
  var NbUnitDetach2 = shtConfig.getRange(10,11).getValue();
  var NbUnitDetach3 = shtConfig.getRange(11,11).getValue();
  var NbUnitMax;
  
  var Detachments = shtConfig.getRange(12,10,13,2).getValues();
  var DetachIncr = 0;
  var DetachTypeArray = new Array(12);
  
  var UnitRoles = shtConfig.getRange(25,10,10,2).getValues();
  var UnitIncr = 0;
  var UnitRoleArray = new Array(9);
  
  var ChUnitEN;
  var ChDetachEN;
  var ChEndEN;

  var ChUnitFR;
  var ChDetachFR;
  var ChEndFR;
  
  var UnitPage = new Array(325);
  var UnitIndex;
  var UnitTitleEN;
  var UnitTitleFR;
  var UnitRoleEN;
  var UnitRoleFR;
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
          case 'Team Name': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // TEAM NAME
            formEN.addTextItem()
            .setTitle("Team Name")
            .setRequired(true);
            
            formFR.addTextItem()
            .setTitle("Nom d'équipe")
            .setRequired(true);
            break;
          }
          case 'Army List': {
            Logger.log('%s - %s',QuestionOrder,colRegRespValues[i][0]); 
            // Army List
            formEN.addPageBreakItem()
            .setTitle("Army List");
//            .setDescription("Please fill up the following to submit your Army List");

            // Faction Keyword 1
            formEN.addTextItem()
            .setTitle("Faction Keyword 1")
            .setRequired(true);  
            
            // Faction Keyword 2
            formEN.addTextItem()
            .setTitle("Faction Keyword 2")
            .setRequired(true);
            
            // Warlord name
            formEN.addTextItem()
            .setTitle("Warlord Name")
            .setRequired(true); 
            
            // Army name
            formEN.addTextItem()
            .setTitle("Army Name")
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
              
              // Number of Models in Unit
              var ModelValidation = FormApp.createTextValidation()
              .setHelpText("Enter a number between 1 and 100.")
              .requireNumberBetween(1, 100)
              .build();
              
              // Power Level of Unit
              var LevelValidation = FormApp.createTextValidation()
              .setHelpText("Enter a number between 1 and 100.")
              .requireNumberBetween(1, 100)
              .build();
              
              for(var UnitNb = 1; UnitNb <= NbUnitMax; UnitNb++){
                
                // Creates the Unit Section
                // Set Index
                UnitIndex = (DetachNb*100) + UnitNb;
                // Title
                UnitTitleEN = "Detachment " + DetachNb + " - Unit " + UnitNb;
                // Set Unit Page
                UnitPage[UnitIndex] = formEN.addPageBreakItem().setTitle(UnitTitleEN);
                Logger.log(UnitIndex);
                // Unit Title
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Unit Title")
                .setRequired(true);
                
                // Unit Role
                UnitRoleEN = formEN.addListItem();
                UnitRoleEN.setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Unit Role")
                UnitRoleEN.setRequired(true)
                UnitRoleEN.setChoiceValues(UnitRoleArray);
                
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Number of Models in Unit")
                .setRequired(true)
                .setValidation(ModelValidation);
                
                formEN.addTextItem()
                .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + " - Unit Power Level")
                .setRequired(true)
                .setValidation(LevelValidation);
                
                // Add Unit or Detachment 
                var AddUnitEN = formEN.addMultipleChoiceItem();
                AddUnitEN.setTitle("Add Another Unit or Another Detachment");
                AddUnitEN.setRequired(true);
                
                // Create the different choices
                ChUnitEN = AddUnitEN.createChoice("Add Another Unit",FormApp.PageNavigationType.CONTINUE);
                ChEndEN = AddUnitEN.createChoice("My Army List is Complete",FormApp.PageNavigationType.SUBMIT);
                
                // If Unit is First Detachment
                if(DetachNb == 1 && NbDetachMax > 1) ChDetachEN = AddUnitEN.createChoice("Add Another Detachment",Detach2EN);
                
                // If Unit is Second Detachment and there are 3 Detachments
                if(DetachNb == 2 && NbDetachMax > 2) ChDetachEN = AddUnitEN.createChoice("Add Another Detachment",Detach3EN);
                
                // Sets the Choices depending on the Unit and Detachment
                if(DetachNb < NbDetachMax){
                  if(UnitNb < NbUnitMax) AddUnitEN.setChoices([ChUnitEN, ChDetachEN, ChEndEN]);
                  if(UnitNb == NbUnitMax) AddUnitEN.setChoices([ChDetachEN, ChEndEN]);
                }
                
                if(DetachNb == NbDetachMax){
                  if(UnitNb < NbUnitMax) AddUnitEN.setChoices([ChUnitEN, ChEndEN]);
                  if(UnitNb == NbUnitMax) AddUnitEN.setChoices([ChEndEN]);
                }
                
                if (DetachNb == NbDetachMax && UnitNb == NbUnitMax) UnitNb = NbUnitMax + 1; 
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