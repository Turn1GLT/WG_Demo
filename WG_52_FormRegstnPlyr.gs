// **********************************************
// function fcnCrtRegstnFormPlyr_WG()
//
// This function creates the Player Registration Form 
// based on the parameters in the Config File
//
// **********************************************

function fcnCrtRegstnFormPlyr_WG() {
  
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
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,20,3).getValues();
  var cfgArmyBuild =       shtConfig.getRange(4,33,16,1).getValues();
  
  // Execution Parameters
  var exeGnrtResp = cfgExecData[3][0];
  
  // Event Properties
  var evntLocation   = cfgEvntParam[0][0];
  var evntGameSystem = cfgEvntParam[5][0];
  var evntName       = cfgEvntParam[7][0];
  var evntFormat     = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  
  var armyBuildRatingVal = cfgArmyBuild[0][0];
  var armyBuildStartVal  = cfgArmyBuild[1][0];
    
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
  if(evntGameSystem == "Warhammer 40k"){
    var shtConfigWH40k = ss.getSheetByName('ConfigWH40k');
    var ArmyBuild40k = shtConfigWH40k.getRange(4,2,20,1).getValues();
    
    var NbFaction     = ArmyBuild40k[1][0];
    var NbDetachMax   = ArmyBuild40k[2][0];
    var NbUnitDetach1 = ArmyBuild40k[3][0];
    var NbUnitDetach2 = ArmyBuild40k[4][0];
    var NbUnitDetach3 = ArmyBuild40k[5][0];
    var UnitModelMin  = 1;
    var UnitModelMax  = ArmyBuild40k[6][0];
    var UnitRatingMin = 1;
    var UnitRatingMax = ArmyBuild40k[7][0];
    
    var DetachList = shtConfigWH40k.getRange(4,3,20,2).getValues();
    var DetachIncr = 0;
    var DetachTypeArray = new Array(12);
    
    var UnitRolesList = shtConfigWH40k.getRange(4,6,10,2).getValues();
    var UnitIncr = 0;
    var UnitRoleArray = new Array(9);
  }
  
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
  
  var shtResp1;
  var shtResp2;
  var shtRespName1;
  var shtRespName2;
  var IndexPlayers = ss.getSheetByName("Players").getIndex();
  var FormsCreated = 0;
  var FormsDeleted = 0;

  var ChcNbDetachArray;
  // Number of Detachment in Army
  if(NbDetachMax == 1) ChcNbDetachArray = ["1"];
  if(NbDetachMax == 2) ChcNbDetachArray = ["1","2"];
  if(NbDetachMax == 3) ChcNbDetachArray = ["1","2","3"];
  
  var ui;
  var title;
  var msg;
  var uiResponse;
  
  // If Event Format is not Single or Team+Players, Pop up Error Message
  if(evntFormat != "Single" && evntFormat != "Team+Players"){
    ui = SpreadsheetApp.getUi();
    title = "Registration Forms Error";
    msg = "The Event does not support Players Registration. Please review Event configuration";
    uiResponse = ui.alert(title, msg, ui.ButtonSet.OK);
  }
  
    // Checks if Event Format is Team or Team+Players
  if(evntFormat == "Single" || evntFormat == "Team+Players"){
    // If Form Exists, Log Error Message
    if(FormIdEN != '' || FormIdFR != ''){
      ErrorVal = 1;
      ui = SpreadsheetApp.getUi();
      title = "Players Registration Forms";
      msg = "The Registration Forms already exist. Click OK to overwrite.";
      uiResponse = ui.alert(title, msg, ui.ButtonSet.OK_CANCEL);
      
      if(uiResponse == "OK"){
        // Clear IDs and URLs
        shtConfig.getRange(rowFormEN, colFormID).clearContent();
        shtConfig.getRange(rowFormFR, colFormID).clearContent();
        shtConfig.getRange(rowFormEN, colFormURL).clearContent(); 
        shtConfig.getRange(rowFormFR, colFormURL).clearContent();
        
        // If Responses Sheets exist, Unlink and Delete them
        shtResp1 = ss.getSheets()[IndexPlayers];
        shtRespName1 = shtResp1.getName();
        shtResp2 = ss.getSheets()[IndexPlayers+1];
        shtRespName2 = shtResp2.getName();
        
        // First Sheet After Responses is MatchResp EN
        if(shtRespName1 == "RegPlyr EN"){
          FormApp.openById(FormIdEN).removeDestination();
          ss.deleteSheet(shtResp1);
        }
        
        // Second Sheet After Responses is MatchResp EN
        if(shtRespName2 == "RegPlyr EN"){
          FormApp.openById(FormIdEN).removeDestination();
          ss.deleteSheet(shtResp2);
        }
        
        // First Sheet After Responses is MatchResp EN
        if(shtRespName1 == "RegPlyr FR"){
          FormApp.openById(FormIdFR).removeDestination();
          ss.deleteSheet(shtResp1);
        }
        
        // Second Sheet After Responses is MatchResp FR
        if(shtRespName2 == "RegPlyr FR"){
          FormApp.openById(FormIdFR).removeDestination();
          ss.deleteSheet(shtResp2);
        }
        // Forms Deleted Flag
        FormsDeleted = 1;
      }
    }
    
    // Create Forms
    if ((FormIdEN == "" && FormIdFR == "") || FormsDeleted == 1){
      // Create Forms
      FormNameEN = evntLocation + " " + evntName + " Player Registration EN";
      formEN = FormApp.create(FormNameEN).setTitle(FormNameEN);
      
      FormNameFR = evntLocation + " " + evntName + " Player Registration FR";
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
            case 'First Name': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("First Name")
              .setHelpText("Please, Remove any space at the beginning or end of the name")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Prénom")
              .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
              .setRequired(true);
              break;
            }
              // LAST NAME
            case 'Last Name': {
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Last Name")
              .setHelpText("Please, Remove any space at the beginning or end of the name")
              .setRequired(true);
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Nom de Famille")
              .setHelpText("SVP, enlevez les espaces au début ou à la fin du nom")
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
              // TEAM NAME
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
              }
              break;
            }   
              // ARMY DEFINITION
            case 'Army Definition': {
              // ENGLISH
              formEN.addPageBreakItem()
              .setTitle("Army Definition")
              .setHelpText("Enter Army's General Information");
              
              // FRENCH
              formFR.addPageBreakItem()
              .setTitle("Définition d'Armée")
              .setHelpText("Entrez les informations générales de votre armée");
              
              break;
            }
              // FACTION KEYWORD 1
            case 'Faction Keyword 1' :{
              // Faction
              if (NbFaction == 1){
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Faction")
                .setHelpText("Enter your Main Faction Keyword")
                .setRequired(true);  
                
                // FRENCH
                formFR.addTextItem()
                .setTitle("Faction")
                .setHelpText("Entrez le mot-clé Faction principal de votre armée")
                .setRequired(true); 
              }
              
              if (NbFaction == 2){
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Faction 1")
                .setHelpText("Enter your First Faction Keyword")
                .setRequired(true);  
                
                // FRENCH
                formFR.addTextItem()
                .setTitle("Faction 1")
                .setHelpText("Entrez le premier mot-clé Faction de votre armée")
                .setRequired(true); 
              }
              break;
            }
              // FACTION KEYWORD 2
            case 'Faction Keyword 2' :{
              if (NbFaction == 2){
                // ENGLISH
                formEN.addTextItem()
                .setTitle("Faction 2")
                .setHelpText("Enter your Second Faction Keyword")
                .setRequired(true);
                
                // Faction Keyword 2
                formFR.addTextItem()
                .setTitle("Faction 2")
                .setHelpText("Entrez le deuxième mot-clé Faction de votre armée")
                .setRequired(true);
              }
              break;
            }
              // WARLORD NAME
            case 'Warlord' :{
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Army Warlord")
              .setHelpText("Enter your Army's Warlord name (or Unit Entry)")
              .setRequired(true); 
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Seigneur de Guerre de l'armée")
              .setHelpText("Entrez le nom (ou type d'unité) du Seigneur de Guerre de votre armée")
              .setRequired(true); 
              break;
            }
              // ARMY NAME
            case 'Army Name' :{ 
              // ENGLISH
              formEN.addTextItem()
              .setTitle("Army Name")
              .setHelpText("Enter your Army's Name (optional)")
              .setRequired(false); 
              
              // FRENCH
              formFR.addTextItem()
              .setTitle("Nom d'Armée")
              .setHelpText("Entrez le nom de votre Armée (optionel)")
              .setRequired(false);
              break;
            }
              // ARMY LIST
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
                  if(armyBuildRatingVal == 'Power Level') ArmyRatingText = " - Power Level";
                  if(armyBuildRatingVal == 'Points')      ArmyRatingText = " - Total Points";
                  formEN.addTextItem()
                  .setTitle("Detachment " + DetachNb + " - Unit " + UnitNb + ArmyRatingText)
                  .setRequired(true)
                  .setValidation(RatingValidationEN);
                  
                  // FRENCH
                  if(armyBuildRatingVal == 'Power Level') ArmyRatingText = " - Niveau Puissance";
                  if(armyBuildRatingVal == 'Points')      ArmyRatingText = " - Total de Points";
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
        // Forms Created Flag
        FormsCreated = 1;
      }
      
      // RESPONSE SHEETS
      // Create Response Sheet in Main File and Rename
      if(exeGnrtResp == "Enabled" && FormsCreated == 1){
        Logger.log("Generating Response Sheets and Form Links");
        var IndexPlayers = ss.getSheetByName("Players").getIndex();
        // English Form
        formEN.setDestination(FormApp.DestinationType.SPREADSHEET, ssID);
        
        // Find and Rename Response Sheet
        ss = SpreadsheetApp.openById(ssID);
        ssSheets = ss.getSheets();
        ssSheets[0].setName('RegPlyr EN');
        // Move Response Sheet to appropriate spot in file
        shtResp = ss.getSheetByName('RegPlyr EN');
        ss.moveActiveSheet(IndexPlayers+1);
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
        ssSheets[0].setName('RegPlyr FR');
        
        // Move Response Sheet to appropriate spot in file
        shtResp = ss.getSheetByName('RegPlyr FR');
        ss.moveActiveSheet(IndexPlayers+2);
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
        
        // Format Players Sheet
        // Hide Unused Columns
        // Loop through RespCol and Hide Matching Table Column if RespCol == "" 
        
        
      }
    }
  }
  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}