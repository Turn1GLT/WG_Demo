// **********************************************
// function fcnRegistrationWG
//
// This function adds the new player to
// the Player's List and calls other functions
// to create its complete profile
//
// **********************************************

function fcnRegistrationWG(shtResponse, RowResponse){

  Logger.log("Routine: fcnRegistrationWG");
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shtConfig = ss.getSheetByName('Config');
  var shtPlayers = ss.getSheetByName('Players');
  
  var shtIDs = shtConfig.getRange(4,7,20,1).getValues();
  var cfgEvntParam = shtConfig.getRange(4,4,32,1).getValues();
  var cfgColRspSht = shtConfig.getRange(4,18,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  var cfgExecData  = shtConfig.getRange(4,24,16,1).getValues();
  var cfgArmyBuild = shtConfig.getRange(4,33,20,1).getValues();
  
  // Log Sheet
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
  // Event Parameters
  var evntEscalation = cfgEvntParam[19][0];
  
  var PlayerData = new Array(10);
  PlayerData[0] = 0 ; // Function Status
  PlayerData[1] = ''; // Number of Players
  PlayerData[2] = ''; // New Player Full Name
  PlayerData[3] = ''; // New Player Email
  PlayerData[4] = ''; // New Player Language
  PlayerData[5] = ''; // New Player Phone Number
  PlayerData[6] = ''; // New Player Team Name
  PlayerData[7] = ''; // New Player Army List
  PlayerData[8] = ''; // New Player Spare
  PlayerData[9] = ''; // New Player Spare
  
  // Log new Registration
  Logger.log( '------- New Player Registration -------');

  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RegRspnVal = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  // Add Player to Player List
  PlayerData = fcnAddPlayerWG(shtIDs, shtConfig, shtPlayers, RegRspnVal, cfgEvntParam, PlayerData);
  var PlayerStatus = PlayerData[0];
  var NbPlayers  = PlayerData[1];
  var PlayerName = PlayerData[2];
  
  // If Player was succesfully added, Generate Army DB, Generate Army List, Generate Startin Pool, Modify Match Report Form and Add Player to Round Booster
  if(PlayerStatus == "New Player" && PlayerStatus != "New Player") {
    // Create Player Army DB
    fcnCrtPlayerArmyDB();
    Logger.log('Army Database Generated');
    
    // Process Player Army List to Army DB 
    fcnProcessArmyList(shtIDs, shtConfig, PlayerName, shtResponse, RowResponse, PlayerData);
    Logger.log('Army Data Processed to Army DB');
    
    // Create Player Army Lists (Player Access)
    fcnCrtPlayerArmyList();
    Logger.log('Army Lists Generated');  
    
    // If Escalation is Enabled, Create Player Escalation Bonus sheet 
    if(evntEscalation == 'Enabled'){
      fcnCrtPlayerEscltBonus();
      Logger.log('Round Unit Sheet Generated');   
    }
    // Add Player to Match Report Forms
    fcnModifyReportFormWG(shtIDs, shtPlayers, evntEscalation);
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings(ss, shtConfig);
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, cfgEvntParam, 0, 1);
    
    // Send Confirmation to New Player
    //fcnSendNewPlayerConf(shtConfig, PlayerData);
    //Logger.log('Confirmation Email Sent');
    
    // Send Confirmation to Location
    // fcnSendNewPlayerConfLocation(shtConfig, PlayerData)
  }

  // Post Log to Log Sheet
  subPostLog(shtLog, Logger.getLog());
}




// **********************************************
// function fcnAddPlayerWG
//
// This function adds the new player to
// the Player's List
//
// **********************************************

function fcnAddPlayerWG(shtIDs, shtConfig, shtPlayers, RegRspnVal, cfgEvntParam, PlayerData) {

  // Opens External Players List File
  var ssExtPlayers = SpreadsheetApp.openById(shtIDs[14][0]);
  var shtExtPlayers = ssExtPlayers.getSheetByName('Players');
  
  // Current Player List
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var NextPlayerRow = NbPlayers + 3;
  var CurrPlayers = shtPlayers.getRange(2, 2, NbPlayers+1, 1).getValues();
  var Status = "New Player";
  
  // Event Properties
  var evntFormat = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,20,3).getValues();
  
  // Response Columns
  var colRspEmail =        cfgRegFormCnstrVal[ 1][1];
  var colRspFullName =     cfgRegFormCnstrVal[ 2][1];
  var colRspFrstName =     cfgRegFormCnstrVal[ 3][1];
  var colRspLastName =     cfgRegFormCnstrVal[ 4][1];
  var colRspLanguage =     cfgRegFormCnstrVal[ 5][1];
  var colRspPhone =        cfgRegFormCnstrVal[ 6][1];
  var colRspTeamName =     cfgRegFormCnstrVal[ 7][1];
  var colRspTeamMembers =  cfgRegFormCnstrVal[ 8][1];
  var colRspArmyDef =      cfgRegFormCnstrVal[ 9][1];
  var colRspArmyName =     cfgRegFormCnstrVal[10][1];
  var colRspArmyFaction1 = cfgRegFormCnstrVal[11][1];
  var colRspArmyFaction2 = cfgRegFormCnstrVal[12][1];
  var colRspArmyWarlord =  cfgRegFormCnstrVal[13][1];
  
  // Player Table Columns
  var colTblEmail =        cfgRegFormCnstrVal[ 1][2];
  var colTblFullName =     cfgRegFormCnstrVal[ 2][2];
  var colTblFrstName =     cfgRegFormCnstrVal[ 3][2];
  var colTblLastName =     cfgRegFormCnstrVal[ 4][2];
  var colTblLanguage =     cfgRegFormCnstrVal[ 5][2];
  var colTblPhone =        cfgRegFormCnstrVal[ 6][2];
  var colTblTeamName =     cfgRegFormCnstrVal[ 7][2];
  var colTblTeamMembers =  cfgRegFormCnstrVal[ 8][2];
  var colTblArmyDef =      cfgRegFormCnstrVal[ 9][2];
  var colTblArmyName =     cfgRegFormCnstrVal[10][2];
  var colTblArmyFaction1 = cfgRegFormCnstrVal[11][2];
  var colTblArmyFaction2 = cfgRegFormCnstrVal[12][2];
  var colTblArmyWarlord =  cfgRegFormCnstrVal[13][2];
  
  var colTblContact = cfgRegFormCnstrVal[18][2];
  var colTblContactGrp = cfgRegFormCnstrVal[19][2];
  
  // Routine Variables
  var PlyrEmail = '';
  var PlyrFullName = '';
  var PlyrFrstName = '';
  var PlyrLastName = '';
  var PlyrLanguage = '';
  var PlyrPhone = '';
  var PlyrTeamName = '';
  var PlyrArmyDef = '';
  var PlyrTeamMember1 = '';
  var PlyrTeamMember2 = '';
  var PlyrTeamMember3 = '';
  var PlyrTeamMember4 = '';
  var PlyrContactInfo = new Array(4); // [0]= First Name, [1]= Last Name, [2]= Email Address, [3]= Language Preference
  
  // Email
  PlyrEmail = RegRspnVal[0][colRspEmail-1];
  
  // Player First and Last Name
  PlyrFrstName = RegRspnVal[0][colRspFrstName-1];
  PlyrLastName = RegRspnVal[0][colRspLastName-1];
  
  // Create Full Name
  PlyrFullName = PlyrFrstName + ' ' + PlyrLastName;
  
  // Player Language Preference
  PlyrLanguage = RegRspnVal[0][colRspLanguage-1];
  
  // Player Phone Number
  if(colRspPhone != '') PlyrPhone = RegRspnVal[0][colRspPhone-1];
  
  // Team Name
  if(colRspTeamName != '') PlyrTeamName = RegRspnVal[0][colRspTeamName-1];
  
  // Team Members
  if(colRspTeamMembers != '') {
    PlyrTeamMember1 = RegRspnVal[0][colRspTeamMembers-1];
    if(evntNbPlyrTeam >= 2) PlyrTeamMember2 = RegRspnVal[0][colRspTeamMembers+1-1];
    if(evntNbPlyrTeam >= 3) PlyrTeamMember3 = RegRspnVal[0][colRspTeamMembers+2-1];
    if(evntNbPlyrTeam >= 4) PlyrTeamMember4 = RegRspnVal[0][colRspTeamMembers+3-1];
  }
  
  // Player Army Definition
  if(colRspArmyDef != '') {
    PlyrArmyDef = RegRspnVal[0][colRspArmyDef-1];
    Logger.log("PlyrArmyDef: %s",PlyrArmyDef);
  }
  
  // Check if Player exists in List
  for(var i = 1; i <= NbPlayers; i++){
    if(PlyrFullName == CurrPlayers[i][0]){
      Status = "Cannot complete registration for " + PlyrFullName + ", Duplicate Player Found in List";
      Logger.log(Status)
    }
  }
  // If New Player
  // Copy Values to Players Sheet at the Next Empty Spot (Number of Players + 3)
  // Copy Values to Players List for Store Access
  if(Status == "New Player"){
    
    // Player Full Name
    shtPlayers.getRange(NextPlayerRow, colTblFullName).setValue(PlyrFullName);
    shtExtPlayers.getRange(NextPlayerRow, colTblFullName).setValue(PlyrFullName);
    Logger.log('Player Name: %s',PlyrFullName);
    
    // Email Address
    shtPlayers.getRange(NextPlayerRow, colTblEmail).setValue(PlyrEmail);
    shtExtPlayers.getRange(NextPlayerRow, colTblEmail).setValue(PlyrEmail);
    Logger.log('Email Address: %s',PlyrEmail);
    
    // Language
    shtPlayers.getRange(NextPlayerRow, colTblLanguage).setValue(PlyrLanguage);
    shtExtPlayers.getRange(NextPlayerRow, colTblLanguage).setValue(PlyrLanguage);
    Logger.log('Language: %s',PlyrLanguage);
    
    // Phone Number
    if(PlyrPhone != ''){
      shtPlayers.getRange(NextPlayerRow, colTblPhone).setValue(PlyrPhone);
      shtExtPlayers.getRange(NextPlayerRow, colTblPhone).setValue(PlyrPhone);
      Logger.log('Phone: %s',PlyrPhone);  
    }
	
    // Team Name
    if(PlyrTeamName != ''){
      shtPlayers.getRange(NextPlayerRow, colTblTeamName).setValue(PlyrTeamName);
      shtExtPlayers.getRange(NextPlayerRow, colTblTeamName).setValue(PlyrTeamName);
      Logger.log('Team Name: %s',PlyrTeamName);  Logger.log('-----------------------------');
	}
    
    // Army Definition
    if(PlyrArmyDef != ""){
      // INSERT NEW FUNCTION TO PROCESS ARMY LIST INFORMATION
      // fcnProcessArmyList();
      shtPlayers.getRange(NextPlayerRow, colTblArmyDef).setValue(PlyrArmyDef);
      shtExtPlayers.getRange(NextPlayerRow, colTblArmyDef.setValue(PlyrArmyDef));
      Logger.log('Army List: %s',PlyrArmyDef);  Logger.log('-----------------------------');
    }

    // Team Members
    if(PlyrTeamMember1 != ''){
      shtPlayers.getRange(NextPlayerRow, colTblTeamMembers).setValue(PlyrTeamMember1);
      shtExtPlayers.getRange(NextPlayerRow, colTblTeamMembers).setValue(PlyrTeamMember1);
      Logger.log('Team Name: %s',PlyrTeamMember1);  Logger.log('-----------------------------');
	}    
  
    // Set Player Contact Info 
    PlyrContactInfo[0]= PlyrFrstName;
    PlyrContactInfo[1]= PlyrLastName;
    PlyrContactInfo[2]= PlyrEmail;
    PlyrContactInfo[3]= PlyrLanguage;
    
    // Add Player Info to Contact List and Contact Group
    var PlyrCntctStatus = subCrtPlayerContact(PlyrContactInfo);
    if(PlyrCntctStatus == 'Player Contact Created' || PlyrCntctStatus == 'Player Contact Updated') {
      // Set Contact Created in Players Sheet
      shtPlayers.getRange(NextPlayerRow, colTblContact).setValue('X');
      shtExtPlayers.getRange(NextPlayerRow, colTblContact).setValue('X');
      // Add to Contact Group   
      var CntctGrpStatus = subAddPlayerContactGroup(shtConfig, PlyrContactInfo);
      if(CntctGrpStatus == 'Player added to Contact Group') {
        // Set Added in Contact Group in Players Sheet
        shtPlayers.getRange(NextPlayerRow, colTblContactGrp).setValue('X');
        shtExtPlayers.getRange(NextPlayerRow, colTblContactGrp).setValue('X');
      }
      else Logger.log('Player Added to Contact Group');
    }
    else Logger.log('Player Contact NOT Created');
 
  }
  
  PlayerData[0] = Status;
  PlayerData[1] = NbPlayers + 1;
  PlayerData[2] = PlyrFullName;
  PlayerData[3] = PlyrEmail;
  PlayerData[4] = PlyrLanguage;
  PlayerData[5] = PlyrPhone;
  PlayerData[6] = PlyrTeamName;
  PlayerData[7] = PlyrArmyDef;
  
  return PlayerData;
}


// **********************************************
// function fcnProcessArmyList
//
// This function processes the Army List Info
// from the Form Response and puts it in
// the player Army List DB
//
// **********************************************

function fcnProcessArmyList(shtIDs, shtConfig, shtPlayers, shtResponse, RegRspnVal, PlayerData){
  
  // Get Response Sheet Name
  var RespSheetName = shtResponse.getSheetName();
  
  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RespHeader = shtResponse.getRange(1,1,1,shtRespMaxCol).getValues();
  
  Logger.log(RespHeader);
}

// **********************************************
// function fcnModifyReportFormWG
//
// This function modifies the Match Report Form
// to add new added players
//
// **********************************************

function fcnModifyReportFormWG(shtIDs, shtPlayers, evntEscalation) {

  var MatchFormEN = FormApp.openById(shtIDs[6][0]);
  var MatchFormItemEN = MatchFormEN.getItems();
  var MatchFormFR = FormApp.openById(shtIDs[7][0]);
  var MatchFormItemFR = MatchFormFR.getItems();
  var MatchFormNbItem = MatchFormItemEN.length;
  
  var EscltBonusFormEN = FormApp.openById(shtIDs[12][0]);
  var EscltBonusFormItemEN = EscltBonusFormEN.getItems();
  var EscltBonusFormFR = FormApp.openById(shtIDs[13][0]);
  var EscltBonusFormItemFR = EscltBonusFormFR.getItems();
  var EscltBonusFormNbItem = EscltBonusFormItemEN.length;

  // Function Variables
  var ItemTitle;
  var ItemPlayerListEN;
  var ItemPlayerListFR;
  var ItemPlayerChoice;
  
  var NbPlayers = shtPlayers.getRange(2, 1).getValue();
  var Players = shtPlayers.getRange(3, 2, NbPlayers, 1).getValues();
  var ListPlayers = [];
  
  // Loops in Match Form to Find Players List
  for(var item = 0; item < MatchFormNbItem; item++){
    ItemTitle = MatchFormItemEN[item].getTitle();
    if(ItemTitle == 'Winning Player' || ItemTitle == 'Losing Player'){
      
      // Get the List Item from the Match Report Form
      ItemPlayerListEN = MatchFormItemEN[item].asListItem();
      ItemPlayerListFR = MatchFormItemFR[item].asListItem();
      
      // Build the Player List from the Players Sheet     
      for (i = 0; i < NbPlayers; i++){
        ListPlayers[i] = Players[i][0];
      }
      // Set the Player List to the Match Report Forms
      ItemPlayerListEN.setChoiceValues(ListPlayers);
      ItemPlayerListFR.setChoiceValues(ListPlayers);
    }
  }
  
  if(evntEscalation == 'Enabled'){
    // Loops in Escalation Bonus Form to Find Players List
    for(var item = 0; item < EscltBonusFormNbItem; item++){
      ItemTitle = EscltBonusFormNbItem[item].getTitle();
      if(ItemTitle == 'Player'){
        
        // Get the List Item from the Round Booster Report Form
        ItemPlayerListEN = EscltBonusFormItemEN[item].asListItem();
        ItemPlayerListFR = EscltBonusFormItemFR[item].asListItem();
        
        // Build the Player List from the Players Sheet     
        for (i = 0; i < NbPlayers; i++){
          ListPlayers[i] = Players[i][0];
        }
        // Set the Player List to the Round Booster Report Forms
        ItemPlayerListEN.setChoiceValues(ListPlayers);
        ItemPlayerListFR.setChoiceValues(ListPlayers);
      }
    }
  }
}