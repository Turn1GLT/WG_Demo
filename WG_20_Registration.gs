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
  if(PlayerStatus == "New Player") {
    // Add Player Info to Contact List and Contact Group
    subCrtPlayerContact(shtConfig, PlayerData);
    subAddPlayerContactGroup(shtConfig, PlayerData);
    Logger.log('Player Contact Generated');
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

  // Opens Players List File
  var ssExtPlayers = SpreadsheetApp.openById(shtIDs[14][0]);
  var shtExtPlayers = ssExtPlayers.getSheetByName('Players');
  
  // Current Player List
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var NextPlayerRow = NbPlayers + 3;
  var CurrPlayers = shtPlayers.getRange(2, 2, NbPlayers+1, 1).getValues();
  var Status = "New Player";
  
  // League Parameters
  var evntFormat = cfgEvntParam[9][0];
  var evntNbPlyrTeam = cfgEvntParam[10][0];
  
  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,16,3).getValues();
  
  // Response Columns
  var colRspEmail = cfgRegFormCnstrVal[1][1] - 1;
  var colRspName = cfgRegFormCnstrVal[2][1] - 1;
  var colRspLanguage = cfgRegFormCnstrVal[3][1] - 1;
  var colRspPhone = cfgRegFormCnstrVal[4][1] - 1;
  var colRspTeamName = cfgRegFormCnstrVal[5][1] - 1;
  var colRspTeamMembers = cfgRegFormCnstrVal[6][1] - 1;
  var colRspArmyDef = cfgRegFormCnstrVal[7][1] - 1;
  var colRspArmyName = cfgRegFormCnstrVal[8][1] - 1;
  var colRspArmyFaction1 = cfgRegFormCnstrVal[9][1] - 1;
  var colRspArmyFaction2 = cfgRegFormCnstrVal[10][1] - 1;
  var colRspArmyWarlord = cfgRegFormCnstrVal[11][1] - 1;
  
  // Player Table Columns
  var colTblEmail = cfgRegFormCnstrVal[1][2];
  var colTblName = cfgRegFormCnstrVal[2][2];
  var colTblLanguage = cfgRegFormCnstrVal[3][2];
  var colTblPhone = cfgRegFormCnstrVal[4][2];
  var colTblTeamName = cfgRegFormCnstrVal[5][2];
  var colTblTeamMembers = cfgRegFormCnstrVal[6][2];
  var colTblArmyDef = cfgRegFormCnstrVal[7][2];
  var colTblArmyName = cfgRegFormCnstrVal[8][2];
  var colTblArmyFaction1 = cfgRegFormCnstrVal[9][2];
  var colTblArmyFaction2 = cfgRegFormCnstrVal[10][2];
  var colTblArmyWarlord = cfgRegFormCnstrVal[11][2];
  
  // Email
  var EmailAddress = RegRspnVal[0][colRspEmail];
  
  // Player Full Name
  var PlayerName = RegRspnVal[0][colRspName];
  
  // Player Language Preference
  var Language = RegRspnVal[0][colRspLanguage];
  
  // Player Phone Number
  if(colRspPhone != '' && colRspPhone != -1) {
    var Phone = RegRspnVal[0][colRspPhone];
  }
  
  // Team Name
  if(colRspTeamName != '' && colRspTeamName != -1) {
    var TeamName = RegRspnVal[0][colRspTeamName];
  }
  
  // Team Members
  if(colRspTeamMembers != '' && colRspTeamMembers != -1) {
    var TeamMember1 = RegRspnVal[0][colRspTeamMembers];
    if(cfgTeamMembers >= 2) var TeamMember2 = RegRspnVal[0][colRspTeamMembers+1];
    if(cfgTeamMembers >= 3) var TeamMember3 = RegRspnVal[0][colRspTeamMembers+2];
    if(cfgTeamMembers >= 4) var TeamMember4 = RegRspnVal[0][colRspTeamMembers+3];
  }
  
  // Player Army List
  if(colRspArmyDef != '' && colRspArmyDef != -1) {
    var ArmyDef = RegRspnVal[0][colRspArmyDef];
  }
  
  // Check if Player exists in List
  for(var i = 1; i <= NbPlayers; i++){
    if(PlayerName == CurrPlayers[i][0]){
      Status = "Cannot complete registration for " + PlayerName + ", Duplicate Player Found in List";
      Logger.log(Status)
    }
  }

  // Copy Values to Players Sheet at the Next Empty Spot (Number of Players + 3)
  // Copy Values to Players List for Store Access
  if(Status == "New Player"){
	// Name
    shtPlayers.getRange(NextPlayerRow, colTblName).setValue(PlayerName);
    shtExtPlayers.getRange(NextPlayerRow, colTblName).setValue(PlayerName);
    Logger.log('Player Name: %s',PlayerName);
    
    // Email Address
    shtPlayers.getRange(NextPlayerRow, colTblEmail).setValue(EmailAddress);
    shtExtPlayers.getRange(NextPlayerRow, colTblEmail).setValue(EmailAddress);
    Logger.log('Email Address: %s',EmailAddress);
    
    // Language
    shtPlayers.getRange(NextPlayerRow, colTblLanguage).setValue(Language);
    shtExtPlayers.getRange(NextPlayerRow, colTblLanguage).setValue(Language);
    Logger.log('Language: %s',Language);
    
    // Phone Number
    if(colTblPhone != '' && colTblPhone != 1){
      shtPlayers.getRange(NextPlayerRow, colTblPhone).setValue(Phone);
      shtExtPlayers.getRange(NextPlayerRow, colTblPhone).setValue(Phone);
      Logger.log('Phone: %s',Phone);  
    }
	
    // Team Name
    if(colTblTeamName != '' && colTblTeamName != 1){
      shtPlayers.getRange(NextPlayerRow, colTblTeamName).setValue(TeamName);
      shtExtPlayers.getRange(NextPlayerRow, colTblTeamName).setValue(TeamName);
      Logger.log('Team Name: %s',TeamName);  Logger.log('-----------------------------');
	}
    
    // Team Name
    if(colTblTeamMembers != '' && colTblTeamMembers != 1){
      shtPlayers.getRange(NextPlayerRow, colTblTeamMembers).setValue(TeamMember1);
      shtExtPlayers.getRange(NextPlayerRow, colTblTeamMembers).setValue(TeamMember1);
      Logger.log('Team Name: %s',TeamName);  Logger.log('-----------------------------');
	}    
    
    // Army List
    if(colTblArmyDef != '' && colTblArmyDef != 1){
      // INSERT NEW FUNCTION TO PROCESS ARMY LIST INFORMATION
      // fcnProcessArmyList();
      shtPlayers.getRange(NextPlayerRow, colTblArmyDef).setValue(ArmyDef);
      shtExtPlayers.getRange(NextPlayerRow, colTblArmyDef.setValue(ArmyDef));
      Logger.log('Army List: %s',ArmyDef);  Logger.log('-----------------------------');
    }

  }
  PlayerData[0] = Status;
  PlayerData[1] = NbPlayers + 1;
  PlayerData[2] = PlayerName;
  PlayerData[3] = EmailAddress;
  PlayerData[4] = Language;
  PlayerData[5] = Phone;
  PlayerData[6] = TeamName;
  PlayerData[7] = ArmyDef;
  
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