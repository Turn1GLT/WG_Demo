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
  var shtLog = SpreadsheetApp.openById(shtIDs[1][0]).getSheetByName('Log');
  
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

  // Add Player to Player List
  PlayerData = fcnAddPlayerWG(shtIDs, shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData);
  var NbPlayers  = PlayerData[1];
  var PlayerName = PlayerData[2];
  
  // If Player was succesfully added, Generate Army DB, Generate Army List, Generate Startin Pool, Modify Match Report Form and Add Player to Round Booster
  if(PlayerData[0] == "New Player" && PlayerData[0] != "New Player" ) {
    fcnCrtPlayerArmyDB();
    Logger.log('Army Database Generated');
    fcnCrtPlayerArmyList();
    Logger.log('Army Lists Generated');  
    fcnCrtPlayerRoundUnit();
    Logger.log('Round Booster Generated');   
    fcnModifyReportFormTCG(ss, shtConfig, shtPlayers, shtIDs);
    
    // Execute Ranking function in Standing tab
    fcnUpdateStandings(ss, shtConfig);
    
    // Copy all data to Standing League Spreadsheet
    fcnCopyStandingsSheets(ss, shtConfig, cfgLgTrParam, 0, 1);
    
    // Send Confirmation to New Player
    fcnSendNewPlayerConf(shtConfig, PlayerData);
    Logger.log('Confirmation Email Sent');
    
    // Send Confirmation to Location
    // fcnSendNewPlayerConfLocation(shtConfig, PlayerData)
  }

  // Post Log to Log Sheet
  subPostLog(shtLog);
}




// **********************************************
// function fcnAddPlayerWG
//
// This function adds the new player to
// the Player's List
//
// **********************************************

function fcnAddPlayerWG(shtIDs, shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData) {

  // Opens Players List File
  var ssPlayersListID = shtConfig.getRange(40,2).getValue();
  var ssPlayersList = SpreadsheetApp.openById(shtIDs[10][0]);
  var shtPlayersList = ssPlayersList.getSheetByName('Players');
  
  // Current Player List
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var NextPlayerRow = NbPlayers + 3;
  var CurrPlayers = shtPlayers.getRange(2, 2, NbPlayers+1, 1).getValues();
  var Status = "New Player";
  
  var cfgTeamMembers = shtConfig.getRange(69,2).getValue();
  
  // Response Columns from Configuration File
  // [x][0] = Response Columns
  // [x][1] = Players Table Columns in Main Sheet
  var colRegRespValues = shtConfig.getRange(56,6,12,3).getValues();
  
  // Response Columns
  var colRespEmail = colRegRespValues[0][1];
  var colRespName = colRegRespValues[1][1];
  var colRespPhone = colRegRespValues[2][1];
  var colRespLanguage = colRegRespValues[3][1];
  var colRespTeamName = colRegRespValues[4][1];
  var colRespTeamMembers = colRegRespValues[5][1];
  var colRespArmyList = colRegRespValues[6][1];
  
  // Player Table Columns
  var colTableEmail = colRegRespValues[0][2] + 1;
  var colTableName = colRegRespValues[1][2] + 1;
  var colTablePhone = colRegRespValues[2][2] + 1;
  var colTableLanguage = colRegRespValues[3][2] + 1;
  var colTableTeamName = colRegRespValues[4][2] + 1;
  var colTableTeamMembers = colRegRespValues[4][2] + 1;
  var colTableArmyList = colRegRespValues[6][2] + 1;
  
  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var Responses = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  // Email
  var EmailAddress = Responses[0][colRespEmail];
  
  // Player Full Name
  var PlayerName = Responses[0][colRespName];
  
  // Player Language Preference
  var Language = Responses[0][colRespLanguage];
  
  // Player Phone Number
  if(colRespPhone != '' && colRespPhone != 0) {
    var Phone = Responses[0][colRespPhone];
  }
  
  // Team Name
  if(colRespTeamName != '' && colRespTeamName != 0) {
    var TeamName = Responses[0][colRespTeamName];
  }
  
  // Team Members
  if(colRespTeamMembers != '' && colRespTeamMembers != 0) {
    var TeamMember1 = Responses[0][colRespTeamMembers];
    if(cfgTeamMembers >= 2) var TeamMember2 = Responses[0][colRespTeamMembers+1];
    if(cfgTeamMembers >= 3) var TeamMember3 = Responses[0][colRespTeamMembers+2];
    if(cfgTeamMembers >= 4) var TeamMember4 = Responses[0][colRespTeamMembers+3];
  }
  
  // Player Army List
  if(colRespArmyList != '' && colRespArmyList != 0) {
    var ArmyList = Responses[0][colRespArmyList];
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
    shtPlayers.getRange(NextPlayerRow, colTableName).setValue(PlayerName);
    shtPlayersList.getRange(NextPlayerRow, colTableName).setValue(PlayerName);
    Logger.log('Player Name: %s',PlayerName);
    
    // Email Address
    shtPlayers.getRange(NextPlayerRow, colTableEmail).setValue(EmailAddress);
    shtPlayersList.getRange(NextPlayerRow, colTableEmail).setValue(EmailAddress);
    Logger.log('Email Address: %s',EmailAddress);
    
    // Language
    shtPlayers.getRange(NextPlayerRow, colTableLanguage).setValue(Language);
    shtPlayersList.getRange(NextPlayerRow, colTableLanguage).setValue(Language);
    Logger.log('Language: %s',Language);
    
    // Phone Number
    if(colTablePhone != '' && colTablePhone != 1){
      shtPlayers.getRange(NextPlayerRow, colTablePhone).setValue(Phone);
      shtPlayersList.getRange(NextPlayerRow, colTablePhone).setValue(Phone);
      Logger.log('Phone: %s',Phone);  
    }
	
    // Team Name
    if(colTableTeamName != '' && colTableTeamName != 1){
      shtPlayers.getRange(NextPlayerRow, colTableTeamName).setValue(TeamName);
      shtPlayersList.getRange(NextPlayerRow, colTableTeamName).setValue(TeamName);
      Logger.log('Team Name: %s',TeamName);  Logger.log('-----------------------------');
	}
    
    // Team Name
    if(colTableTeamMembers != '' && colTableTeamMembers != 1){
      shtPlayers.getRange(NextPlayerRow, colTableTeamMembers).setValue(TeamMember1);
      shtPlayersList.getRange(NextPlayerRow, colTableTeamMembers).setValue(TeamMember1);
      Logger.log('Team Name: %s',TeamName);  Logger.log('-----------------------------');
	}    
    
    // Army List
    if(colTableArmyList != '' && colTableArmyList != 1){
      // INSERT NEW FUNCTION TO PROCESS ARMY LIST INFORMATION
      // fcnProcessArmyList();
      shtPlayers.getRange(NextPlayerRow, colTableArmyList).setValue(ArmyList);
      shtPlayersList.getRange(NextPlayerRow, colTableArmyList).setValue(ArmyList);
      Logger.log('Army List: %s',ArmyList);  Logger.log('-----------------------------');
  }

  }
  PlayerData[0] = Status;
  PlayerData[1] = NbPlayers + 1;
  PlayerData[2] = PlayerName;
  PlayerData[3] = EmailAddress;
  PlayerData[4] = Language;
  PlayerData[5] = Phone;
  PlayerData[6] = TeamName;
  PlayerData[7] = ArmyList;
  
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

function fcnProcessArmyList(shtIDs, shtConfig, shtPlayers, shtResponse, RowResponse, PlayerData){
  
  // Get Response Sheet Name
  var RespSheetName = shtResponse.getSheetName();
  
  // Get All Values from Response Sheet
  var shtRespMaxCol = shtResponse.getMaxColumns();
  var RespHeader = shtResponse.getRange(1,1,1,shtRespMaxCol).getValues();
  var Responses  = shtResponse.getRange(RowResponse,1,1,shtRespMaxCol).getValues();
  
  Logger.log(RespHeader);
}

// **********************************************
// function fcnModifyReportFormTCG
//
// This function modifies the Match Report Form
// to add new added players
//
// **********************************************

function fcnModifyReportFormTCG(ss, shtConfig, shtPlayers, shtIDs) {

  var MatchFormEN = FormApp.openById(shtIDs[6][0]);
  var MatchFormItemEN = MatchFormEN.getItems();
  var MatchFormFR = FormApp.openById(shtIDs[7][0]);
  var MatchFormItemFR = MatchFormFR.getItems();
  var NbMatchFormItem = MatchFormItemFR.length;
  
  var RoundUnitFormEN = FormApp.openById(shtIDs[12][0]);
  var RoundUnitFormItemEN = RoundUnitFormEN.getItems();
  var RoundUnitFormFR = FormApp.openById(shtIDs[13][0]);
  var RoundUnitFormItemFR = RoundUnitFormFR.getItems();
  var NbRoundUnitFormItem = RoundUnitFormItemFR.length;

  // Function Variables
  var ItemTitle;
  var ItemPlayerListEN;
  var ItemPlayerListFR;
  var ItemPlayerChoice;
  
  var NbPlayers = shtPlayers.getRange(2, 1).getValue();
  var Players = shtPlayers.getRange(3, 2, NbPlayers, 1).getValues();
  var ListPlayers = [];
  
  // Loops in Match Form to Find Players List
  for(var item = 0; item < NbMatchFormItem; item++){
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
  
  // Loops in Round Unit Form to Find Players List
  for(var item = 0; item < NbRoundUnitFormItem; item++){
    ItemTitle = RoundUnitFormItemEN[item].getTitle();
    if(ItemTitle == 'Player'){
      
      // Get the List Item from the Round Booster Report Form
      ItemPlayerListEN = RoundUnitFormItemEN[item].asListItem();
      ItemPlayerListFR = RoundUnitFormItemFR[item].asListItem();
      
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