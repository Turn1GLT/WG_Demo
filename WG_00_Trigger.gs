// **********************************************
// function fcnSubmitWG()
//
// This function analyzes the form submitted
// and executes the appropriate functions
//
// **********************************************

function onSubmitWG_Demo40K(e) {
      
  // Get Row from New Response
  var RowResponse = e.range.getRow();
    
  // Get Sheet from New Response
  var shtResponse = SpreadsheetApp.getActiveSheet();
  var ShtName = shtResponse.getSheetName();
  
  Logger.log('------- New Response Received -------');
  Logger.log('Sheet: %s',ShtName);
  Logger.log('Response Row: %s',RowResponse);
  
  // If Form Submitted is a Match Report, process results
  if(ShtName == 'Responses EN' || ShtName == 'Responses FR') {
    fcnProcessMatchWG();
  }
  
  // If Form Submitted is a Player Subscription, Register Player
  if(ShtName == 'Registration EN' || ShtName == 'Registration FR'){
    fcnRegistrationWG(shtResponse, RowResponse);
  }
  
  // If Form Submitted is a Weekly Booster, Enter Weekly Booster
  if(ShtName == 'WeekUnit EN' || ShtName == 'WeekUnit FR'){
    fcnWeekUnitWG(shtResponse, RowResponse);
  }
} 


// **********************************************
// function OnOpenWG()
//
// Function executed everytime the Spreadsheet is
// opened or refreshed
//
// **********************************************

function onOpenWG_Demo40K() {
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cfgEscalation = ss.getSheetByName('Config').getRange(68,2).getValue();
    
  var AnalyzeDataMenu  = [];
  AnalyzeDataMenu.push({name: 'Analyze New Match Entry', functionName: 'fcnMainWG_Grim40K'});
  AnalyzeDataMenu.push({name: 'Reset Match Entries', functionName:'fcnResetLeagueMatch'});
  
  var LeagueMenu = [];
  LeagueMenu.push({name:'Update Config ID & Links', functionName:'fcnUpdateLinksIDs'});
  LeagueMenu.push({name:'Create Match Report Forms', functionName:'fcnCreateReportForm_WG_S'});
  LeagueMenu.push({name:'Setup Response Sheets',functionName:'fcnSetupResponseSht'});
  LeagueMenu.push({name:'Create Registration Forms', functionName:'fcnCreateRegForm_WG_S'});
  if(cfgEscalation == 'Enabled') LeagueMenu.push({name:'Create Weekly Unit Forms', functionName:'fcnCreateWeekUnitForm_WG_S'});
  LeagueMenu.push({name:'Initialize League', functionName:'fcnInitLeague'});
  LeagueMenu.push(null);
  LeagueMenu.push({name:'Create Players Army DB', functionName:'fcnCrtPlayerArmyDB'});
  LeagueMenu.push({name:'Create Players Army Lists', functionName:'fcnCrtPlayerArmyList'});
  if(cfgEscalation == 'Enabled') LeagueMenu.push({name:'Create Players Weekly Units', functionName:'fcnCrtPlayerWeekUnit'});
  LeagueMenu.push(null);
  LeagueMenu.push({name:'Delete Players Army DB', functionName:'fcnDelPlayerArmyDB'});
  LeagueMenu.push({name:'Delete Players Army Lists', functionName:'fcnDelPlayerArmyList'});
  if(cfgEscalation == 'Enabled') LeagueMenu.push({name:'Delete Players Weekly Units', functionName:'fcnDelPlayerWeekUnit'});
  
  ss.addMenu("Manage League", LeagueMenu);
  ss.addMenu("Process Data", AnalyzeDataMenu);
}

// **********************************************
// function fcnWeekChangeWG()
//
// When the Week number changes, this function analyzes all
// generates a weekly report 
//
// **********************************************

function onWeekChangeWG_Demo40K(){

  // Main Spreadsheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Open Configuration Spreadsheet
  var shtConfig = ss.getSheetByName('Config');
  
  // League Name EN
  var Location = shtConfig.getRange(12,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(14,2).getValue();
  var LeagueNameEN = shtConfig.getRange(3,2).getValue() + ' ' + LeagueTypeEN;
  var LeagueNameFR = LeagueTypeFR + ' ' + shtConfig.getRange(3,2).getValue();
  
  // Open Cumulative Spreadsheet
  var shtCumul = ss.getSheetByName('Cumulative Results');
  var Week = shtCumul.getRange(2,3).getValue();
  var LastWeek = Week - 1;
  var WeekShtName = 'Week'+Week;
  var shtWeek = ss.getSheetByName(WeekShtName);
  var PenaltyTable;
  var EmailSubject;
  var EmailMessage;
  var MatchPlyd;
  var MatchesPlayed = 0;
  
  // Players Array to return Penalty Losses
  var PlayerData = new Array(32); // 0= Player Name, 1= Penalty Losses
  for(var plyr = 0; plyr < 32; plyr++){
    PlayerData[plyr] = new Array(2); 
    for (var val = 0; val < 2; val++) PlayerData[plyr][val] = '';
  }
  
  // Get Amount of matches played this week.
  MatchPlyd = shtWeek.getRange(5, 5, 32, 1).getValues();
  for(var plyr=0; plyr<32; plyr++){
    if(MatchPlyd[plyr][0] > 0) MatchesPlayed = MatchesPlayed + MatchPlyd[plyr][0];
  }

  // Analyze if Players have missing matches to apply Loss Penalties
  PlayerData = fcnAnalyzeLossPenalty(ss, Week, PlayerData);
  
  for(var row = 0; row<32; row++){
    if (PlayerData[row][0] != '') Logger.log('Player: %s - Missing: %s',PlayerData[row][0], PlayerData[row][1]);
  }
  
  // Populate the Penalty Table for the Weekly Report
  PenaltyTable = subEmailPlayerPenaltyTable(PlayerData);
  
  // Send Weekly Report Email
  EmailSubject = LeagueNameEN +' - Week ' + LastWeek + ' Report';
  EmailMessage = 'Week ' + LastWeek + ' is now complete and Week '+ Week +' has started. <br><br>Here is the week report for Week ' + LastWeek + '.<br><br>' +
    MatchesPlayed +' matches were played this week.<br>'+
      'etc etc etc...<br><br>';
  
  EmailMessage += PenaltyTable;
  
  MailApp.sendEmail('triadgaminglt@gmail.com', EmailSubject, EmailMessage,{name:'Triad Gaming Booster League Manager',htmlBody:EmailMessage});
  
  // Execute Ranking function in Standing tab
  fcnUpdateStandings(ss);
  
  // Copy all data to League Spreadsheet
  fcnCopyStandingsResults(ss, shtConfig);
  
}