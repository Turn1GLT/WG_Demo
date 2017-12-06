// **********************************************
// function subLogger()
//
// This function posts the log to a spreadsheet
//
// **********************************************

function subPostLog(shtLog, Log) {
  shtLog.insertRowBefore(2);
  shtLog.getRange(2,1).setValue(new Date()).setNumberFormat('yyyy-MM-dd / HH:mm:ss');
  shtLog.getRange(2,2).setValue(Log);
}


// **********************************************
// function subCreateArray(X,Y)
//
// This function creates and array of two dimensions X-Y
// First
//
// **********************************************

function subCreateArray(X,Y){
  
  var newArray;
  
  // If dimension X is greater than zero
  if(X > 0){
    // Create Array of dimension X
    newArray = new Array(X)
    // Loops in dimension X to create dimension Y
    for(var x = 0; x < X; x++){
      // If a dimension Y is greater than zero, dimension Y exists
      if(Y > 0){
        newArray[x] = new Array(Y);
        for (var y = 0; y < Y; y++) newArray[x][y] = '';
      }
      // If not, Array is one dimension X
      else{
        newArray[x] = '';
      }
    }
  }
  else{
    newArray = '';
  }
  return newArray;
} 

// **********************************************
// function SubDelPlayerSheets()
//
// This function deletes all Players Sheets 
// from the Parameter 
//
// **********************************************

function subDelPlayerSheets(shtID){
  
  Logger.log("Routine: fcnDelPlayerSheets");
  
  // Spreadsheet
  var ssDel = SpreadsheetApp.openById(shtID); 
  var shtTemplate = ssDel.getSheetByName('Template');
  var sheets = ssDel.getSheets();
  var NbSheet = ssDel.getNumSheets();
  
  // Routine Variables
  var shtCurr;
  var shtCurrName;
  
  // Activates Template Sheet
  ssDel.setActiveSheet(shtTemplate);
  
  // Loop through the Spreadsheet to delete Sheets
  for (var sht = 0; sht < NbSheet - 1; sht ++){
    // Get Sheet Name
    shtCurrName = sheets[0].getSheetName();
    // If First Sheet is Template
    if(shtCurrName != 'Template') {
      // Delete Sheet
      ssDel.deleteSheet(sheets[0]);
      // Update Sheets
      sheets = ssDel.getSheets()
      NbSheet--;
      sht--;
    }
    // If First Sheet is Not Template
    if(shtCurrName == 'Template' && NbSheet > 1) {
      // Delete Sheet
      ssDel.deleteSheet(sheets[1]);
      // Update Sheets
      sheets = ssDel.getSheets()
      NbSheet--;
      sht--;
    }
  }
}



// **********************************************
// function subFindPlayerRow()
//
// This function finds the Player Row in the Spreadsheet
//
// **********************************************

function subFindPlayerRow(sheet, rowStart, colPlyr, length, PlayerName){
  Logger.log("Routine: subFindPlayerRow: %s",PlayerName);
  var RsltRow = 0;
  
  var RndPlyrList = sheet.getRange(rowStart,colPlyr,length,1).getValues();
  
  // Find the Winning and Losing Player in the Round Result Tab
  for (var row = rowStart; row < length+rowStart; row ++){
    if (RndPlyrList[row - 5][0] == PlayerName) {
      RsltRow = row;
      row = 37;
    }
  }
  Logger.log("Player Row Found for %s : %s",PlayerName,RsltRow);
  return RsltRow;
}

// **********************************************
// function subGetEmailAddressSngl()
//
// This function gets the email addresses for a
// single player from the configuration file
//
// **********************************************

function subGetEmailAddressSngl(Player, shtPlayers, EmailAddresses){
  
  // Players Sheets for Email addresses
  var colEmail = 3;
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var rowPlayer = 0;
  var PlyrRowStart = 3;
  
  var PlayerNames = shtPlayers.getRange(PlyrRowStart,2,NbPlayers,1).getValues();
  
  // Find Players rows
  for (var row = 0; row < NbPlayers; row++){
    if (PlayerNames[row] == Player) rowPlayer = row + PlyrRowStart;
    if (rowPlayer > 0) row = NbPlayers + 1;
  }
  
  // Get Email addresses using the players rows
  EmailAddresses[0] = shtPlayers.getRange(rowPlayer,colEmail+1).getValue();
  EmailAddresses[1] = shtPlayers.getRange(rowPlayer,colEmail).getValue();
    
  return EmailAddresses;
}

// **********************************************
// function subGetEmailAddressDbl()
//
// This function gets the email addresses for both
// players from the configuration file
//
// **********************************************

function subGetEmailAddressDbl(ss, Addresses, WinPlyr, LosPlyr){
  
  // Players Sheets for Email addresses
  var shtPlayers = ss.getSheetByName('Players');
  var colEmail = 3;
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var rowWinr = 0;
  var rowLosr = 0;
  var PlyrRowStart = 3;
  
  var PlayerNames = shtPlayers.getRange(PlyrRowStart,2,NbPlayers,1).getValues();
  
  // Find Players rows
  for (var row = 0; row < NbPlayers; row++){
    if (PlayerNames[row] == WinPlyr) rowWinr = row + PlyrRowStart;
    if (PlayerNames[row] == LosPlyr) rowLosr = row + PlyrRowStart;
    if (rowWinr > 0 && rowLosr > 0) row = NbPlayers + 1;
  }
   
  // Get Email addresses using the players rows
  Addresses[1][0] = shtPlayers.getRange(rowWinr,colEmail+1).getValue(); // Language
  Addresses[1][1] = shtPlayers.getRange(rowWinr,colEmail).getValue();   // Email Address
  Addresses[2][0] = shtPlayers.getRange(rowLosr,colEmail+1).getValue(); // Language
  Addresses[2][1] = shtPlayers.getRange(rowLosr,colEmail).getValue();   // Email Address
    
  return Addresses;
}

// **********************************************
// function subGetEmailRecipients()
//
// This function searches for all players in the  
// list with the selected language
//
// **********************************************

function subGetEmailRecipients(shtPlayers, Language){
  
  // Function Variables
  var EmailRecipients = '';
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  var PlayersData = shtPlayers.getRange(3,2,NbPlayers,2).getValues(); // ..[0]= Email Address  ..[1]= Language 
  
  // Loop through all players selected languages and concatenate their email addresses 
  // if it matches the Language sent in parameter 
  for(var i = 0; i < NbPlayers; i++){
    if(PlayersData[i][1] == Language) {
      if(EmailRecipients !='') EmailRecipients += ', '
      EmailRecipients += PlayersData[i][0];
    }
  }
  
  return EmailRecipients;
}



// **********************************************
// function subCrtPlayerContact()
//
// This function creates the Player Contact in Gmail Account
//
// **********************************************

function subCrtPlayerContact(PlyrContactInfo){
  
  //  PlyrContactInfo[0]= First Name
  //  PlyrContactInfo[1]= Last Name
  //  PlyrContactInfo[2]= Email
  //  PlyrContactInfo[3]= Language

  Logger.log("Routine: subCrtPlayerContact");
  Logger.log("Player: %s, %s",PlyrContactInfo[1], PlyrContactInfo[0]);
    
  var Status = 'Player Contact Not Created';
      
  // Check if player is already a contact
  var PlayerContact = ContactsApp.getContact(PlyrContactInfo[2]);
  
  // If Player is a contact, update First and Last Name
  if(PlayerContact != null){
    PlayerContact.setGivenName(PlyrContactInfo[0]);
    PlayerContact.setFamilyName(PlyrContactInfo[1]);
    Logger.log('Contact Updated: %s',PlayerContact.getFullName())
    Status = 'Player Contact Updated';
  }
  
  // If Player is not a contact, create it
  else { 
    PlayerContact = ContactsApp.createContact(PlyrContactInfo[0], PlyrContactInfo[1], PlyrContactInfo[2]);
    // Get Player Contact
    PlayerContact = ContactsApp.getContact(PlyrContactInfo[2]);
    if(PlayerContact != null) Status = 'Player Contact Created';
  }
  
  Logger.log(Status);
  
  return Status;
}
    

// **********************************************
// function subAddPlayerContactGroup()
//
// This function adds a player to the Contact Group   
//
// **********************************************

function subAddPlayerContactGroup(shtConfig, PlyrContactInfo){

  //  PlyrContactInfo[0]= First Name
  //  PlyrContactInfo[1]= Last Name
  //  PlyrContactInfo[2]= Email
  //  PlyrContactInfo[3]= Language
  
  // Event Parameters
  var cfgEvntParam = shtConfig.getRange(4,4,48,1).getValues();
  
  var evntLocation = cfgEvntParam[0][0];
  var evntNameEN = cfgEvntParam[7][0];
  var evntNameFR = cfgEvntParam[8][0];
  var evntCntctGrpNameEN = evntLocation + " " + evntNameEN;
  var evntCntctGrpNameFR = evntLocation + " " + evntNameFR;

  var Status;
  var ContactGroupEN;
  var ContactGroupFR;
    
  Logger.log("Routine: subAddPlayerContactGroup");
  Logger.log("Player: %s, %s",PlyrContactInfo[1],PlyrContactInfo[0]);
    
  // Get Player Contact
  var PlayerContact = ContactsApp.getContact(PlyrContactInfo[2]);
  
  if(PlayerContact != null){
    if(PlyrContactInfo[3] == "English"){
      // Get Contact Group
      ContactGroupEN = ContactsApp.getContactGroup(evntCntctGrpNameEN);
      // If Contact Group does not exist, create it
      if(ContactGroupEN == null) ContactGroupEN = ContactsApp.createContactGroup(evntCntctGrpNameEN);
      // Add Contact to Contact Group
      ContactGroupEN.addContact(PlayerContact);
      Status = 'Player added to Contact Group';
    }
    if(PlyrContactInfo[3] == "Français"){
      // Get Contact Group
      ContactGroupFR = ContactsApp.getContactGroup(evntCntctGrpNameFR);
      // If Contact Group does not exist, create it
      if(ContactGroupFR == null) ContactGroupFR = ContactsApp.createContactGroup(evntCntctGrpNameFR);
      // Add Contact to Contact Group
      ContactGroupFR.addContact(PlayerContact);
      Status = 'Player added to Contact Group';
    }
  }
  
  if(Status != 'Player added to Contact Group') Status = 'Contact Group Error'
  
  Logger.log(Status);
  
  return Status;  
}


// **********************************************
// function subCheckDataConflict()
//
// This function verifies that two arrays of data 
// are the same. If two values are different,
// the function returns the Data ID where they
// differ. If no conflict is found, returns 0;
//
// **********************************************

function subCheckDataConflict(DataArray1, DataArray2, ColStart, ColEnd) {
  
  var DataConflict = 0;
  
  // Compare New Response Data and Match Data. If Data is not equal to the other
  for (var j = ColStart; j <= ColEnd; j++){
       
    // If Data Conflict is found, sets the data and sends email
    if (DataArray1[0][j] != DataArray2[0][j]) {
      DataConflict = j;
      j = ColEnd + 1;
    }
  }
  return DataConflict;
}

// **********************************************
// function subPlayerMatchValidation()
//
// This function verifies that the player was allowed 
// to play this match. It checks in the total amount of matches
// played by the player to allow the game to be posted
// The function returns 1 if the game is valid and 0 if not valid
//
// **********************************************

function subPlayerMatchValidation(ss, shtConfig, ParticipantName, MatchValidation) {
  
  // Get Configuration Data
  var cfgEventData = shtConfig.getRange(4, 2,16,1).getValues();
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  
  // Column Values for Rounds Sheets
  var colRndPlyr =     cfgColRndSht[ 0][0];
  var colRndStatus =   cfgColRndSht[ 1][0];
  var colRndMP =       cfgColRndSht[ 2][0];
  
  // Opens Cumulative Results tab
  var shtCumul = ss.getSheetByName('Cumulative Results');
    
  // Get Data from Cumulative Results
  var RoundNum =      cfgEventData[ 3][0];
  var CumulMaxMatch = cfgEventData[ 8][0];
  var NbPlyrs =       cfgEventData[ 9][0];
  var NbTeams =       cfgEventData[10][0];
  var EvntFormat =    cfgEventData[11][0];
  var shtRound = ss.getSheetByName('Round' + RoundNum);
  
  var PlayerStatus;
  var PlayerMatchPlayed;
  var Participants;
  var NbParticipants;
  
  // Select Participant if Single Players or Teams  
  if(EvntFormat == "Single"){
    Participants =   "Players";
    NbParticipants = NbPlyrs;
  }
  if(EvntFormat == "Team"){
    Participants =   "Teams";
    NbParticipants = NbTeams;
  }  
  // Look for Player Row and if Player is still Active or Eliminated
  //subFindPlayerRow(sheet, rowStart, colPlyr, length, PlayerName)
  var PlyrRow = subFindPlayerRow(shtCumul,5,colRndPlyr,NbParticipants,ParticipantName);
  
  PlayerMatchPlayed = shtRound.getRange(PlyrRow,colRndMP).getValue();
  PlayerStatus =      shtRound.getRange(PlyrRow,colRndStatus).getValue();
  MatchValidation[1] = PlayerMatchPlayed;

  // If Player is Active and Number of Matches Played is below or equal to the maximum permitted
  if (PlayerStatus == 'Active' && PlayerMatchPlayed + 1 <= CumulMaxMatch) MatchValidation[0] = 1;
  
  // If Player is Eliminated, return -1
  if (PlayerStatus == 'Eliminated') MatchValidation[0] = -1;
  
  // If Player has played more games (counting the one to be posted) than permitted, return -2
  if (PlayerMatchPlayed + 1 > CumulMaxMatch && PlayerStatus != 'Eliminated') MatchValidation[0] = -2;
  
  return MatchValidation;
}

// **********************************************
// function subGenErrorMsg()
//
// This function generates the Error Message according to 
// the value sent in argument
//
// **********************************************

function subGenErrorMsg(Status, ErrorVal,Param) {

  switch (ErrorVal){

    case -10 : Status[0] = ErrorVal; Status[1] = 'Duplicate Entry Found at Row ' + Param; break; 
    case -11 : Status[0] = ErrorVal; Status[1] = 'Winning Player is Eliminated from League'; break;  
    case -12 : Status[0] = ErrorVal; Status[1] = 'Winning Player has played too many matches'; break;  
    case -21 : Status[0] = ErrorVal; Status[1] = 'Losing Player is Eliminated from League'; break;  
    case -22 : Status[0] = ErrorVal; Status[1] = 'Losing Player has played too many matches'; break;  
    case -31 : Status[0] = ErrorVal; Status[1] = 'Both Players are Eliminated from the League'; break;  
    case -32 : Status[0] = ErrorVal; Status[1] = 'Winning Player is Eliminated from the League and Losing Player has played too many matches'; break;
    case -33 : Status[0] = ErrorVal; Status[1] = 'Winning Player has player too many matches and Losing Player is Eliminated from the League'; break;
    case -34 : Status[0] = ErrorVal; Status[1] = 'Both Players have played too many matches'; break;
    case -50 : Status[0] = ErrorVal; Status[1] = 'Illegal Match, Same Player selected for Win and Loss'; break; 
    case -60 : Status[0] = ErrorVal; Status[1] = 'Card Name not Found for Card Number: ' + Param; break;  // Card Name not Found
      
    case -97 : Status[0] = ErrorVal; Status[1] = 'Match Results Post Not Executed'; break;   
    case -98 : Status[0] = ErrorVal; Status[1] = 'Matching Response Search Not Executed'; break; 
    case -99 : Status[0] = ErrorVal; Status[1] = 'Duplicate Entry Search Not Executed'; break;    

//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission
//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission
//    case  : Status[0] = ErrorVal; Status[1] = ''; break;  // Add Error Message for Data Conflict on Dual Submission

}
  
return Status;
}


// **********************************************
// function subUpdateStatus()
//
// This function updates the status of 
// the entry currently processing
//
// **********************************************

function subUpdateStatus(shtRspn, RspnRow, ColStatus, ColStatusMsg, StatusNum) {
  
  var StatusMsg
  
  switch(StatusNum){
    case  0: StatusMsg = 'Not Processed'; break;
    case  1: StatusMsg = 'Process Starting'; break;
    case  2: StatusMsg = 'Finding Duplicate'; break;
    case  3: StatusMsg = 'Finding Dual Response'; break;
    case  4: StatusMsg = 'Post Results in Round Tab'; break;
    case  5: StatusMsg = 'Update DB and List'; break;
    case  6: StatusMsg = 'Data Processed'; break;
    case  7: StatusMsg = 'Sending Confirmation Email'; break;
    case  8: StatusMsg = 'Sending Process Error Email'; break;
    case  9: StatusMsg = 'Updating Match ID'; break;
    case 10: StatusMsg = 'Match Processed'; break;
	
  }
   
  // Updating Status Data
  shtRspn.getRange(RspnRow, ColStatus).setValue(StatusNum);
  shtRspn.getRange(RspnRow, ColStatusMsg).setValue(StatusMsg);

  return StatusMsg;
}

// **********************************************
// function subPlayerWithMost()
//
// This function searches for the player with the 
// most "Param" for a given Round
//
// **********************************************

function subPlayerWithMost(shtConfig, PlayerMostData, NbPlayers, shtRound){
 
  
  var cfgColRndSht = shtConfig.getRange(4,21,16,1).getValues();
  
  // Column Values
  var colPlyr = cfgColRndSht[0][0];
  var colTeam = cfgColRndSht[1][0];
  var colWin = cfgColRndSht[3][0];
  var colLos = cfgColRndSht[4][0];
  var colTie = cfgColRndSht[5][0];
  var colPoints = cfgColRndSht[6][0];
  var colWinPerc = cfgColRndSht[7][0];
  var colLocation = cfgColRndSht[8][0];
  var colBalanceBonus = cfgColRndSht[10][0];
  
  var colParam;
  var Rank = 0;
  var MostValue = 0;
  var TestValue = 0;
  
  // Get Round Data Array from Sheet
  // Rows :    Players  
  // Columns:  Round Sheet Column Value - 1
  var RoundData = shtRound.getRange(4,1,33,10).getValues();
  
  // Select Appropriate Column according to Param
  switch (PlayerMostData[0][0]){
    case 'Wins'    : colParam = colWin-1; break;
    case 'Loss'    : colParam = colLos-1; break;
    case 'Points'  : colParam = colPoints-1; break;
    case 'Win%'    : colParam = colWinPerc-1; break;
    case 'Store'   : colParam = colLocation-1; break;
  }
  
  // Loop through Selected Column to find the Player with the Most...
  for(var i=1; i<=NbPlayers; i++){
    TestValue = RoundData[i][colParam];
    // If an Equal Value is found
    if(TestValue == MostValue){
      Rank += 1;
      PlayerMostData[Rank][0] = RoundData[i][1];
      PlayerMostData[Rank][1] = MostValue;
    }

    // If a new Highest Value is found
    if(TestValue > MostValue) {
      // Clear Array
      for(var j=0; j<NbPlayers; j++){
        PlayerMostData[j][0] = '';
        PlayerMostData[j][1] = '';
      }
      // Write New Value
      MostValue = TestValue;
      PlayerMostData[0][0] = RoundData[i][1];
      PlayerMostData[0][1] = MostValue;
    }
  }
  return PlayerMostData; 
}


// **********************************************
// function subCrtMatchRepPlyrList()
//
// This function creates the Player List for
// the Match Report Form 
//
// **********************************************

function subCrtMatchRepPlyrList(shtConfig, shtPlayers, cfgEvntParam){

  // Number of Players
  var NbPlyr = shtConfig.getRange(13,2).getValue();
  
  // Routine Variables
  var Players;
  var PlayerList;
  
  // Transfers Players Double Array to Single Array
  if (NbPlyr > 0){
    Players = shtPlayers.getRange(3,2,NbPlyr,1).getValues();
    PlayerList = new Array(NbPlyr);
    for(var i = 0; i < NbPlyr; i++){
      PlayerList[i] = Players[i][0];
    }
  }
   
  return PlayerList;
}
    
// **********************************************
// function subCrtMatchRepTeamList()
//
// This function creates the Team List for
// the Match Report Form 
//
// **********************************************

function subCrtMatchRepTeamList(shtConfig, shtTeams, cfgEvntParam){

  // Event Parameters
  var evntTeamNbPlyr = cfgEvntParam[10][0];
  var evntTeamMatch = cfgEvntParam[11][0];
  
  //  Number of Players and Teams
  var NbPlyr = shtConfig.getRange(13,2).getValue();
  var NbTeam = shtConfig.getRange(14,2).getValue();
  
  // Routine Variables
  var Teams;
  var TeamList;
  var TeamListLen;
  var pntrTeam;
  var pntrPlyr;
  var dataList;
  
  // Get Teams Data
  // [0]= N/A, [1]= Team Name, [2-5]= N/A, [6-13]= Members 1-8 
  Teams = shtTeams.getRange(3,2,NbTeam,14).getValues();
  
  // Set the Team and Player Pointer values to 0 and 6;
  pntrTeam = 0;
  pntrPlyr = 6;
      
  // If Teams are 2 players and Matches are 1v1
  if(evntTeamNbPlyr == '2' && evntTeamMatch == '1v1'){
    // Create the Team List Array
    TeamListLen = NbPlyr;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0] + ' - ' + Teams[pntrTeam][pntrPlyr];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      if(pntrPlyr == 6)  pntrPlyr = 7;
      if(pntrPlyr == 7){
        pntrPlyr = 6;
        pntrTeam++;
      }
    }
  }
  
  // If Teams are 2 players and Matches are 2v2
  if(evntTeamNbPlyr == '2' && evntTeamMatch == '2v2'){
    // Create the Team List Array
    TeamListLen = NbTeam;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      pntrTeam++;
    }
  }
  
  // If Teams are 3 players and Matches are 1v1
  if(evntTeamNbPlyr == '3' && evntTeamMatch == '1v1'){
    // Create the Team List Array
    TeamListLen = NbPlyr;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0] + ' - ' + Teams[pntrTeam][pntrPlyr];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      if(pntrPlyr == 6)  pntrPlyr = 7;
      if(pntrPlyr == 7)  pntrPlyr = 8;
      if(pntrPlyr == 8){
        pntrPlyr = 6;
        pntrTeam++;
      }
    }
  }
  
  // If Teams are 3 players and Matches are 3v3
  if(evntTeamNbPlyr == '3' && evntTeamMatch == '3v3'){
    // Create the Team List Array
    TeamListLen = NbTeam;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      pntrTeam++;
    }
  }
  
  // If Teams are 4 players and Matches are 1v1
  if(evntTeamNbPlyr == '4' && evntTeamMatch == '1v1'){
    // Create the Team List Array
    TeamListLen = NbPlyr;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0] + ' - ' + Teams[pntrTeam][pntrPlyr];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      if(pntrPlyr == 6)  pntrPlyr = 7;
      if(pntrPlyr == 7)  pntrPlyr = 8;
      if(pntrPlyr == 8)  pntrPlyr = 9;
      if(pntrPlyr == 9){
        pntrPlyr = 6;
        pntrTeam++;
      }
    }
  }
  
  // If Teams are 4 players and Matches are 2v2 (Team A - Team B)
  if(evntTeamNbPlyr == '4' && evntTeamMatch == '2v2'){
    // Create the Team List Array
    TeamListLen = NbTeam;
    //TeamListLen = NbTeam * 2;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      pntrTeam++;
    }
    
//    for(var i = 0; i < TeamListLen; i++){
//      // Get the Team Name and Team Member Combination (P1 + P2, P3 + P4, P1 + P3, P2 + P4, P1 + P4, P2 + P3)
//      if(pntrPlyr == 6) dataList = Teams[pntrTeam][1] + ' -  A';
//      if(pntrPlyr == 7) dataList = Teams[pntrTeam][1] + ' -  B';
//      // Populate Array
//      TeamList[i] = dataList;
//      // Update Pointers accordingly
//      if(pntrPlyr == 6)  pntrPlyr = 7;
//      if(pntrPlyr == 7){
//        pntrPlyr = 6;
//        pntrTeam++;
//      }
//    }
  }
  
  // If Teams are 4 players and Matches are 4v4
  if(evntTeamNbPlyr == '4' && evntTeamMatch == '4v4'){
    // Create the Team List Array
    TeamListLen = NbTeam;
    TeamList = new Array(TeamListLen);
    // Fill the Array
    for(var i = 0; i < TeamListLen; i++){
      // Get the Team Name and Team Member
      dataList = Teams[pntrTeam][0];
      // Populate Array
      TeamList[i] = dataList;
      // Update Pointers accordingly
      pntrTeam++;
    }
  }
  return TeamList;   
}

// **********************************************
// function subUpdatePlayerMember()
//
// This function updates the Player Sheet with 
// the Member Info
//
// **********************************************

function subUpdatePlayerMember(shtConfig, shtPlayers, Member){

  // Registration Form Construction 
  // Column 1 = Category Name
  // Column 2 = Category Order in Form
  // Column 3 = Column Value in Player/Team Sheet
  var cfgRegFormCnstrVal = shtConfig.getRange(4,26,20,3).getValues();
  var colTblMemberFileID = cfgRegFormCnstrVal[17][2];
  var NbPlayers = shtPlayers.getRange(2,1).getValue();
  
  shtPlayers.getRange(NbPlayers+2,colTblMemberFileID).setValue(Member[7]);

}

// **********************************************
// function subUpdatePlyrEvntRecord()
//
// This function updates the Player Event Record
//
// **********************************************

function subUpdatePlyrEvntRecord(cfgEvntParam, RndRecPlyr, GameResult, MatchDataPts){
  
  var evntPtsPerWin =      cfgEvntParam[29][0];
  var evntPtsPerLoss =     cfgEvntParam[30][0];
  var evntPtsPerTie =      cfgEvntParam[31][0];
  var evntPtsGainedMatch = cfgEvntParam[32][0];
  
  // Initializes Player Round Record
  if (RndRecPlyr[0][0] == '') RndRecPlyr[0][0] = 0; // MP
  if (RndRecPlyr[0][1] == '') RndRecPlyr[0][1] = 0; // Win
  if (RndRecPlyr[0][2] == '') RndRecPlyr[0][2] = 0; // Loss
  if (RndRecPlyr[0][3] == '') RndRecPlyr[0][3] = 0; // Tie
  if (RndRecPlyr[0][4] == '') RndRecPlyr[0][4] = 0; // Points
  if (RndRecPlyr[0][5] == '') RndRecPlyr[0][5] = 0; // Win %
  
  // Update Player Matches Played
  RndRecPlyr[0][0] = RndRecPlyr[0][0] + 1;
  
  // Update Player Wins
  if(GameResult == "Win") RndRecPlyr[0][1] = RndRecPlyr[0][1] + 1;
  
  // Update Player Loss
  if(GameResult == "Loss") RndRecPlyr[0][2] = RndRecPlyr[0][2] + 1;
  
  // Update Player Tie
  if(GameResult == "Tie") RndRecPlyr[0][3] = RndRecPlyr[0][3] + 1;
  
   // Update Points
  // If Points per Game are not used, Points are equal to the sum of Wins*PtsPerWin + Loss*PtsPerLoss + Ties*PtsPerTie
  if(evntPtsGainedMatch == "Disabled"){
    RndRecPlyr[0][4] = (RndRecPlyr[0][1] * evntPtsPerWin) + (RndRecPlyr[0][2] * evntPtsPerLoss) + (RndRecPlyr[0][3] * evntPtsPerTie);
    }
  
  // If Points per Game are used, Points are equal to the sum of all points made during all matches
  if(evntPtsGainedMatch == "Enabled"){
    RndRecPlyr[0][4] = RndRecPlyr[0][4] + MatchDataPts;
  }
  
  // Update Win Percentage
  if(RndRecPlyr[0][0] > 0) RndRecPlyr[0][5] = RndRecPlyr[0][1] / RndRecPlyr[0][0];
  
  return RndRecPlyr;
}