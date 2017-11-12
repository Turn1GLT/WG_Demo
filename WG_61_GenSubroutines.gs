// **********************************************
// function subLogger()
//
// This function posts the log to a spreadsheet
//
// **********************************************

function subPostLog(shtLog) {
  
  //var shtLog = SpreadsheetApp.setActiveSheet(1);
  
  shtLog.insertRowAfter(2);
  shtLog.getRange(3,1).setValue('=if(INDIRECT("R[0]C[1]",FALSE)<>"",1,"")');
  shtLog.getRange(3,2).setValue(new Date()).setNumberFormat('yyyy-MM-dd / HH:mm:ss');
  shtLog.getRange(3,3).setValue(Logger.getLog());
//  Logger.clear();
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

function subCheckDataConflict(DataArray1, DataArray2, ColStart, ColEnd, shtTest) {
  
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

function subPlayerMatchValidation(ss, PlayerName, MatchValidation, shtTest) {
  
  // Opens Cumulative Results tab
  var shtCumul = ss.getSheetByName('Cumulative Results');
    
  // Get Data from Cumulative Results
  var CumulMaxMatch = shtCumul.getRange(4,3).getValue();
  var CumulPlyrData = shtCumul.getRange(5,1,32,11).getValues();
  var WeekNum = shtCumul.getRange(2,3).getValue();
  var shtWeek = ss.getSheetByName('Week' + WeekNum);
  var WeekPlyrData = shtWeek.getRange(5,1,32,11).getValues(); // Data[i][j] i = Player List 1-32, j = ID(0), Name(1), Initials(2), MP(3), W(4), L(5), %(6), Penalty(7), Matches in Store(8) Packs(9), Status(10)
  
  var PlayerStatus;
  var PlayerMatchPlayed;
  
  // Look for Player Row and if Player is still Active or Eliminated
  for (var i = 0; i < 32; i++) {
    // Player Found, Number of Match Played and Status memorized
    if (PlayerName == WeekPlyrData[i][1]){
      PlayerMatchPlayed = WeekPlyrData[i][3];
      PlayerStatus = CumulPlyrData[i][10];
      MatchValidation[1] = PlayerMatchPlayed;
      i = 32; // Exit Loop
    }
  }

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
    case  4: StatusMsg = 'Post Results in Week Tab'; break;
    case  5: StatusMsg = 'Update Card DB and Card List'; break;
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
// function fcnPlayerWithMost()
//
// This function searches for the player with the 
// most "Param" for a given week
//
// **********************************************

function fcnPlayerWithMost(PlayerMostData, NbPlayers, shtWeek){
 
  var ColParam;
  var Rank = 0;
  var MostValue = 0;
  var TestValue = 0;
  
  // Get Week Data Array from Sheet
  // Rows
  // Players
  
  // Columns
  // 0 = Player Name, 3 = Matches Played, 4 = Wins, 5 = Loss, 6 = Win%, 7 = Penalty Loss, 8 = Matches Played in Store, 9 = Punishment Packs
  var WeekData = shtWeek.getRange(4,1,33,10).getValues();
  
  // Select Appropriate Column according to Param
  switch (PlayerMostData[0][0]){
    case 'Wins'    : ColParam = 4; break;
    case 'Loss'    : ColParam = 5; break;
    case 'Win%'    : ColParam = 6; break;
    case 'Store'   : ColParam = 8; break;
    case 'PunPack' : ColParam = 9; break;
  }
  
  // Loop through Selected Column to find the Player with the Most...
  for(var i=1; i<=NbPlayers; i++){
    TestValue = WeekData[i][ColParam];
    // If an Equal Value is found
    if(TestValue == MostValue){
      Rank += 1;
      PlayerMostData[Rank][0] = WeekData[i][1];
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
      PlayerMostData[0][0] = WeekData[i][1];
      PlayerMostData[0][1] = MostValue;
    }
  }
  return PlayerMostData; 
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
      
      
   
      
      
      