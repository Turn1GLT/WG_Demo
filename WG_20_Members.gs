// **********************************************
// function fcnRegistrationWG
//
// This function searches for a Member of the 
// Turn1 GLT Member List and returns the 
// missing information (Filee ID and Link)
//
// **********************************************

function fcnSearchMember(Member){
  
  Logger.log("Routine: fcnSearchMember");
  
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Record File ID
  //  Member[ 2] = Member Record File Link
  //  Member[ 3] = Member Full Name
  //  Member[ 4] = Member First Name
  //  Member[ 5] = Member Last Name
  //  Member[ 6] = Member Email
  //  Member[ 7] = Member Language
  //  Member[ 8] = Member Phone Number
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  // Defines the Search
  var searchFor ='title contains "' + Member[3] + '"';
  var names =   [];
  var fileIds = [];
  
  // Start the Search
  var files = DriveApp.searchFiles(searchFor);
  
  // Populate the Array
  while (files.hasNext()) {
    var file = files.next();
    var fileId = file.getId();// To get FileId of the file
    fileIds.push(fileId);
    var name = file.getName();
    names.push(name);
  }

  // If Member is Not Found
  if(fileIds.length == 0){
    Member[1] = 'Member Not Found';
    Member[2] = '';
  }
  
  // If Member is Found
  if(fileIds.length > 0){
    Member[1] = fileIds[0];
    Member[2] = 'https://docs.google.com/spreadsheets/d/' + fileIds[0];
  }
  return Member;
}


// **********************************************
// function fcnCreateMember
//
// This function adds the new player to the 
// Turn1 GLT Member List
//
// **********************************************

function fcnCreateMember(Member) {
  
  Logger.log("Routine: fcnCreateMember");
  
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Record File ID
  //  Member[ 2] = Member Record File Link
  //  Member[ 3] = Member Full Name
  //  Member[ 4] = Member First Name
  //  Member[ 5] = Member Last Name
  //  Member[ 6] = Member Email
  //  Member[ 7] = Member Language
  //  Member[ 8] = Member Phone Number
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  // Create Player Record Spreadsheet
  var ss = SpreadsheetApp.create(Member[3]);
  var id = ss.getId();
  // Get File
  var file = DriveApp.getFileById(id);
  var folders = DriveApp.getFoldersByName('Members');
  while (folders.hasNext()) {
    var folder = folders.next();
  }
  // Move File to Folder
  folder.addFile(file);
  // Remove File from Root Folder
  DriveApp.getRootFolder().removeFile(file);
  
  // Open Member List (File _Member List)
  var shtList = SpreadsheetApp.openById("1wQVwIIFJoRTNfK3oV4ioeFU3hw6JnEA4MXmc-GI8Vgw").getSheetByName("List");
  var listNbMembers = shtList.getRange(2,1).getValue();
  var listMaxCol = shtList.getMaxColumns();
  var listMaxRow = shtList.getMaxRows();
  var listNextMemberRow = listNbMembers + 3; // First Member is at Row 3
  
  var NewMemberID = listNbMembers + 1;
  
  // Open Member Spreadsheet
  var ssMember = SpreadsheetApp.openById(id);
  
  // Copy Templates to Member Spreadsheet
  var ssTemplates = SpreadsheetApp.openById("1HPnrUdIen2X0YeV1R2eNf5CEOZ9Yq_GYJEWTfAuIAMU");
  // English Files
  if(Member[7] == "English"){
    ssTemplates.getSheetByName("Member Info EN").copyTo(ssMember);
    ssTemplates.getSheetByName("WG Record Template EN").copyTo(ssMember);
    ssTemplates.getSheetByName("TCG Record Template EN").copyTo(ssMember);
    ssTemplates.getSheetByName("BG Record Template EN").copyTo(ssMember);
  }
  //  French Files
  if(Member[7] == "Français"){
    ssTemplates.getSheetByName("Member Info FR").copyTo(ssMember);
    ssTemplates.getSheetByName("WG Record Template FR").copyTo(ssMember);
    ssTemplates.getSheetByName("TCG Record Template FR").copyTo(ssMember);
    ssTemplates.getSheetByName("BG Record Template FR").copyTo(ssMember);
  }
  // Delete Sheet1
  ssMember.deleteSheet(ssMember.getSheets()[0]);
  
  // Get Number of Sheets
  var memberNbSheets = ssMember.getNumSheets();
  var memberSheets = ssMember.getSheets();
  var SheetName;
  var Sheet;
  
  // Rename all Copied Sheets
  for(var sht= 0; sht < memberNbSheets; sht++){
    Sheet = memberSheets[sht];
    SheetName = Sheet.getName();
    // Rename all tabs
    switch(SheetName){
      case "Copy of Member Info EN"         : Sheet.setName("Info"); break;
      case "Copy of WG Record Template EN"  : Sheet.setName("WG Record"); break;
      case "Copy of TCG Record Template EN" : Sheet.setName("TCG Record"); break;
      case "Copy of BG Record Template EN"  : Sheet.setName("BG Record"); break;
      case "Copy of Member Info FR"         : Sheet.setName("Info"); break;
      case "Copy of WG Record Template FR"  : Sheet.setName("WG Fiche"); break;
      case "Copy of TCG Record Template FR" : Sheet.setName("TCG Fiche"); break;
      case "Copy of BG Record Template FR"  : Sheet.setName("BG Fiche"); break;
    }
  }
  
  // Update Member Info with New Data
  Member[0] = NewMemberID;
  Member[1] = id;
  Member[2] = 'https://docs.google.com/spreadsheets/d/' + id;
  
  
  // Write Member Values in _Member List
  var shtInfo = ssMember.getSheetByName("Info");
  var valMemberSheet = shtInfo.getRange(1, 2, 16, 1).getValues();
  var valMemberList =  shtList.getRange(listNextMemberRow,1, 1, listMaxCol).getValues();
  
  // Set Member Info Values for both sheets
  for(var i= 0; i<16; i++){
    valMemberList[0][i+1] = Member[i];
    valMemberSheet[i][0]  = Member[i];
    
  }
  // Add New Member to _Member List
  shtList.insertRowBefore(listMaxRow);
  shtList.getRange(listNextMemberRow+1,1, 1, listMaxCol).setValues(valMemberList);
  shtList.getRange(listNextMemberRow+1, 1).setValue('=if(INDIRECT("R[0]C[1]",false)<>"",1,0)');
  
  // Write Member Info to Member Spreadsheet
  shtInfo.getRange(1, 2, 16, 1).setValues(valMemberSheet);
  return Member;
}