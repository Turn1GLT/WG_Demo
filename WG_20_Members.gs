// **********************************************
// function fcnRegistrationWG
//
// This function searches for a Member of the 
// Turn1 GLT Member List and returns the 
// missing information (Filee ID and Link)
//
// **********************************************

function fcnSearchMember(Member){

  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Full Name
  //  Member[ 2] = Member First Name
  //  Member[ 3] = Member Last Name
  //  Member[ 4] = Member Email
  //  Member[ 5] = Member Language
  //  Member[ 6] = Member Phone Number
  //  Member[ 7] = Member Record File ID
  //  Member[ 8] = Member Record File Link
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  // Defines the Search
  var searchFor ='title contains "' + Member[0] + '"';
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
    Member[7] = 'File Not Found';
    Member[8] = '';
  }
  
  // If Member is Found
  if(fileIds.length > 0){
    Member[7] = fileIds[0];
    Member[8] = 'https://docs.google.com/spreadsheets/d/' + fileIds[0];
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
  //  Member[ 0] = Member ID
  //  Member[ 1] = Member Full Name
  //  Member[ 2] = Member First Name
  //  Member[ 3] = Member Last Name
  //  Member[ 4] = Member Email
  //  Member[ 5] = Member Language
  //  Member[ 6] = Member Phone Number
  //  Member[ 7] = Member Record File ID
  //  Member[ 8] = Member Record File Link
  //  Member[ 9] = Member Spare
  //  Member[10] = Member Spare
  //  Member[11] = Member Spare
  //  Member[12] = Member Spare
  //  Member[13] = Member Spare
  //  Member[14] = Member Spare
  //  Member[15] = Member Spare
  
  // Create Player Record Spreadsheet
  var ss = SpreadsheetApp.create(Member[1]);
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
  ssTemplates.getSheetByName("Member Info").copyTo(ssMember);
  ssTemplates.getSheetByName("WG Record Template").copyTo(ssMember);
  ssTemplates.getSheetByName("TCG Record Template").copyTo(ssMember);
  ssTemplates.getSheetByName("BG Record Template").copyTo(ssMember);
  
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
    Logger.log("sht: %s - SheetName: %s",sht,SheetName);
    // Rename all tabs
    switch(SheetName){
      case "Copy of Member Info"         : Sheet.setName("Info"); break;
      case "Copy of WG Record Template"  : Sheet.setName("WG Record"); break;
      case "Copy of TCG Record Template" : Sheet.setName("TCG Record"); break;
      case "Copy of BG Record Template"  : Sheet.setName("BG Record"); break;
    }
  }
  
  // Update Member Info with New Data
  Member[0] = NewMemberID;
  Member[7] = id;
  Member[8] = 'https://docs.google.com/spreadsheets/d/' + id;
  
  
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
  
  Logger.log(valMemberList);
  Logger.log(valMemberSheet);
  return Member;
}