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
// function fcnSendConfirmEmailEN()
//
// This function generates the confirmation email in English
// after a match report has been submitted
//
// **********************************************

function fcnSendConfirmEmailEN(shtConfig, Address, MatchData) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(17,2,3,1).getValues();
  var StandingsUrl = UrlValues[0][0];
  var CardPoolUrl = UrlValues[1][0];
  var MatchReporterUrl = UrlValues[2][0];
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(25,2).getValue();
  var ssEmail = SpreadsheetApp.openById(ssEmailID);
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,6,29,1).getValues();
  
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueNameEN = Location + ' ' + LeagueTypeEN;
  
  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
  
  // Set Email Subject
  EmailSubject = LeagueNameEN + " - Week " + Week + " - Match Result" ;
    
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been received and succesfully processed for the ' + LeagueNameEN + ', Week ' + Week + 
    '<br><br>Here is your match result:<br><br>';
    
  // Generate Match Data Table
  EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'EN');
  
  EmailMessage += "<br>Click below to access the League Standings and Results:"+
    "<br>"+ StandingsUrl +
      "<br><br>Click below to access your Card Pool:"+
        "<br>"+ CardPoolUrl +
          "<br><br>Click below to send another Match Report:"+
            "<br>"+ MatchReporterUrl +
              "<br><br>If you find any problems with your match result, please reply to this message and describe the situation as best you can. You will receive a response once it has been processed."+
                "<br><br>Thank you for using Wargames League Manager from Triad Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Sends email to both players with the Match Data
  if (Address[1][0] == 'English' && Address[1][1] != '') MailApp.sendEmail(Address[1][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  if (Address[2][0] == 'English' && Address[2][1] != '') MailApp.sendEmail(Address[2][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
}


// **********************************************
// function fcnSendErrorEmailEN()
//
// This function generates the error email in English
// after a match report has been submitted
//
// **********************************************

function fcnSendErrorEmailEN(shtConfig, Address, MatchData, MatchID, Status) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var EmailName1 = '';
  var EmailName2 = '';
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(17,2,3,1).getValues();
  var StandingsUrl = UrlValues[0][0];
  var CardPoolUrl = UrlValues[1][0];
  var MatchReporterUrl = UrlValues[2][0];
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(25,2).getValue();
  var ssEmail = SpreadsheetApp.openById(ssEmailID);
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,6,29,1).getValues();
  
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueNameEN = Location + ' ' + LeagueTypeEN;
  
  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
  
  var StatusMsg;
   
  // Selects the Appropriate Error Message
  switch (Status[0]){
  
    case -10 : StatusMsg = 'Match Result has already been received and processed.'; break; // Administrator + Players
    case -11 : StatusMsg = '<b>'+Winr+'</b> is eliminated from League.'; break;    // Administrator + Players
    case -12 : StatusMsg = '<b>'+Winr+'</b> has played too many matches this week. Matches played: '+MatchData[4][1]; break;  // Administrator + Players 
    case -21 : StatusMsg = '<b>'+Losr+'</b> is eliminated from League.'; break;    // Administrator + Players
    case -22 : StatusMsg = '<b>'+Losr+'</b> has played too many matches this week. Matches played: '+MatchData[5][1]; break;  // Administrator + Players 
    case -31 : StatusMsg = 'Both players are eliminated from League.'; break; // Administrator + Players 
    case -32 : StatusMsg = '<b>'+Winr+'</b> is eliminated from League.<br><b>'+Losr+'</b> has played too many matches this week. Matches played: '+MatchData[5][1]; break;  // Administrator + Players
    case -33 : StatusMsg = '<b>'+Winr+'</b> has player too many matches this week. Matches played: <b>'+MatchData[4][1]+'</b>.<br><b>'+Losr+'</b> is eliminated from League.'; break;  // Administrator + Players
    case -34 : StatusMsg = 'Both Players have played too many matches this week.<br><b>'+Winr+'</b> Matches played: <b>'+MatchData[4][1]+'</b><br><b>'+Losr+'</b> Matches played: <b>'+MatchData[5][1]+'</b>'; break; // Administrator + Players
    case -50 : StatusMsg = 'Same player selected for Win and Loss.<br>Winner: <b>'+Winr+'</b><br>Loser: <b>' +Losr+ '</b>'; break; // Administrator + Players
    case -60 : StatusMsg = Status[1]; break;  // Administrator + Players
	case -97 : StatusMsg = 'Process Error, Match Results Post Not Executed'; break;        // Administrator
    case -98 : StatusMsg = 'Process Error, Matching Response Search Not Executed'; break;  // Administrator
    case -99 : StatusMsg = 'Process Error, Duplicate Entry Search Not Executed'; break;    // Administrator
  }
  
  // Set Email Subject
  EmailSubject = LeagueNameEN + ' - Week ' + Week + ' - Match Report Error' ;
  
  // Start of Email Message
  EmailMessage = '<html><body>';

  // If Error prevented Match Data to be processed (Duplicate Entry or Player Match is not valid)  
  if (Status[0] < 0 && Status[0] > -60) {
    EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been succesfully received for the ' + LeagueNameEN + ', Week ' + Week + 
      "<br><br>An error has been detected in one of the player's record. Unfortunately, this error prevented us to process the match report.<br><br>"+
        "<b>Error Detected</b><br>" + StatusMsg +
          '<br><br>Here is your match result:<br><br>';
    
    // Populate the Match Data Table
    EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'EN');
  }

  // If Error did not prevent Match Data to be processed (Card Name not Found for Card Number X)    
  if (Status[0] == -60){
    EmailMessage += 'Hi ' + Winr + ' and ' + Losr + ',<br><br>Your match result has been succesfully received for the ' + LeagueNameEN + ', Week ' + Week + 
      "<br><br>We were able to process the match data but an error has been detected in the submitted form.<br>Please contact us to resolve this error as soon as possible<br><br>"+
        "<b>Error Detected</b><br>" + StatusMsg +
          '<br><br>Here is your match result:<br><br>';
    
    // Populate the Match Data Table
    EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'EN');
  }

  // If Process Error was Detected 
  if (Status[0] < -60) {
    EmailMessage += 'Process Error was detected<br><br>'+
        "<b>Error Detected</b><br>" + StatusMsg;
  }
  
  if (Status[0] >= -60) {
    EmailMessage += "<br>Click below to access the League Standings and Results:"+
      "<br>"+ StandingsUrl +
        "<br><br>Click below to access your Card Pool:"+
          "<br>"+ CardPoolUrl +
            "<br><br>Click below to send another Match Report:"+
              "<br>"+ MatchReporterUrl +
                "<br><br>If you find any problems with your match result, please reply to this message and describe the situation as best you can. You will receive a response once it has been processed."+
                  "<br><br>Thank you for using Wargames League Manager from Triad Gaming Leagues & Tournaments";
  }
  
  // End of Email Message
  EmailMessage += '</body></html>';
   
  // Send email to Administrator
  MailApp.sendEmail(Address[0][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  
  // If Error is between 0 and -60, send email to players. If not, only send to Administrator
  if (Status[0] >= -60){
    // Sends email to both players with the Match Data
    if (Address[1][0] == 'English' && Address[1][1] != '') {
      MailApp.sendEmail(Address[1][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
    }
    if (Address[2][0] == 'English' && Address[2][1] != ''&& Address[1][1] != Address[2][1]) {
      MailApp.sendEmail(Address[2][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
    }
  }
}


// **********************************************
// function fcnSendConfirmEmailFR()
//
// This function generates the confirmation email in French
// after a match report has been submitted
//
// **********************************************

function fcnSendConfirmEmailFR(shtConfig, Address, MatchData) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(20,2,3,1).getValues();
  var StandingsUrl = UrlValues[0][0];
  var CardPoolUrl = UrlValues[1][0];
  var MatchReporterUrl = UrlValues[2][0];
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(25,2).getValue();
  var ssEmail = SpreadsheetApp.openById(ssEmailID);
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,7,29,1).getValues();
  
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(14,2).getValue();
  var LeagueNameFR = LeagueTypeFR + ' ' + Location;

  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];

  // Set Email Subject
  EmailSubject = LeagueNameFR + " - Week " + Week + " - Rapport de Match" ;
    
  // Start of Email Message
  EmailMessage = "<html><body>";
  
  EmailMessage += "Bonjour " + Winr + " et " + Losr + ",<br><br>Nous confirmons que nous avons bien reçu et traité le rapport de votre match de la " + LeagueNameFR + ", Semaine " + Week + 
    "<br><br>Voici le sommaire de votre match:<br><br>";
    
  // Generate Match Data Table
  EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'FR');
  
  EmailMessage += "<br>Cliquez ci-dessous pour accéder au classement et statistiques de la ligue:"+
    "<br>"+ StandingsUrl +
      "<br><br>Cliquez ci-dessous pour accéder à votre pool de cartes:"+
        "<br>"+ CardPoolUrl +
          "<br><br>Cliquez ci-dessous pour envoyer un autre rapport de match:"+
            "<br>"+ MatchReporterUrl +
              "<br><br>Si vous remarquez quel problème que ce soit dans ce rapport, SVP répondez à ce courriel en décrivant la situation de votre mieux. Vous recevrez une réponse dès que la situation sera traitée."+
                "<br><br>Merci d'utiliser Wargames League Manager de Triad Gaming Leagues & Tournaments";
  
  // End of Email Message
  EmailMessage += "</body></html>";
  
  // Sends email to both players with the Match Data
  if (Address[1][0] == 'Français' && Address[1][1] != '') MailApp.sendEmail(Address[1][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  if (Address[2][0] == 'Français' && Address[2][1] != '') MailApp.sendEmail(Address[2][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
}


// **********************************************
// function fcnSendErrorEmailFR()
//
// This function generates the error email in French
// after a match report has been submitted
//
// **********************************************

function fcnSendErrorEmailFR(shtConfig, Address, MatchData, MatchID, Status) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  var EmailName1 = '';
  var EmailName2 = '';
  
  // Get Document URLs
  var UrlValues = shtConfig.getRange(20,2,3,1).getValues();
  var StandingsUrl = UrlValues[0][0];
  var CardPoolUrl = UrlValues[1][0];
  var MatchReporterUrl = UrlValues[2][0];
  
  // Open Email Templates
  var ssEmailID = shtConfig.getRange(25,2).getValue();
  var ssEmail = SpreadsheetApp.openById(ssEmailID);
  var shtEmailTemplates = ssEmail.getSheetByName('Templates');
  var Headers = shtEmailTemplates.getRange(3,7,29,1).getValues();
    
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeFR = shtConfig.getRange(14,2).getValue();
  var LeagueNameFR = LeagueTypeFR + ' ' + Location;

  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
  
  var StatusMsg;
   
  // Selects the Appropriate Error Message
  switch (Status[0]){
  
    case -10 : StatusMsg = 'Le résultat de ce match a déjà été reçu et traité.'; break; // Administrator + Players
    case -11 : StatusMsg = '<b>'+Winr+'</b> est éliminé(e) de la ligue.'; break;    // Administrator + Players
    case -12 : StatusMsg = '<b>'+Winr+'</b> a joué le maximum de match permis. Matches joués: '+MatchData[4][1]; break;  // Administrator + Players 
    case -21 : StatusMsg = '<b>'+Losr+'</b> est éliminé(e) de la ligue.'; break;    // Administrator + Players
    case -22 : StatusMsg = '<b>'+Losr+'</b> a joué le maximum de match permis. Matches joués: '+MatchData[5][1]; break;  // Administrator + Players 
    case -31 : StatusMsg = 'Les deux joueurs sont éliminés de la ligue.'; break; // Administrator + Players 
    case -32 : StatusMsg = '<b>'+Winr+'</b> est éliminé(e) de la ligue.<br><b>'+Losr+'</b> a joué le maximum de match permis. Matches joués: '+MatchData[5][1]; break;  // Administrator + Players
    case -33 : StatusMsg = '<b>'+Winr+'</b> a joué le maximum de match permis. Matches joués: <b>'+MatchData[4][1]+'</b>.<br><b>'+Losr+'</b> est éliminé(e) de la ligue.'; break;  // Administrator + Players
    case -34 : StatusMsg = 'Les deux joueurs ont joué le maximum de match permis.<br><b>'+Winr+'</b> Matches joués: <b>'+MatchData[4][1]+'</b><br><b>'+Losr+'</b> Matches joués: <b>'+MatchData[5][1]+'</b>'; break; // Administrator + Players
    case -50 : StatusMsg = 'Le même joueur a été sélectionné comme joueur gagnant et perdant.<br>Joueur gagnant: <b>'+Winr+'</b><br>Joueur perdant: <b>' +Losr+ '</b>'; break; // Administrator + Players
    case -60 : StatusMsg = Status[1]; break;  // Administrator + Players
	case -97 : StatusMsg = 'Process Error, Match Results Post Not Executed'; break;        // Administrator
    case -98 : StatusMsg = 'Process Error, Matching Response Search Not Executed'; break;  // Administrator
    case -99 : StatusMsg = 'Process Error, Duplicate Entry Search Not Executed'; break;    // Administrator
  }
  
  // Set Email Subject
  EmailSubject = LeagueNameFR + ' - Week ' + Week + ' - Erreur Rapport de Match' ;
  
  // Start of Email Message
  EmailMessage = "<html><body>";

  // If Error prevented Match Data to be processed (Duplicate Entry or Player Match is not valid)  
  if (Status[0] < 0 && Status[0] > -60) {
    EmailMessage += "Bonjour " + Winr + " et " + Losr + ",<br><br>Nous confirmons que nous avons bien reçu le résultat de votre match de la " + LeagueNameFR + ", Semaine " + Week + 
      "<br><br>Nous avons détecté une erreur dans la fiche d'un joueur qui nous a empêché de traiter le rapport du match.<br><br>"+
        "<b>Erreur détectée</b><br>" + StatusMsg +
          "<br><br>Voici le sommaire de votre match:<br><br>";
    
    // Populate the Match Data Table
    EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'FR');
  }

  // If Error did not prevent Match Data to be processed (Card Name not Found for Card Number X)    
  if (Status[0] == -60){
    EmailMessage += "Bonjour " + Winr + " et " + Losr + ",<br><br>Nous confirmons que nous avons bien reçu le résultat de votre match de la " + LeagueNameFR + ", Semaine " + Week + 
      "<br><br>Nous avons été en mesure de traiter le rapport de votre match mais avons détecté une erreur dans les informations reçues.<br>SVP, contactez-nous le plus rapidement possible pour corriger cette erreur<br><br>"+
        "<b>Erreur détectée</b><br>" + StatusMsg +
          "<br><br>Voici le sommaire de votre match:<br><br>";
    
    // Populate the Match Data Table
    EmailMessage = subMatchReportTable(EmailMessage, Headers, MatchData,'FR');
  }

  // If Process Error was Detected 
  if (Status[0] < -60) {
    EmailMessage += "Process Error was detected<br><br>"+
      "<b>Erreur détectée</b><br>" + StatusMsg;
  }
  
  if (Status[0] >= -60) {
    EmailMessage += "<br>Cliquez ci-dessous pour accéder au classement et statistiques de la ligue:"+
      "<br>"+ StandingsUrl +
        "<br><br>Cliquez ci-dessous pour accéder à votre pool de cartes:"+
          "<br>"+ CardPoolUrl +
            "<br><br>Cliquez ci-dessous pour envoyer un autre rapport de match:"+
              "<br>"+ MatchReporterUrl +
                "<br><br>Si vous remarquez quel problème que ce soit dans ce rapport, SVP répondez à ce courriel en décrivant la situation de votre mieux. Vous recevrez une réponse dès que la situation sera traitée."+
                  "<br><br>Merci d'utiliser Wargames League Manager de Triad Gaming Leagues & Tournaments";
  }
  
  // End of Email Message
  EmailMessage += "</body></html>";
   
  // Send email to Administrator
  // MailApp.sendEmail(Address[0][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
  
  // If Error is between 0 and -60, send email to players. If not, only send to Administrator
  if (Status[0] >= -60){
    // Sends email to both players with the Match Data
    if (Address[1][0] == 'Français' && Address[1][1] != '') {
      MailApp.sendEmail(Address[1][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
    }
    if (Address[2][0] == 'Français' && Address[2][1] != ''&& Address[1][1] != Address[2][1]) {
      MailApp.sendEmail(Address[2][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
    }
  }
}


// **********************************************
// function subMatchReportTable()
//
// This function generates the HTML table that displays 
// the Match Data and Booster Pack Data
//
// **********************************************

function subMatchReportTable(EmailMessage, Headers, MatchData, Param){
  
  for(var row=0; row<=6; ++row){

    //if(row == 1) ++row;
    
    // Start of Match Table
    if(row == 0) EmailMessage += '<table style="border-collapse:collapse;" border = 1 cellpadding = 5><tr>';
    
    // Match Data
    if(row <= 6) {
      
      // Translate MatchData if necessary
      if (Param == 'EN' && MatchData[row][0] == 'Oui') MatchData[row][0] = 'Yes';
      if (Param == 'EN' && MatchData[row][0] == 'Non') MatchData[row][0] = 'No' ;
      if (Param == 'FR' && MatchData[row][0] == 'Yes') MatchData[row][0] = 'Oui';
      if (Param == 'FR' && MatchData[row][0] == 'No' ) MatchData[row][0] = 'Non';
            
      // If Match is NOT a Tie, Do not change anything
      if(MatchData[6][0] == 'No' || MatchData[6][0] == 'Non') {
        if(row < 6) EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][0]+'</td></tr>';
      }
      
      // If Match is a Tie, Change some Header Values
      if (MatchData[6][0] == 'Yes' || MatchData[6][0] == 'Oui') {
                
        if(row < 4) EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][0]+'</td></tr>';
        
        // If Match is Tie, replace Winning Player by Player 1
        if(row == 4 && Param == 'EN') EmailMessage += '<tr><td>Player 1</td><td>'+MatchData[row][0]+'</td></tr>';
        
        // If Match is Tie, replace Losing Player by Player 2
        if(row == 5 && Param == 'EN') EmailMessage += '<tr><td>Player 2</td><td>'+MatchData[row][0]+'</td></tr>';
        
        // If Match is Tie, replace Joueur Gagnant by Joueur 1
        if(row == 4 && Param == 'FR') EmailMessage += '<tr><td>Joueur 1</td><td>'+MatchData[row][0]+'</td></tr>';
        
        // If Match is Tie, replace Joueur Perdant by Joueur 2
        if(row == 5 && Param == 'FR') EmailMessage += '<tr><td>Joueur 2</td><td>'+MatchData[row][0]+'</td></tr>';
        
        if(row == 6) EmailMessage += '<tr><td>'+Headers[row][0]+'</td><td>'+MatchData[row][0]+'</td></tr>';
      }
    }
    
    // End of Table
    if(row == 6) EmailMessage += '</table><br>';   
  }
  return EmailMessage +'</table>';
}


// **********************************************
// function fcnSendFeedbackEmail()
//
// This function generates the feedback email 
//
// **********************************************

function fcnSendFeedbackEmail(shtConfig, Address, MatchData, Feedback) {
  
  // Variables
  var EmailSubject;
  var EmailMessage;
  
  // League Name
  var Location = shtConfig.getRange(11,2).getValue();
  var LeagueTypeEN = shtConfig.getRange(13,2).getValue();
  var LeagueNameEN = shtConfig.getRange(3,2).getValue() + ' ' + LeagueTypeEN;
    
  // Match Data Assignation
  var MatchID = MatchData[2][0];
  var Week    = MatchData[3][0];
  var Winr    = MatchData[4][0];
  var Losr    = MatchData[5][0];
  
  // Set Email Subject
  EmailSubject = Location + ' ' + LeagueNameEN + ' - Week ' + Week + ' - Player Feedback' ;
  
  // Start of Email Message
  EmailMessage = '<html><body>';
  
  EmailMessage += 'Match ID: ' + MatchID + '<br>' +
    'Week: ' + Week + '<br>' +
      'Winning Player: ' + Winr + '<br>' +
        'Losing Player: ' + Losr + '<br><br>';
  EmailMessage += 'Here is the feedback received by:<br><br>'+
    Address[1][1]+'<br>'+
      Address[2][1]+'<br><br>'+
        Feedback;
  
  // End of Email Message
  EmailMessage += '</body></html>';
  
  // Send email to Administrator
  MailApp.sendEmail(Address[0][1], EmailSubject, EmailMessage,{name:'Turn 1 Gaming League Manager',htmlBody:EmailMessage});
}


// **********************************************
// function subEmailPlayerPenaltyTable()
//
// This function generates the HTML table that displays 
// the Match Data and Booster Pack Data
//
// **********************************************

function subEmailPlayerPenaltyTable(PlayerData){
  
  var EmailMessage;
  
  for(var row=0; row<33; ++row){

    if(PlayerData[row][0] != ''){
      
      // Start of Table
      if(row == 0) {
        EmailMessage = 'Players who have not completed the minimum number of matches have received penalty losses on their record.<br>Here is the list of penalty losses this week.<br><br><font size="4"><b><table style="border-collapse:collapse;" border = 1 cellpadding = 5><tr>';
        EmailMessage += '<tr><td><b>Player Name</b></td><td><b>Penalty Losses</b></td></tr>';
      }
      
      // Player Data
      EmailMessage += '<tr><td>'+PlayerData[row][0]+'</td><td>'+PlayerData[row][1]+'</td></tr>';
    }
    if(PlayerData[row][0] == '') row = 33;
  }
  return EmailMessage +'</table>';
}

