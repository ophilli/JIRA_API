function sendErrorMessage(message, useEmail) {
  
  if (!useEmail) {
    Browser.msgBox(message);
  }  
  else {
    var currentApp = 'Jira Sprint Report';
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), "Error With" + currentApp + ': ' + message, "");
  }  
  
}

function sendEmail(subject) {
  
    GmailApp.sendEmail(Session.getActiveUser().getEmail(), subject, "");
  
}

function testSendErrorMessage() {
  var x = new Date().toISOString()
  sendErrorMessage("It doesn't work",true);
  
}  
