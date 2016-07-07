
 


function canProcessIssue(issue,startDate) {
  
  // If it's been updated to resolved prior to start of sprint, then ignore
  if (issue.fields.status.name == 'Resolved' || issue.fields.status.name == 'Closed' ) {
    if (issue.fields.updated < startDate) {
      Logger.log("Ignore %s from %s",issue.fields.summary,issue.fields.parent.fields.summary);
      return false;
    }  
  }

  return true;  
  
}  

function getJQLResults(jql) {
  
  var url = "https://" + jiraHost + "/rest/api/2/search?jql=" + jql + "&fields=*all&expand=changelog&maxResults=1000" 

  
  
  var headers = { "Accept":"application/json", 
              "Content-Type":"application/json", 
              "method": "GET",
              "headers": {"Authorization": authentication}
             };
  
  var resp = UrlFetchApp.fetch(url,headers );
  
  var x=resp.getContentText();

  return Utilities.jsonParse(x);
  
}

function getSprintDetails(issue,sprintNumber) {
  
  var sprintDetails = {startDate:'', endDate:''};

   for (var j = 0; j < issue.fields.customfield_10007.length; j++) {
    
      var sprint = issue.fields.customfield_10007[j];
      
      if (sprint.indexOf("name=Sprint " + sprintNumber) > 0) {
        
        var endDatePos = sprint.indexOf("endDate=");
        if (endDatePos > 0) {
          sprintDetails.endDate = sprint.substr(endDatePos+8,29);
        }
        
        var startDatePos = sprint.indexOf("startDate=");
        if (startDatePos > 0) {
          sprintDetails.startDate = sprint.substr(startDatePos+10,29);
        }                                             
                                                                                      
        return sprintDetails;
         
        
      }  
   } 
  
}

function getIssuesForSprint(sprintNumber) {
  
  var data = getJQLResults("project%20%3D%20" + projectPrefix + "%20and%20sprint%20%3D%20%22Sprint%20" + sprintNumber + "%22" + '%20order%20by%20rank');
  
  return data.issues;
  
}  

