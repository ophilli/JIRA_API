// ---------------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2014 Iain Brown - https://littlebmonkey.squarespace.com/blog/get-more-accurate-burndown-chart-for-jira
//
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:
//
// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

// Set these variables for you Jira Configuration
// --------------------------------------
var jiraHost = "lbmscrum.atlassian.net";
var projectPrefix = "EXA";
var sprintNumber = 47;
var authentication = "Basic SWFrasi5Ccm93bjpqa33rfc3Vjf4444f"; // This should be set to "Basic xxxxxxxxxxxxxxxx" where xxxxxxxxxxxx is the result of the base64 encoding of username:password for your jira login - e.g. Tom.Smith:ilovejira You can easily get he base64 encoding by visiting http://www.base64encode.org/
// --------------------------------------


function addTriggersIfNecessary() {
  
  var x = ScriptApp.getProjectTriggers();
  if (x.length == 0) {
    
    // Our standup is at 9.30am, so we have 2 goes at updating any estimates prior to the standup.
    
    ScriptApp.newTrigger("jiraRefresh").timeBased().everyDays(1).atHour(9).nearMinute(0).create();
    ScriptApp.newTrigger("jiraRefresh").timeBased().everyDays(1).atHour(9).nearMinute(25).create();
    sendEmail("Triggers Added for " + SpreadsheetApp.getActiveSpreadsheet().getName());
    
   
  }  
  
}  


// This is to automatically remove triggers after the sprint ends.
function removeTriggersIfNecessary() {
  
  var triggers = ScriptApp.getProjectTriggers();
  if (triggers.length > 0) {
    for (var y = 0; y < triggers.length; y++) {
    
      ScriptApp.deleteTrigger(triggers[y]);
   
    }
    sendEmail("Triggers removed for " + SpreadsheetApp.getActiveSpreadsheet().getName());
  }  
  
}  

function onOpen(){
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [{name: "Refresh Jira", functionName: "jiraRefresh"}]; 
  ss.addMenu("Jira", menuEntries);
  
}

// Call to initially load Jira into a spreadsheet.

function jiraRefresh() {
  
  addTriggersIfNecessary();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Stories");
  
  var range = ss.getRange(199, 1, 1,48);
  
  range.copyTo(ss.getRange(3, 1, 196,48));
  
  populateStoriesSheet();
  
}


function populateStoriesSheet() {
  
  var currentSS = SpreadsheetApp.getActiveSpreadsheet();
  // sprintNumber = currentSS.getName().split(' ')[1]; // This can override the sprint number to be the second token in the filename.
  
  ss = currentSS.getSheetByName("Stories");
  
  dates = ss.getRange(1, 4, 1, 40).getValues()[0];
  
  
  var issues = getIssuesForSprint(sprintNumber);
  var desc = new Array();
  for (var i=0;i<issues.length;i++) {
    desc.push(issues[i].key + ' ' + issues[i].fields.summary);
  }  
  var stories = new Array(); 
  var storyInfo = {key:new Array(),points:new Array(),dateAdded:new Array()};
  
  for (var i = 0; i < issues.length; i++) {
    
    var issue = issues[i];
    
    if (i == 0) {
      
      var sprintDetails = getSprintDetails(issue,sprintNumber);
      
      if (sprintDetails.endDate.substring(0,10) < new Date().toISOString().substring(0,10)) {
        removeTriggersIfNecessary();
      }  
    }  
    
    if (canProcessIssue(issue,sprintDetails.startDate)) {
      
     
      
      var desc = issue.key + ' - ' + issue.fields.summary;
      var time = issue.fields.timeoriginalestimate/3600;
      
      if (issue.fields.issuetype.name == 'Story') {
        if (issue.fields.summary.toUpperCase().indexOf("SWIMLANE") >= 0) {
          desc = "#swimlane";
        }
        else {
          storyInfo = addStoryInfo(storyInfo,issue,sprintDetails);
          desc = "#" + issue.key.replace(issue.fields.project.key + '-', '') + ' ' + issue.fields.summary;
        }  
        time = '';
      } 
      
       issue = getIssueTimeKeeping(issue,sprintDetails,storyInfo);
      
      
      if (issue.fields.issuetype.name == 'Story' || issue.fields.issuetype.name == 'Sub-task') {
        var row = [desc,time];
        stories.push(row);
      }
      else if (issue.fields.parent.fields.summary.toUpperCase().indexOf("SWIMLANE") >= 0) {
        desc = issue.fields.issuetype.name + ': ' + issue.fields.summary + ' (' + issue.key + ')' ;
        var row = [desc,time];
        stories.push(row);
      }  
      else {
        
        var row = ['*' + desc,time];
        stories.push(row);

      }  
      if (issue.workdates) {
        for (var w = 0; w < issue.workdates.dates.length;w++) {
          for (var d = 0; d < dates.length; d++) {
            var aDate = dates[d];
            if (aDate != "" && aDate.substring(0,10) == issue.workdates.dates[w]) {
              processHours(ss,issue.workdates.hoursSpent[w],issue.workdates.hoursRemaining[w],d,stories.length); 
              break;
            }   
            else if (aDate != "" && aDate.substring(0,10) > issue.workdates.dates[w]) {
              processHours(ss,issue.workdates.hoursSpent[w],issue.workdates.hoursRemaining[w],d-2,stories.length);  
              break;
            }   
          }
        }  
      }
    }
  }  
  
  var z = stories;
  ss.getRange(3, 1, stories.length, 2).setValues(stories);
  
  dateAdded = new Array();
  for (var i = 0; i < storyInfo.key.length;i++) {
    dateAdded.push([storyInfo.dateAdded[i]]);
  }  
  points = new Array();
  for (var i = 0; i < storyInfo.key.length;i++) {
    points.push([storyInfo.points[i]]);
  }  
  
  var review = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Review");
  review.getRange(4, 4, 25, 1).clearContent();
  review.getRange(4, 6, 25, 1).clearContent();
  review.getRange(4, 4, storyInfo.key.length, 1).setValues(points);
  review.getRange(4, 6, storyInfo.key.length, 1).setValues(dateAdded);
  
  
}  

function addStoryInfo(storyInfo,issue,sprintDetails) {
 
  storyInfo.key.push(issue.key);
  storyInfo.points.push(issue.fields.customfield_10004);
  storyInfo.dateAdded.push(getDateAdded(issue,sprintDetails));
  return storyInfo;
  
} 

function getDateAdded(issue,sprintDetails) {
  
   for (var i =0; i < issue.changelog.histories.length; i++) {
     var history = issue.changelog.histories[i];
    
     for (var j =0; j < history.items.length;j++) {
       var change = history.items[j];
       if (change.field == 'Sprint' && history.created.substring(0,10) > sprintDetails.startDate.substring(0,10)) {
         Logger.log("History %s, startdate %s, for issue %s",history.created,sprintDetails.startDate,issue.summary);
         return history.created.substring(0,10);
       }
     }  
   }

  return "";  
}   
  
  
  


function processHours(ss,hoursSpent,hoursRemaining,dateIndex,rowIndex) {
  
  if (hoursSpent >= 0) {
    var r = ss.getRange(2+rowIndex, 4+dateIndex);
    var current = r.getValue();
    r.setValue(r.getValue()+hoursSpent);
  }  
  
  if (hoursRemaining >= 0) {
    var r = ss.getRange(2+rowIndex, 5+dateIndex);
    r.setValue(hoursRemaining);
  }  
  
}  

function getIssueTimeKeeping(issue,sprintDetails,storyInfo) {
  
  var workdates = {dates:new Array(),hoursSpent:new Array(),hoursRemaining: new Array()};
  
  if (issue.fields.issuetype.name == "Story") {
    return issue;
  }  
  
  var index = storyInfo.key.indexOf(issue.fields.parent.key);
  if (index >= 0 && storyInfo.dateAdded[index] != "") {
    workdates = addToEntryForDate(workdates, storyInfo.dateAdded[index],-1,issue.fields.timeoriginalestimate/3600);
    workdates = addToEntryForDate(workdates, sprintDetails.startDate.substring(0,10),-1,0);
  }  
  issue.workdates = getWorklogSummaries(issue,workdates,sprintDetails,storyInfo);
  
  
  
  return issue;
  
}  

function addToEntryForDate(workdates,workdate,hoursSpent,hoursRemaining) {
  
  var i = workdates.dates.indexOf(workdate);
  
  if (i < 0) {
    workdates.dates.push(workdate);
    i = workdates.dates.length - 1;
    workdates.hoursSpent.push(0);
    workdates.hoursRemaining.push(-1);
  }  
  
  if (hoursSpent >= 0 && workdates.hoursSpent[i] < 0 ) {
    workdates.hoursSpent[i] = 0;
  }  
  if (hoursSpent >= 0) {
    workdates.hoursSpent[i] += hoursSpent;
  }
  
  if (hoursRemaining > -1 ) {
    workdates.hoursRemaining[i] = hoursRemaining;
  }
  return workdates;
  
}  

function getWorklogSummaries(issue,workdates,sprintDetails,storyInfo) {
  
  var worklogHoursRemaining = -1;
  var newHoursRemaining = -1;
  var newHoursSpent = -1;
  var worklogid = 0;
  var changeDate = "9999-12-31T00:00:00.000+1300";
  
  for (var i =0; i < issue.changelog.histories.length; i++) {
     Logger.log("New History");
     var history = issue.changelog.histories[i];
     newHoursRemaining = -1; 
     newHoursSpent = -1;
     worklogid = 0;
     changeDate = "9999-12-31T00:00:00.000+1300";
    
     for (var j =0; j < history.items.length;j++) {
       var change = history.items[j];
       Logger.log("New Change For History");
       if (change.field == 'WorklogId') {
         worklogid = change.from;
       }
       else if (change.field == "timeestimate") {
         newHoursRemaining = Math.round(change.to/3600 * 100) / 100;
         
         if (changeDate > history.created) {
           changeDate = history.created;
         }  
         Logger.log("Estimated time updated to %s for issue %s at %s with changeDate %s", newHoursRemaining,issue.fields.summary,history.created,changeDate);
       }  
       else if (change.field == "timespent") {
         newHoursSpent = Math.round((change.to-change.from)/3600* 100) / 100;
         
         if (changeDate > history.created) {
           changeDate = history.created;
         }
         Logger.log("Actual time updated to %s for issue %s at %s with changeDate %s", newHoursRemaining,issue.fields.summary,history.created,changeDate);
       } 
       else if (change.field == "status" && change.toString == "Resolved" && issue.fields.timeestimate != 0 && (issue.fields.status == "Resolved" || issue.fields.status == "Closed")) {
         newHoursRemaining = 0;
         if (changeDate > history.created) {
           changeDate = history.created;
         }  
         Logger.log("Resolved time updated to %s for issue %s at %s with changeDate %s", newHoursRemaining,issue.fields.summary,history.created,changeDate);
       }  
     }  
    
    if (worklogid != 0) {
      changeDate = getWorklogDate(issue,worklogid);
       Logger.log("ChangeDate changed to %s by worklog",changeDate);
    }
    
    if (worklogid != 0 && changeDate == '') {
      //do nothing if work log not found
    }  
    else if ((newHoursRemaining > -1 || newHoursSpent > -1) && (changeDate.substring(0,10) >= sprintDetails.startDate.substring(0,10) && changeDate.substring(0,10) <= sprintDetails.endDate.substring(0,10))) {
      Logger.log("Adding entry for date");
      workdates = addToEntryForDate(workdates,changeDate.substring(0,10),newHoursSpent,newHoursRemaining);
    }
    else {
      Logger.log("Ignored entry with nhr=%s,nhs=%s,cd=%s,st=%s,se=%s",newHoursRemaining,newHoursSpent,changeDate,sprintDetails.startDate,sprintDetails.endDate);
    }  
  }   
  
  return workdates;
}  
  


function getWorklogDate(issue,worklogid) {
  
  for (var w=0;w < issue.fields.worklog.worklogs.length; w++) {
    var worklog = issue.fields.worklog.worklogs[w];
    if (worklog.id == worklogid) {
      return worklog.started;
    }
  } 
  
  
  return '';
 
  
}