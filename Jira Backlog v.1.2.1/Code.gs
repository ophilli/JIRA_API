// ---------------------------------------------------------------------------------------------------------------------------------------------------
// The MIT License (MIT)
// 
// Copyright (c) 2014 Iain Brown - http://www.littlebluemonkey.com/blog/automatically-import-jira-backlog-into-google-spreadsheet
// 
//
// Inspired by http://gmailblog.blogspot.co.nz/2011/07/gmail-snooze-with-apps-script.html
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
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.

var C_MAX_RESULTS = 1000;

// Generates JIRA Backlog control menu
function onOpen(e){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var menuEntries = [{name: "Configure Jira", functionName: "jiraConfigure"},
                       {name: "Refresh Data Now", functionName: "jiraPullManual"},
                       {name: "Schedule Automatic Refresh", functionName: "scheduleRefresh"},
                       {name: "Stop Automatic Refresh", functionName: "removeTriggers"}]; 
  
    ss.addMenu("Jira", menuEntries);

    //menuEntries = [{name: "Create cards", functionName: "createCardsFromBacklog"}, 
    //               {name: "Create cards from selected rows", functionName: "createCardsFromSelectedRowsInBacklog"}];
  
    //ss.addMenu("Story Cards", menuEntries);
}

// Configures constants for JIRA Backlog
function jiraConfigure() {
    
    var prefix = Browser.inputBox("Enter the 3-4 digit prefix for your Jira Project. e.g. RCPQ", "Prefix", Browser.Buttons.OK);
    PropertiesService.getUserProperties().setProperty("prefix", prefix.toUpperCase());
    
    var host = Browser.inputBox("Enter the host name of your on demand instance e.g. revvy-modeln.atlassian.net", "Host", Browser.Buttons.OK);
    PropertiesService.getUserProperties().setProperty("host", host);
    
    var userAndPassword = Browser.inputBox("Enter your Jira On Demand User id and Password in the form User:Password. e.g. Tommy.Smith:ilovejira (Note: This will be base64 Encoded and saved as a property on the spreadsheet)", "Userid:Password", Browser.Buttons.OK_CANCEL);
    var x = Utilities.base64Encode(userAndPassword);
    PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
    
    var issueTypes = Browser.inputBox("Enter a comma separated list of the types of issues you want to import    e.g. story or story,epic,bug", "Issue Types", Browser.Buttons.OK);
    PropertiesService.getUserProperties().setProperty("issueTypes", issueTypes);


    Browser.msgBox("Jira configuration saved successfully.");
}    

// Adds automatic refresh triggers
function scheduleRefresh() {
    var time = Browser.inputBox("Enter desired time between refreshes in hours. e.g. 4", "Time", Browser.Buttons.OK)
  
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    
    ScriptApp.newTrigger("jiraPull").timeBased().everyHours(time).create();
    
    Browser.msgBox("Spreadsheet will refresh automatically every " + time + " hours.");
}    

// Removes automatic refresh triggers
function removeTriggers() {
    var triggers = ScriptApp.getProjectTriggers();
    for (var i = 0; i < triggers.length; i++) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
    
    Browser.msgBox("Spreadsheet will no longer refresh automatically.");
}    

// Manually refreshes JIRA data
function jiraPullManual() {
    jiraPull();
    Browser.msgBox("Jira backlog successfully imported");
}    

// Manages and processes requests from API
function jiraPull() {
    
    var allFields = getAllFields();
    
    var data = getStories();
    
    if (allFields === "" || data === "") {
        Browser.msgBox("Error pulling data from Jira - aborting now.");
        return;
    }    
    
    var ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backlog");
    var headings = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
    
    
    
    var y = new Array();
    for (i=0;i<data.issues.length;i++) {
        var d=data.issues[i];
        y.push(getStory(d,headings,allFields));
    }    
    
    ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Backlog");
    var last = ss.getLastRow();
    if (last >= 2) {
        ss.getRange(2, 1, ss.getLastRow()-1,ss.getLastColumn()).clearContent();    
    }    
    
    if (y.length > 0) {
        ss.getRange(2, 1, data.issues.length,y[0].length).setValues(y);
    }
}


function getAllFields() {
    
    var theFields = getFields();
    var allFields = new Object();
    allFields.ids = new Array();
    allFields.names = new Array();
    
    for (var i = 0; i < theFields.length; i++) {
            allFields.ids.push(theFields[i].id);
            allFields.names.push(theFields[i].name.toLowerCase());
    }    
    
    return allFields;
}    


function getFields() {
    return JSON.parse(getDataFromAPI("field"));  
}    

// Parses API data 
function getStories() {
    var allData = {issues:[]};
    var data = {startAt:0,maxResults:0,total:1};
    var startAt = 0;
    
    while (data.startAt + data.maxResults < data.total) {
        Logger.log("Making request for %s entries", C_MAX_RESULTS);
        data = JSON.parse(getDataFromAPI("search?jql=project%20%3D%20" 
                                        + PropertiesService.getUserProperties().getProperty("prefix") 
                                        + "%20and%20status%20!%3D%20resolved%20and%20type%20in%20("
                                        + PropertiesService.getUserProperties().getProperty("issueTypes") 
                                        + ")%20order%20by%20rank%20&maxResults=" 
                                        + C_MAX_RESULTS + "&startAt=" + startAt));

        allData.issues = allData.issues.concat(data.issues);
        startAt = data.startAt + data.maxResults;
    }    
    
    return allData;
}    

// JIRA API Call
function getDataFromAPI(path) {
    var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
    var digestfull = PropertiesService.getUserProperties().getProperty("digest");
    
    var headers = { "Accept":"application/json", 
                            "Content-Type":"application/json", 
                            "method": "GET",
                            "headers": {"Authorization": digestfull},
                            "muteHttpExceptions": true
                            };
    
    var resp = UrlFetchApp.fetch(url,headers );
    if (resp.getResponseCode() != 200) {
        Browser.msgBox("Error retrieving data for url" + url + ":" + resp.getContentText());
        return "";
    }    
    else {
        return resp.getContentText();
    }    
}    


function getStory(data,headings,fields) {
    
    var story = [];
    for (var i = 0;i < headings.length;i++) {
        if (headings[i] !== "") {
            story.push(getDataForHeading(data,headings[i].toLowerCase(),fields));
        }    
    }                
    
    return story;
}    


function getDataForHeading(data,heading,fields) {
    
            if (data.hasOwnProperty(heading)) {
                return data[heading];
            }    
            else if (data.fields.hasOwnProperty(heading)) {
                return data.fields[heading];
            }    
    
            var fieldName = getFieldName(heading,fields);
    
            if (fieldName !== "") {
                if (data.hasOwnProperty(fieldName)) {
                    return data[fieldName];
                }    
                else if (data.fields.hasOwnProperty(fieldName)) {
                    return data.fields[fieldName];
                }    
            }
    
            var splitName = heading.split(" ");
    
            if (splitName.length == 2) {
                if (data.fields.hasOwnProperty(splitName[0]) ) {
                    if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
                        return data.fields[splitName[0]][splitName[1]];
                    }
                    return "";
                }    
            }    
    
            return "Could not find value for " + heading;
}    

// 
function getFieldName(heading,fields) {
    var index = fields.names.indexOf(heading);
    if ( index > -1) {
         return fields.ids[index]; 
    }
    return "";
}
