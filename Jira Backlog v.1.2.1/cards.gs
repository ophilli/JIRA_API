
// START: Template functions
// You may need to update these functions if the template is changed.

function getTemplateArea() {
    return "A1:F10";
}

function setCardId(backlogItem, card) {
    card.getCell(1, 3).setValue(backlogItem['Key']);
    card.getCell(6, 3).setValue(backlogItem['Key']);
}

function setCardName(backlogItem, card) {
    card.getCell(2, 3).setValue(backlogItem['Summary']);
}

function setUserStory(backlogItem, card) {
    card.getCell(4, 3).setValue(backlogItem['User Story']);
}

function setHowToTest(backlogItem, card) {
    card.getCell(9, 3).setValue(backlogItem['Acceptance Criteria']);
}



function setEstimate(backlogItem, card) {
    card.getCell(1, 6).setValue(backlogItem['Story Points']);
}

function setPrerequisite(backlogItem, card) {
    card.getCell(7, 3).setValue(backlogItem['Prerequisites']);
}

function getTemplateStartColumn() {
    return getTemplateArea().substring(0,1);
}

function getTemplateStartRow() {
    return parseInt(getTemplateArea().substring(1,2), 10);
}

function getTemplateLastColumn() {
    return getTemplateArea().substring(3,4);
}

function getTemplateLastRow() {
    return parseInt(getTemplateArea().substring(4), 10);
}
// END: Template functions


// START: Get sheets
function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet(); 
}

function getBacklogSheet() {
    return getSpreadsheet().getActiveSheet();
}

function getTemplateSheet() {
    return getSpreadsheet().getSheetByName("Card Template");
}

function getCardSheet() {
    return getSpreadsheet().getSheetByName("Generated Cards");
}

function getPreparedCardSheet(template, numberOfItems, numberOfRows) {
    var rowsNeeded = numberOfItems * numberOfRows;

    var sheet = getCardSheet();
    sheet.clear();

    setColumnWidthTo(sheet, template);

    var rows = sheet.getMaxRows();

    if (rows < rowsNeeded) {
        sheet.insertRows(1, (rowsNeeded - rows));
    }

    setRowHeightTo(sheet, numberOfRows, numberOfItems);

    return sheet;
}
// END: Get sheets


// START: Get range within sheets
function getTemplateRange() {
    return getTemplateSheet().getRange(getTemplateArea());
}

function getHeadersRange(backlog) {
    return backlog.getRange(1, 1, 1, backlog.getLastColumn());
}

function getItemsRange(backlog) {
    var numRows = backlog.getLastRow() - 1;

    return backlog.getRange(2, 1, numRows, backlog.getLastColumn());
}

function getSelectedItemsRange(backlog) {
    var range = getSpreadsheet().getActiveRange();
    var startRow = range.getRowIndex();
    var rows = range.getNumRows();

    if (startRow < 2 ) { 
        startRow = 2; 
        rows = (rows > 1 ? rows-1 : rows);
    }

    return backlog.getRange(startRow, 1, rows, backlog.getLastColumn());
}
// END: Get range within sheets


function setRowHeightTo(cardSheet, numberOfRows, numberOfItems) {
    var templateSheet = getTemplateSheet();

    for (var i = 0; i < numberOfItems; i++) {
        for (var j = 1; j < (numberOfRows+1); j++) {
            var currentRow = (i*numberOfRows)+j;
            var currentHeight = templateSheet.getRowHeight(j);
            cardSheet.setRowHeight(currentRow, currentHeight);
        }
    }
}

function setColumnWidthTo(cardSheet, templateRange) {
    var templateSheet = getTemplateSheet();
    var max = templateRange.getLastColumn() + 1;

    for (var i = 1; i < max; i++) {
        var currentWidth = templateSheet.getColumnWidth(i);
        cardSheet.setColumnWidth(i, currentWidth);
    }
}

/* Get backlog items as objects with property name and values from the backlog. */
function getBacklogItems(selectedOnly) {
    var backlog = getBacklogSheet();

    var rowsRange = (selectedOnly ? getSelectedItemsRange(backlog) : getItemsRange(backlog));
    var rows = rowsRange.getValues();
    var headers = getHeadersRange(backlog).getValues()[0];

    var backlogItems = [];

    for (var i = 0; i < rows.length; i++) {
        var backlogItem = {};

        for (var j = 0; j < rows[i].length; j++) {
            backlogItem[headers[j]] = rows[i][j];
        }

        backlogItems.push(backlogItem);
    }

    return backlogItems;
}

function assertCardSheetExists() {
    if (getCardSheet() == null) {
        getSpreadsheet().insertSheet("Generated Cards", 2);
        Browser.msgBox("The 'Cards' sheet was missing and has now been added. Please try again.");
        return false;
    }

    return true;
}

function createCardsFromBacklog() {
    if (!assertCardSheetExists()) {
        return;
    }

    var backlogItems = getBacklogItems(false);
    createCards(backlogItems);
}

function createCardsFromSelectedRowsInBacklog() {
    if (!assertCardSheetExists()) {
        return;
    }

    if (getBacklogSheet().getName() != SpreadsheetApp.getActiveSheet().getName()) {
        Browser.msgBox("The Backlog sheet need to be active when creating cards from selected rows. Please try again.");
        return;
    }

    var backlogItems = getBacklogItems(true);
    createCards(backlogItems);
}

function createCards(backlogItems) {
    var numberOfRows = getTemplateLastRow();
    var template = getTemplateRange();
    var cardSheet = getPreparedCardSheet(template, backlogItems.length, numberOfRows);

    var startRow = getTemplateStartRow();    
    var lastRow = getTemplateLastRow();
    var startColumn = getTemplateStartColumn();
    var lastColumn = getTemplateLastColumn();

    for (var i = 0; i < backlogItems.length; i++) {
        var rangeVal = startColumn + startRow + ":" + lastColumn + lastRow;

        var card = cardSheet.getRange(rangeVal);
        template.copyTo(card);

        setCardId(backlogItems[i], card);
        setCardName(backlogItems[i], card);
        setUserStory(backlogItems[i], card);
        //setImportance(backlogItems[i], card);
        setHowToTest(backlogItems[i], card);
        setEstimate(backlogItems[i], card);
        //setPrerequisite(backlogItems[i], card);

        startRow += numberOfRows;
        lastRow += numberOfRows;
    }

    Browser.msgBox("Done!");
}
