function sendFormResponsesToLineNotify(data) {
    var {spreadsheetId, lineToken} = getSheetAndToken("");
    var message = "";
    var totalSum = 0; // Initialize the sum variable
    var hasError = false; // Flag to track if there was an error

    try {
        var formResponses = data.response.getItemResponses();
        for (var i = 0; i < formResponses.length; i++) {
            var item = formResponses[i].getItem().getTitle();
            var response = formResponses[i].getResponse();

            if (item.startsWith("[計分項目]")) {
                totalSum += parseFloat(response) || 0;
            }
            if (item === "小組隊員") {
                message += response;
            }
            // message += item + " ： " + response + "\n";
        }
        // Add the total sum to the message
        message += " 本次回報得分為：" + totalSum + "\n";
    } catch (err) {
        message += "No Answers for message.\n";
        Logger.log(err);
        hasError = true; // Set the flag to indicate an error occurred
    }

    if (!hasError) {
        sendLineNotify(message, lineToken);
    }
}

function getSheetAndToken(groupId) {
    // Open the spreadsheet containing the 'group sheet'
    var mainSpreadsheetId = '';
    var mainSpreadsheet = SpreadsheetApp.openById(mainSpreadsheetId);
    var groupSheet = mainSpreadsheet.getSheetByName('group sheet');

    // Get the data from the 'group sheet'
    var data = groupSheet.getDataRange().getValues();
    var {spreadsheetId, lineToken} = getSheetIDAndToken(data, groupId);
    return {spreadsheetId, lineToken};
}

function getSheetIDAndToken(data, groupId) {
    var spreadsheetId = null;
    var lineToken = null;
    for (var i = 1; i < data.length; i++) { // Start from 1 to skip the header row
        if (data[i][0] == groupId) {
            spreadsheetId = data[i][1];
            lineToken = data[i][2];
            break;
        }
    }
    return {spreadsheetId, lineToken};
}

function sendSummary(summary, summaryDesc) {
    var {spreadsheetId, lineToken} = getSheetAndToken("");

    // Open the target spreadsheet
    var targetSpreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var dailySummarySheet = targetSpreadsheet.getSheetByName(summary);

    if (dailySummarySheet) {
        var data = dailySummarySheet.getDataRange().getValues();
        var message = summaryDesc + ":\n";

        for (var i = 2; i < data.length; i++) { // Start from 1 to skip the header row
            message += data[i][0] + ": " + data[i][1] + " 分\n";
        }

        // Send LINE notification
        sendLineNotify(message, lineToken);
    } else {
        Logger.log('Daily Summary sheet not found in the target spreadsheet.');
    }
}

function sendDailySummary() {
    var summary = 'Daily Summary';
    var summaryDesc = "每日統計";
    sendSummary(summary, summaryDesc);
}

function sendWeeklySummary() {
    var summary = 'Weekly Summary';
    var summaryDesc = "每週統計";
    sendSummary(summary, summaryDesc);
}

function sendLineNotify(message, token) {
    var options = {
        "method": "post",
        "payload": {
            "message": message
        },
        "headers": {
            "Authorization": "Bearer " + token
        }
    };
    try {
        UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
    } catch (err) {
        // 異常時進行紀錄
        Logger.log(err)
    }
}
