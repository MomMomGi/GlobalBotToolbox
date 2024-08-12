function getWeeklyGroupScoresRankingAndNotify() {
    var mainSpreadsheetId = '';
    var mainSpreadsheet = SpreadsheetApp.openById(mainSpreadsheetId);
    var groupSheet = mainSpreadsheet.getSheetByName('ranking group');
    var data = groupSheet.getDataRange().getValues();
    
    var allScores = [];
    var weeklyScores = {};
    var groupTokens = {};

    // Iterate through each row in the 'ranking group' sheet
    for (var i = 1; i < data.length; i++) { // Start from 1 to skip the header row
        var groupName = data[i][0];
        var spreadsheetId = data[i][1];
        var lineToken = data[i][2];
        Logger.log("Parsed " + groupName + "\n");

        if (spreadsheetId) {
            var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
            var totalGroupScoresSheet = spreadsheet.getSheetByName('Total Group Scores');
            var weeklyGroupScoresSheet = spreadsheet.getSheetByName('Weekly Group Scores');
            
            if (totalGroupScoresSheet) {
                var totalScoresData = totalGroupScoresSheet.getRange('A2:B').getValues();
                for (var j = 0; j < totalScoresData.length; j++) {
                    if (totalScoresData[j][0]) {
                        allScores.push([groupName, totalScoresData[j][0]]);
                    }
                }
            }
            
            if (weeklyGroupScoresSheet) {
                var weeklyScoresData = weeklyGroupScoresSheet.getRange('A2:B').getValues();
                for (var j = 0; j < weeklyScoresData.length; j++) {
                    if (weeklyScoresData[j][0]) {
                        weeklyScores[groupName] = weeklyScoresData[j][0];
                        Logger.log("Weekly scores " + groupName + ": " + weeklyScoresData[j][0] + "\n");
                    }
                }
            }
        }
        
        // Store the line token for each group
        if (lineToken) {
            groupTokens[groupName] = lineToken;
        }
    }
    
    // Sort the scores by total points in descending order
    allScores.sort(function(a, b) {
        return b[1] - a[1];
    });
    
    // // Create the ranking message
    var messages = [];
    var currentMessage = '本週小隊排名:\n\n名次    小隊   總得分   本週得分\n';
    for (var k = 0; k < allScores.length; k++) {
        var groupName = allScores[k][0];
        var totalScore = allScores[k][1];
        var weeklyScore = weeklyScores[groupName] || 0;
        var newLine = (k + 1) + '.   ' + groupName + ': ' + totalScore + ' 分  (本週: ' + weeklyScore + ' 分)\n';
        
        // Check if adding this line would exceed the character limit
        if (currentMessage.length + newLine.length > 900) { // Leave some buffer
            messages.push(currentMessage);
            currentMessage = '本週小隊排名:\n\n名次    小隊   總得分   本週得分\n';
        }
        
        currentMessage += newLine;
    }
    
    if (currentMessage.length > 0) {
        messages.push(currentMessage);
    }
    
    // Send the ranking messages to each group via LINE Notify
    for (var group in groupTokens) {
        for (var m = 0; m < messages.length; m++) {
            sendLineNotify(messages[m], groupTokens[group]);
        }
    }
}

function calculateTop3Players() {
    var mainSpreadsheetId = ''; // Replace with your main spreadsheet ID
    var mainSpreadsheet = SpreadsheetApp.openById(mainSpreadsheetId);
    var groupSheet = mainSpreadsheet.getSheetByName('ranking group');
    var data = groupSheet.getDataRange().getValues();
    
    var allScores = [];
    // var weeklyScores = {};
    var groupTokens = {};

    for (var i = 1; i < data.length; i++) { // Start from 1 to skip the header row
        var groupName = data[i][0];
        var spreadsheetId = data[i][1];
        var lineToken = data[i][2];
        Logger.log("Parsed " + groupName + "\n");

        if (spreadsheetId) {
            var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
            var totalSummarySheet = spreadsheet.getSheetByName('Total Summary');
            
            if (totalSummarySheet) {
                var totalSummaryData = totalSummarySheet.getRange('A3:B').getValues();
                for (var j = 0; j < totalSummaryData.length; j++) {
                    if (totalSummaryData[j][0]) {
                        allScores.push([groupName, totalSummaryData[j][0], totalSummaryData[j][1]]);
                    }
                }
            }
        }
        
        // Store the line token for each group
        if (lineToken) {
            groupTokens[groupName] = lineToken;
        }
    }

    // Sort the scores by points in descending order
    allScores.sort(function(a, b) {
        return b[2] - a[2];
    });

    // Extract the top 3 players
    var top3Players = allScores.slice(0, 10);

    // Create the ranking message
    var message = 'Top 3 Players:\n\n';
    for (var k = 0; k < top3Players.length; k++) {
        message += (k + 1) + '. ' + top3Players[k][1] + '(' + top3Players[k][0] + ')' +': ' + top3Players[k][2] + ' points\n';
    }

    // Log the message
    Logger.log(message);

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
    Logger.log('Error sending LINE notification: ' + err);
  }
}
