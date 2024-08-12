function getGroupTitles(spreadsheet, groupSheet) {
    var formTitlesSheet = spreadsheet.getSheetByName(groupSheet);
    var formTitlesData = formTitlesSheet.getDataRange().getValues();
    var formTitles = formTitlesData.map(function (row) {
        return "超級創造班" + row[0] + " 回報表單" ;
    });
    return formTitles;
}

function getMemberLists(spreadsheet, memberSheet) {
    var memberNamesSheet = spreadsheet.getSheetByName(memberSheet);
    var memberNamesData = memberNamesSheet.getDataRange().getValues();
    var memberNames = transpose(memberNamesData.slice(1)); // Skip the header row and transpose
    return memberNames;
}

function copyItemsFromMaster(masterItems, newForm, memberNames, i) {
    var memberDropdown = null;
    masterItems.forEach(function (item, index) {
        switch (item.getType()) {
            case FormApp.ItemType.TEXT:
                var textItem = item.asTextItem();
                newForm.addTextItem()
                    .setTitle(textItem.getTitle())
                    .setHelpText(textItem.getHelpText())
                    .setRequired(textItem.isRequired());
                break;
            case FormApp.ItemType.MULTIPLE_CHOICE:
                var mcItem = item.asMultipleChoiceItem();
                var newMcItem = newForm.addMultipleChoiceItem()
                    .setTitle(mcItem.getTitle())
                    .setChoices(mcItem.getChoices())
                    .setHelpText(mcItem.getHelpText())
                    .setRequired(mcItem.isRequired());
                if (mcItem.hasOtherOption()) {
                    newMcItem.showOtherOption(true);
                }
                break;
            case FormApp.ItemType.LIST:
                var listItem = item.asListItem();
                newForm.addListItem()
                    .setTitle(listItem.getTitle())
                    .setChoiceValues(listItem.getChoiceValues())
                    .setHelpText(listItem.getHelpText())
                    .setRequired(listItem.isRequired());
                break;
            case FormApp.ItemType.DATE:
                var dateItem = item.asDateItem();
                newForm.addDateItem()
                    .setTitle(dateItem.getTitle())
                    .setHelpText(dateItem.getHelpText())
                    .setRequired(dateItem.isRequired());
                break;
            case FormApp.ItemType.CHECKBOX:
                var checkboxItem = item.asCheckboxItem();
                newForm.addCheckboxItem()
                    .setTitle(checkboxItem.getTitle())
                    .setChoices(checkboxItem.getChoices())
                    .setHelpText(checkboxItem.getHelpText())
                    .setRequired(checkboxItem.isRequired());
                break;
            case FormApp.ItemType.IMAGE:
                var imageItem = item.asImageItem();
                var newImage = newForm.addImageItem();
                newImage.setImage(imageItem.getImage())
                    .setTitle(imageItem.getTitle())
                    .setHelpText(imageItem.getHelpText())
                    .setAlignment(imageItem.getAlignment());
                // Insert the "member name" dropdown right after the image
                if (imageItem.getTitle() === '加值規則' && !memberDropdown) {
                    memberDropdown = newForm.addListItem();
                    memberDropdown.setTitle('小組隊員');
                    memberDropdown.setChoiceValues(memberNames[i]);
                    memberDropdown.setRequired(true);
                    newForm.moveItem(memberDropdown.getIndex(), newImage.getIndex() + 1);
                }
                break;
            // Add cases for other item types as needed
            default:
                Logger.log('Unsupported item type: ' + item.getType());
        }
    });
}

function createResponseSheet(formTitle, newForm) {
    // Create a new spreadsheet for responses
    var responseSpreadsheet = SpreadsheetApp.create(formTitle + ' Responses');
    var responseSpreadsheetId = responseSpreadsheet.getId();

    // Link the form to the new spreadsheet
    newForm.setDestination(FormApp.DestinationType.SPREADSHEET, responseSpreadsheetId);
    return {responseSpreadsheet, responseSpreadsheetId};
}

function getFormIdSheet(spreadsheet) {
    var formIdsSheet = spreadsheet.getSheetByName('Form IDs');
    if (!formIdsSheet) {
        formIdsSheet = spreadsheet.insertSheet('Form IDs');
        formIdsSheet.appendRow(['Form ID', 'Spreadsheet ID', 'Form URL', 'Sheet URL']);
    }
    return formIdsSheet;
}

function createFromAndTitle(i, formTitles, memberNames) {
    // create the form with title
    var newForm = FormApp.create(formTitles[i]);
    // set the form description to include member names
    var newFormId = newForm.getId();
    var form = FormApp.openById(newFormId);

    var teamLeader = memberNames[i][0];
    var teamMembers = memberNames[i].slice(1).join(', ');
    var memberNamesDescription = '組長: ' + teamLeader + '\n組員: ' + teamMembers;
    form.setDescription(memberNamesDescription);
    return newForm;
}

function createFormsBatch() {
    // The ID of your Google Sheet containing the form titles and member names
    var sheetId = '';
    // var sheetId = '1tnyHkR1NCW00j0nokKhBxsf5KV6FGn5GT1wXJGqNx4c';
    var batchSize = 5; // Number of forms to create per batch

    // Open the Google Sheet
    var spreadsheet = SpreadsheetApp.openById(sheetId);

    // Read the form titles from the "Form Titles" sheet
    const groupSheet = 'group names';
    var formTitles = getGroupTitles(spreadsheet, groupSheet);

    // Read the member names from the "Member Names" sheet
    const memberSheet = 'member names';
    var memberNames = getMemberLists(spreadsheet, memberSheet);

    // Ensure the lengths of formTitles and memberNames are equal
    if (formTitles.length !== memberNames.length) {
        Logger.log('The number of titles and member name arrays must be the same.');
        return;
    }

    // Read the last processed index from the "State" sheet
    var stateSheet = spreadsheet.getSheetByName('State');
    if (!stateSheet) {
        stateSheet = spreadsheet.insertSheet('State');
        stateSheet.getRange('A1').setValue(0);
    }
    var lastProcessedIndex = stateSheet.getRange('A1').getValue();

    // The ID of your master form
    var masterFormId = '';

    // Get the master form
    var masterForm = FormApp.openById(masterFormId);
    var masterItems = masterForm.getItems();

    // Get the sheet to store form and spreadsheet IDs
    var formIdsSheet = getFormIdSheet(spreadsheet);

    // Create new forms in batches
    for (var i = lastProcessedIndex; i < lastProcessedIndex + batchSize && i < formTitles.length; i++) {
        var newForm = createFromAndTitle(i, formTitles, memberNames);

        // Copy items from the master form
        copyItemsFromMaster(masterItems, newForm, memberNames, i);

        // create and link the response sheet
        var {responseSpreadsheet, responseSpreadsheetId} = createResponseSheet(formTitles[i], newForm);

        // Log the URLs of the new form and spreadsheet
        Logger.log('Created form: ' + newForm.getPublishedUrl());
        Logger.log('Response spreadsheet: ' + responseSpreadsheet.getUrl());

        // Store the form ID and spreadsheet ID in the sheet
        formIdsSheet.appendRow([newForm.getId(), responseSpreadsheetId, newForm.getPublishedUrl(), responseSpreadsheet.getUrl()]);
    }

    // Update the last processed index
    stateSheet.getRange('A1').setValue(lastProcessedIndex + batchSize);
}

function transpose(matrix) {
    return matrix[0].map((_, colIndex) => matrix.map(row => row[colIndex]));
}

