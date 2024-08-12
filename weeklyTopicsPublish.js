function updateImagesInFormsFromMaster() {
    // The ID of your master form
    var masterFormId = '';
    // The ID of your Google Sheet containing the form IDs
    var sheetId = '';
    // Image titles to update
    var imageTitles = ['主題親證項目', '主題親證項目二', '親證秘密任務'];

    // Open the Google Sheet
    var spreadsheet = SpreadsheetApp.openById(sheetId);
    var formIdsSheet = spreadsheet.getSheetByName('Form IDs');
    var formIdsData = formIdsSheet.getDataRange().getValues();
    var formIds = formIdsData.slice(1).map(function(row) { return row[0]; });

    // Get the master form
    var masterForm = FormApp.openById(masterFormId);
    var masterItems = masterForm.getItems();

    // Find the master image items by titles
    var masterImageItems = {};
    imageTitles.forEach(function(imageTitle) {
        var masterImageItem = masterItems.find(function(item) {
            return item.getTitle() === imageTitle && item.getType() === FormApp.ItemType.IMAGE;
        });
        if (masterImageItem) {
            masterImageItems[imageTitle] = masterImageItem.asImageItem().getImage();
            Logger.log("master form get " + imageTitle );
        } else {
            Logger.log('Image item with title "' + imageTitle + '" not found in the master form.');
        }
    });
        // Update each form
    formIds.forEach(function(formId) {
        try {
            var form = FormApp.openById(formId);
            var items = form.getItems();

            imageTitles.forEach(function(imageTitle) {
                var imageItem = items.find(function(item) {
                    return item.getTitle() === imageTitle && item.getType() === FormApp.ItemType.IMAGE;
                });

                if (imageItem && masterImageItems[imageTitle]) {
                    // Update the image
                    imageItem.asImageItem().setImage(masterImageItems[imageTitle]);
                    Logger.log('Updated image "' + imageTitle + '" in form: ' + form.getPublishedUrl());
                } else if (!imageItem) {
                    Logger.log('Image item with title "' + imageTitle + '" not found in form ID ' + formId);
                }
            });
        } catch (e) {
            Logger.log('Error updating form ID ' + formId + ': ' + e.message);
        }
    });
}
