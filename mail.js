function onOpen(e) {
    SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Show documentation', 'showSidebar')
    .addToUi();
}

function onInstall(e) {
    onOpen(e);
}

function showSidebar() {
    var ui = HtmlService 
    .createHtmlOutputFromFile('Sidebar')
    .setTitle('Finance Functions');
    SpreadsheetApp.getUi().showSidebar(ui);
}