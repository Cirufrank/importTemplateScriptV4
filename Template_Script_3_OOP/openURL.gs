function openURL() {
  SpreadsheetApp.getUi()
   .showModalDialog(
     HtmlService.createHtmlOutputFromFile('myURL').setHeight(100),
     'Opening Script Logs'
   )
}
