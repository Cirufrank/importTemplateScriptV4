////////////////////////////////////////////////////////////
//Opens a dialogue box that will direct users to the script termination page when the Script Termination Page button is clicked
////////////////////////////////////////////////////////////

function openURL() {
    SpreadsheetApp.getUi()
    //shows a modal dialogue box outputting the specified HTML
     .showModalDialog(
       //creates HTML output from the myURL.html file
       HtmlService.createHtmlOutputFromFile('myURL').setHeight(100),
       'Opening Script Logs'
     )
  }
  