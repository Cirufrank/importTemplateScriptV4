////////////////////////////////////////////////////////////
//Removes header (first row) comments from the sheet that it is called on
////////////////////////////////////////////////////////////

function removeHeaderCellComments() {
  let template = new Template();
  template.removeHeaderComments();
  SpreadsheetApp.getUi().alert("Header comments removed");
}
