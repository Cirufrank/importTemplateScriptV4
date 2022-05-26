function removeHeaderCellComments() {
    let template = new Template();
    template.removeHeaderComments();
    SpreadsheetApp.getUi().alert("Header comment removed");
  }
  