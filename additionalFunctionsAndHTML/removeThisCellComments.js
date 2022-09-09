////////////////////////////////////////////////////////////
//Removes header (first row) comments from the sheet that it is called on
////////////////////////////////////////////////////////////

//Executes when the Remove Cell Comments button is clicked
try {
    function removeCellComments() {
    const template = new Template();
    template.removeCellComments();
    //Has a pop-up tell the user the header comments have been removed
    SpreadsheetApp.getUi().alert("Cell comments removed");
  }
} catch (err) {
  Logger.log(err);
  throw new Error(`An error occured and the cell comments remover did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
}
