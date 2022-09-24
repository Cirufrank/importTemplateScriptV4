////////////////////////////////////////////////////////////
//Runs the checks for the Hours/Reponses Import Template
////////////////////////////////////////////////////////////

try {
  //Executes when the Check Responses and Hours Template button is clicked
  function hoursAndResponsesCheck() {
    const responsesAndHoursTemplate = new HoursAndResponsesTemplate();
    const reportSheet = responsesAndHoursTemplate.createReportSheet();

    responsesAndHoursTemplate.sheet.setName(responsesAndHoursTemplate.templateName);
    

    responsesAndHoursTemplate.ss.setActiveSheet(responsesAndHoursTemplate.getSheet(responsesAndHoursTemplate.templateName));


    responsesAndHoursTemplate.setFrozenRows(responsesAndHoursTemplate.sheet, 1);
    responsesAndHoursTemplate.setFrozenRows(reportSheet, 1);

    try {
            responsesAndHoursTemplate.removeWhiteSpaceFromCells();
          } catch (err) {
            Logger.log(err);
            responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedRemovedWhiteSpaceFromCellsComment;
            throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }

          try {
          responsesAndHoursTemplate.removeFormattingFromSheetCells();
          } catch (err) {
            Logger.log(err);
            responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedRemovedFormattingFromCellsComment;
            throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
    
          responsesAndHoursTemplate.clearSheetSummaryColumn(responsesAndHoursTemplate.sheet);
          responsesAndHoursTemplate.clearSheetSummaryColumn(reportSheet);
          responsesAndHoursTemplate.setErrorColumnHeaderInMainSheet();

          try {
          responsesAndHoursTemplate.formatAllDatedColumns(reportSheet);
        } catch(err) {
          Logger.log(err);
          responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedFormatDateColumns;
          throw new Error(`Check not ran for formatting of the dated column(s). Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
    
        try {
          responsesAndHoursTemplate.checkForInvalidEmails(reportSheet);
        } catch(err) {
            Logger.log(err);
            responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedInvalidEmailCheckMessage;
            throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
          } 

          try{
      responsesAndHoursTemplate.checkFirstFiveColumnsForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedFirstFiveColumnsMissingValuesCheck;
            throw new Error(`Check not ran for missing values within first five columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first five colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for five]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

          try {
          responsesAndHoursTemplate.setCommentsOnReportCell(reportSheet);
        } catch(err) {
          Logger.log(err);
          throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
        //Has a pop-up tell the user the check has been completed
        SpreadsheetApp.getUi().alert("Responses and Hours Import Check Complete");

  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the responses and hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

