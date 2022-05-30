class NeedAndOpportunityResponsesCheck extends UserTemplate {
  constructor() {
    super();
  }
}

try {
  function needAndOpportunityResponsesCheck() {
      let responseTemplate = new NeedAndOpportunityResponsesCheck();
      let reportSheet =responseTemplate.createReportSheet();

      responseTemplate.sheet.setName(responseTemplate.templateName);
      

      responseTemplate.ss.setActiveSheet(responseTemplate.getSheet(responseTemplate.templateName));


      responseTemplate.setFrozenRows(responseTemplate.sheet, 1);
      responseTemplate.setFrozenRows(reportSheet, 1);

      try {
            responseTemplate.removeWhiteSpaceFromCells();
          } catch (err) {
            Logger.log(err);
            responseTemplate.reportSummaryComments = responseTemplate.failedRemovedWhiteSpaceFromCellsComment;
            throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }

      try {
        responseTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          responseTemplate.reportSummaryComments = responseTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

      responseTemplate.clearSheetSummaryColumn(responseTemplate.sheet);
      responseTemplate.clearSheetSummaryColumn(reportSheet);
      responseTemplate.setErrorColumnHeaderInMainSheet();

      try {
        responseTemplate.formatAllDatedColumns(reportSheet);
      } catch(err) {
          Logger.log(err);
          responseTemplate.reportSummaryComments = responseTemplate.failedFormatDateColumns;
          throw new Error(`Check not ran for formatting of the dated column(s). Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
      
      try {
        responseTemplate.checkForInvalidEmails(reportSheet);
      } catch(err) {
          Logger.log(err);
          responseTemplate.reportSummaryComments = responseTemplate.failedInvalidEmailCheckMessage;
          throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        } 

      try{
        responseTemplate.checkFirstThreeColumnsForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            responseTemplate.reportSummaryComments = responseTemplate.failedCheckFirstThreeColumnsForMissingValues;
            throw new Error(`Check not ran for missing values within first three columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first five colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for five]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
          }

      try {
        responseTemplate.setCommentsOnReportCell(reportSheet);
      } catch(err) {
          Logger.log(err);
          throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
      
      try {
        responseTemplate.checkForInvalidNumbers(reportSheet);
      } catch(err) {
          Logger.log(err);
          responseTemplate.reportSummaryComments = responseTemplate.failedInvalidPhoneNumbersCheck;
          throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }

      SpreadsheetApp.getUi().alert("Responses Import Check Complete");
  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the responses import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
