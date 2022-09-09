////////////////////////////////////////////////////////////
//Runs the checks for the User Groups Import Template
////////////////////////////////////////////////////////////

try {
    //Executes when the Check User Groups Template button is clicked
    function userGroupsTemplateCheck() {
      const userGroupsTemplate = new UserGroupsTempalte();
      const reportSheet = userGroupsTemplate.createReportSheet();
  
      userGroupsTemplate.sheet.setName(userGroupsTemplate.templateName);
      
  
     userGroupsTemplate.ss.setActiveSheet(userGroupsTemplate.getSheet(userGroupsTemplate.templateName));
  
      
  
  
      userGroupsTemplate.setFrozenRows(userGroupsTemplate.sheet, 1);
      userGroupsTemplate.setFrozenRows(reportSheet, 1);
  
     try {
            userGroupsTemplate.removeWhiteSpaceFromCells();
          } catch (err) {
            Logger.log(err);
            userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedRemovedWhiteSpaceFromCellsComment;
            throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
  
          try {
            userGroupsTemplate.removeFormattingFromSheetCells();
          } catch (err) {
            Logger.log(err);
            userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedRemovedFormattingFromCellsComment;
            throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
      
      userGroupsTemplate.clearSheetSummaryColumn(userGroupsTemplate.sheet);
      userGroupsTemplate.clearSheetSummaryColumn(reportSheet);
      userGroupsTemplate.setErrorColumnHeaderInMainSheet(userGroupsTemplate.sheet);
  
      try{
        userGroupsTemplate.checkFirstColumnForBlanks(reportSheet);
          } catch(err) {
              Logger.log(err);
              userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedFirstColumnMissingValuesCheckMessage;
              throw new Error(`Check not ran for missing values within first column. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
  
      try {
        userGroupsTemplate.formatCommaSeparatedLists(reportSheet);
      } catch (err) {
          Logger.log(err);
          userGroupsTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
          throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
  
      try {
      userGroupsTemplate.checkForInvalidCommaSeparatedEmails(reportSheet);
      } catch(err) {
          Logger.log(err);
          userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedCheckNotRanForInvalidCommaSeparatedEmail;
          throw new Error(`Check not ran for invalid User Group Members emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the user group members column is titled "User Group Members (emails)" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        } 
  
      try {
    userGroupsTemplate.checkForInvalidAllowedDomainNames(reportSheet);
    } catch(err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedCheckNotRanForInvalidAllowedEmailDomains;
        throw new Error(`Check not ran for invalid allowed domain names. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the allowed domain names column is titled "Allowed Email Domains" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
  
    try {
      userGroupsTemplate.setCommentsOnReportCell(reportSheet);
    } catch(err) {
      Logger.log(err);
      throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
      //Has a pop-up tell the user the check has been completed
     SpreadsheetApp.getUi().alert("User Groups Import Check Complete");
    }
  
  } catch (err) {
      Logger.log(err);
      throw new Error(`An error occured and the User Groups import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    
  