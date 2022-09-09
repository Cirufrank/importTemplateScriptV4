////////////////////////////////////////////////////////////
//Runs the checks for the User Import Template
////////////////////////////////////////////////////////////

try {
    //Executes when the User Import Template Checker is clicked
    function checkUserImportTemplate() {
    const userImportTemplate = new UserTemplate();
    
    const reportSheet = userImportTemplate.createReportSheet();
    
    //Makes the main sheet the active sheet so that actions are performed on the non-report sheet
    userImportTemplate.ss.setActiveSheet(userImportTemplate.getSheet(userImportTemplate.templateName));
   
  
    userImportTemplate.setFrozenRows(userImportTemplate.sheet, 1);
    userImportTemplate.setFrozenRows(reportSheet, 1);
  
    
    try {
      userImportTemplate.removeWhiteSpaceFromCells();
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedRemovedWhiteSpaceFromCellsComment;
      throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.removeFormattingFromSheetCells();
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedRemovedFormattingFromCellsComment;
      throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    userImportTemplate.clearSheetSummaryColumn(userImportTemplate.sheet);
    userImportTemplate.clearSheetSummaryColumn(reportSheet);
    userImportTemplate.setErrorColumnHeaderInMainSheet();
    
    try {
      userImportTemplate.checkForDuplicateEmails(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForDuplicateEmails;
      throw new Error(`Emails not checked for duplicates. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    try {
      userImportTemplate.checkForDuplicateEmailsAndSetErrors(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedErrorColumnsNotSetForDuplicateEmails;
      throw new Error(`Error columns not set for duplicate emails. Reason: ${err.name}: ${err.message}. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try{
      userImportTemplate.checkFirstThreeColumnsForBlanks(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckFirstThreeColumnsForMissingValues;
      throw new Error(`Check not ran for missing values within first three columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first three colums (i.e. "First Name, Last Name, and Email is running a User Import Check), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.formatAllDatedColumns(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedFormatDateColumns; //User Only
      throw new Error(`Check not ran for formatting of dated columns. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    try {
        userImportTemplate.formatCommaSeparatedLists(reportSheet);
      } catch (err) {
          Logger.log(err);
          userImportTemplate.reportSummaryComments = userImportTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
          throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
    
    try {
      userImportTemplate.convertStatesToTwoLetterCode(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForConvertingStatesToTwoLetterCodes; //User Programs Needs Agencies
      throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.validateStates(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidStatesCheckMessage; //Users Needs Agencies
      throw new Error(`Check not ran for invalid states: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.validateGenderOptions(reportSheet);
    } catch (err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForInvalidGenderOptions; //Users Only
      throw new Error(`Check not ran for invalid gender options: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      userImportTemplate.checkForInvalidEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidEmailCheckMessage; //All
        throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
    
    try {
      userImportTemplate.checkForInvalidPostalCodes(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedCheckNotRanForInvalidPostalCodes; //Usrs Neesd Agencies
      throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      userImportTemplate.checkForInvalidNumbers(reportSheet);
    } catch(err) {
      Logger.log(err);
      userImportTemplate.reportSummaryComments = userImportTemplate.failedInvalidPhoneNumbersCheck; //Users Needs Agencies
      throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      userImportTemplate.setCommentsOnReportCell(reportSheet);
    } catch(err) {
      Logger.log(err);
      throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`); //All
    }
      //Has a pop-up tell the user the check has been completed
     SpreadsheetApp.getUi().alert("User Import Check Complete");
    
  }
  
  } catch(err) {
  Logger.log(err);
  throw new Error(`An error occured the the user import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  
  