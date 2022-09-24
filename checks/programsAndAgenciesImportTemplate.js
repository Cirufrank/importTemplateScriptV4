////////////////////////////////////////////////////////////
//Runs the checks for the Programs/Agencies Import Template
////////////////////////////////////////////////////////////

try {
  //Executes when the Check Agencies/Programs button is clicked
  function programsAndAgenciesTemplateCheck() {
    const programsAndAgenciesTemplate = new ProgramsAndAgenciesTemplate();
    const reportSheet = programsAndAgenciesTemplate.createReportSheet();

    programsAndAgenciesTemplate.sheet.setName(programsAndAgenciesTemplate.templateName);
    

   programsAndAgenciesTemplate.ss.setActiveSheet(programsAndAgenciesTemplate.getSheet(programsAndAgenciesTemplate.templateName));

    


    programsAndAgenciesTemplate.setFrozenRows(programsAndAgenciesTemplate.sheet, 1);
    programsAndAgenciesTemplate.setFrozenRows(reportSheet, 1);

   try {
          programsAndAgenciesTemplate.removeWhiteSpaceFromCells();
        } catch (err) {
          Logger.log(err);
          programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemovedWhiteSpaceFromCellsComment;
          throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

        try {
          programsAndAgenciesTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
    
    programsAndAgenciesTemplate.clearSheetSummaryColumn(programsAndAgenciesTemplate.sheet);
    programsAndAgenciesTemplate.clearSheetSummaryColumn(reportSheet);
    programsAndAgenciesTemplate.setErrorColumnHeaderInMainSheet(programsAndAgenciesTemplate.sheet);

    try{
      programsAndAgenciesTemplate.checkFirstFourColumnsForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedFirstFourColumnsMissingValuesCheckMessage;
            throw new Error(`Check not ran for missing values within first four columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

    try {
    programsAndAgenciesTemplate.convertStatesToTwoLetterCode(reportSheet);
  } catch (err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForConvertingStatesToTwoLetterCodes ;
    throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    programsAndAgenciesTemplate.validateStates(reportSheet);
  } catch (err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidStatesCheckMessage;
    throw new Error(`Check not ran for invalid states: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  try {
    programsAndAgenciesTemplate.checkForInvalidEmails(reportSheet);
  } catch(err) {
      Logger.log(err);
      programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidEmailCheckMessage;
      throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    } 

    try {
      programsAndAgenciesTemplate.formatCommaSeparatedLists(reportSheet);
    } catch (err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
        throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    programsAndAgenciesTemplate.checkForInvalidCommaSeparatedEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForInvalidCommaSeparatedEmails;
        throw new Error(`Check not ran for invalid additional contact emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
  
  try {
    programsAndAgenciesTemplate.checkForInvalidPostalCodes(reportSheet);
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidPostalCodeCheck;
    throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidNumbers(reportSheet);
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedInvalidPhoneNumbersCheck;
    throw new Error(`Check not ran for invalid phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'general');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForGeneralURL;
    throw new Error(`Check not ran for invalid general URL. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'youtube');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForYouTubeLink;
    throw new Error(`Check not ran for invalid YouTube link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'twitter');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid Twitter link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'linkedin');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid LinkedIn link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'instagram');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForTwitterLink;
    throw new Error(`Check not ran for invalid Instagram link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.checkForInvalidURL(reportSheet, 'facebook');
  } catch(err) {
    Logger.log(err);
    programsAndAgenciesTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedCheckNotRanForFacebookLink;
    throw new Error(`Check not ran for invalid Facebook link. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  try {
    programsAndAgenciesTemplate.setCommentsOnReportCell(reportSheet);
  } catch(err) {
    Logger.log(err);
    throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
    //Has a pop-up tell the user the check has been completed
   SpreadsheetApp.getUi().alert("Agencies/Programs Import Check Complete");
  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the Agencies/Programs import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

