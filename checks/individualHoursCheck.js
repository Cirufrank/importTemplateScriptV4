////////////////////////////////////////////////////////////
//Runs the checks for the Individual Hours Import Template
////////////////////////////////////////////////////////////

try {
    //Executes when the Check Individual Hour Template button is pressed
    function checkIndvImportTemplate() {
        const individualHoursTemplate = new IndividualHoursTemplate();
        const reportSheet = individualHoursTemplate.createReportSheet();
  
        individualHoursTemplate.sheet.setName(individualHoursTemplate.templateName);
        
  
        individualHoursTemplate.ss.setActiveSheet(individualHoursTemplate.getSheet(individualHoursTemplate.templateName));
      
  
        individualHoursTemplate.setFrozenRows(individualHoursTemplate.sheet, 1);
        individualHoursTemplate.setFrozenRows(reportSheet, 1);
  
        
        try {
          individualHoursTemplate.removeWhiteSpaceFromCells();
        } catch (err) {
          Logger.log(err);
          individualHoursTemplate.reportSummaryComments = individualHoursTemplate.failedRemovedWhiteSpaceFromCellsComment;
          throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

        try {
          individualHoursTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          individualHoursTemplate.reportSummaryComments = individualHoursTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
  
        individualHoursTemplate.clearSheetSummaryColumn(individualHoursTemplate.sheet);
        individualHoursTemplate.clearSheetSummaryColumn(reportSheet);
        individualHoursTemplate.setErrorColumnHeaderInMainSheet();
  
        try {
          individualHoursTemplate.checkForHourTypeColumn();
          individualHoursTemplate.checkFirstThreeColumnsForBlanksAndAddIndvColumn(reportSheet);
        } catch(err) {
          Logger.log(err);
          individualHoursTemplate.reportSummaryComments = individualHoursTemplate.failedCheckFirstThreeColumnsForBlanksAndAddIndvColumnMessage;
          throw new Error(`Check not ran for missing values within first three columns and 'Hour Type' column not filled with the value 'Individual'. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first three colums (i.e. "First Name, Last Name, and Email is running a User Import Check), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
  
        try {
        individualHoursTemplate.formatAllDatedColumns(reportSheet);
      } catch(err) {
        Logger.log(err);
        individualHoursTemplate.reportSummaryComments = individualHoursTemplate._failedFormatDateColumns;
        throw new Error(`Check not ran for formatting of the date served column. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
  
       try {
        individualHoursTemplate.checkForInvalidEmails(reportSheet);
      } catch(err) {
          Logger.log(err);
          individualHoursTemplate.reportSummaryComments = individualHoursTemplate.failedInvalidEmailCheckMessage;
          throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        } 
  
       try {
        individualHoursTemplate.setCommentsOnReportCell(reportSheet);
      } catch(err) {
        Logger.log(err);
        throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
      //Has a pop-up tell the user the check has been completed
      SpreadsheetApp.getUi().alert("Individual Hours Import Check Complete");
  
  
    }
  } catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the individual hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  