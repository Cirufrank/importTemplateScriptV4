/*

Goals met:

removed white space
Check for invalid additional contact emails YASSSSS
Checks for invalid emails 
checks for missing required values (first four columns-this is position based and imports)


Things to keep in mind:

Need columns for singular email check to be titles "Email"

This may be okay due to position being what matters in the future */


function checkFirstFourColumnsForBlanks(sheetBinding, reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
    let rowStartPosition = 2;
    let columnStartPosition = 1;
    let secondColumnPosition = 2;
    let thirdColumnPositon = 3;
    let fourthColumnPosition = 4;
    let maxRows = sheetBinding.getMaxRows();
    let totalColumsToCheck = 4;
    let range = sheetBinding.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
    let values = range.getValues();
  
    for (let row = 0; row < values.length; row += 1) {
      let headerRowPosition = 1;
      let cellRow = row + 2;
      let currentRow = values[row];
      let item1 = currentRow[0];
      let item1CurrentCell = getSheetCell(sheetBinding, cellRow, columnStartPosition);
      let item1ReportCurrentCell = getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
      let item1HeaderCell = getSheetCell(sheetBinding, headerRowPosition, columnStartPosition);
      let item2 = currentRow[1];
      let item2CurrentCell = getSheetCell(sheetBinding, cellRow, secondColumnPosition);
      let item2ReportCurrentCell = getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
      let item2HeaderCell = getSheetCell(sheetBinding, headerRowPosition, secondColumnPosition);
      let item3 = currentRow[2];
      let item3CurrentCell = getSheetCell(sheetBinding, cellRow, thirdColumnPositon);
      let item3ReportCell = getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      let item3HeaderCell = getSheetCell(sheetBinding, headerRowPosition, thirdColumnPositon);
      let item4 = currentRow[3];
      let item4CurrentCell = getSheetCell(sheetBinding, cellRow, fourthColumnPosition);
      let item4ReportCell = getSheetCell(reportSheetBinding, cellRow, fourthColumnPosition);
      let item4HeaderCell = getSheetCell(sheetBinding, headerRowPosition, fourthColumnPosition);
  
      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0) {
        currentRow.forEach((val, index) => {
          if (val === "") {
            switch (index) {
              case 0:
                setSheetCellBackground(item1CurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item1ReportCurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item1HeaderCell, LIGHT_RED_HEX_CODE);
                insertCommentToSheetCell(item1ReportCurrentCell, VALUES_MISSING_COMMENT);
                setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
                insertHeaderComment(item1HeaderCell, VALUES_MISSING_COMMENT);
                break;
              case 1:
                setSheetCellBackground(item2CurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item2ReportCurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item2HeaderCell, LIGHT_RED_HEX_CODE);
                insertCommentToSheetCell(item2ReportCurrentCell, VALUES_MISSING_COMMENT);
                setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
                insertHeaderComment(item2HeaderCell, VALUES_MISSING_COMMENT);
                break;
              case 2:
                setSheetCellBackground(item3CurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item3ReportCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item3HeaderCell, LIGHT_RED_HEX_CODE);
                insertCommentToSheetCell(item3ReportCell, VALUES_MISSING_COMMENT);
                setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
                insertHeaderComment(item3HeaderCell, VALUES_MISSING_COMMENT);
                break;
              case 3:
                setSheetCellBackground(item4CurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item4ReportCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item4HeaderCell, LIGHT_RED_HEX_CODE);
                insertCommentToSheetCell(item4ReportCell, VALUES_MISSING_COMMENT);
                setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
                insertHeaderComment(item4HeaderCell, VALUES_MISSING_COMMENT);
                break;
            }
          }
        });
      }
    }
  
  reportSummaryCommentsBinding.push("Success: checked for missing values in first four columns");
  
  }
  
  function formatAdditionalContactEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) {
    let headerRow = 1;
    let emailColumnRange = getColumnRange('Additional Contacts (emails)', sheetBinding, columnsHeadersBinding);
  
    if (emailColumnRange) {
    let emailColumnValues = getValues(emailColumnRange);
    let emailColumnPosition = emailColumnRange.getColumn();
  
    emailColumnValues.forEach((emailString, index) => {
      let row = index + 2;
      let emailArray = String(emailString).split(",");
  
  
        if (emailArray.length > 1) {
          let emailArrayNew = emailArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(", ");
          let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
          let reportSheetHeaderCell = getSheetCell(reportSheetBinding, headerRow, emailColumnPosition);
          let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
          let removedWhiteSpaceAndBlanksValue = emailArrayNew;
          currentCell.setValue(removedWhiteSpaceAndBlanksValue);
  
            insertHeaderComment(reportSheetHeaderCell, "Removed whitespace and empty values from Additional Contact Emails on main sheet");
            setSheetCellBackground(reportSheetHeaderCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
        }
  
        });
  
        reportSummaryCommentsBinding.push("Success: removed whitespace and empty values from Additional Contact Emails");
  
        } else {
      reportSummaryCommentsBinding.push("Success: removed whitespace and empty values from Additional Contact Emails, but did note find column");
    }
    
  }
  
  //TODOOOOO
  
  function checkForInvalidAdditionalContactEmails(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
    let headerRow = 1;
    let emailColumnRange = getColumnRange('Additional Contacts (emails)', sheetBinding, columnsHeadersBinding);
  
    if (emailColumnRange) {
      let emailColumnValues = getValues(emailColumnRange);
      let emailColumnPosition = emailColumnRange.getColumn();
  
    emailColumnValues.forEach((emailArray, index) => {
      let row = index + 2;
  
        if (typeof emailArray === 'string') {
          let currentEmail = emailArray;
          if (currentEmail !== "" && !validateEmail(currentEmail)) {
            let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
            let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
            let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  
            //Line below for testing
  
            // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
            //testing line ends here
            setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
            insertCommentToSheetCell(reportSheetCell, INVALID_EMAIL_COMMENT);
            insertHeaderComment(mainSheetHeaderCell, "Invalid Email/Emails found");
            setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
          }
  
        } else {
            emailArray.forEach(email => {
              let currentEmail = email.trim();
            if (currentEmail !== "" && !validateEmail(currentEmail)) {
            let currentCell = getSheetCell(sheetBinding, row, emailColumnPosition);
            let reportSheetCell = getSheetCell(reportSheetBinding, row, emailColumnPosition);
            let mainSheetHeaderCell = getSheetCell(sheetBinding, headerRow, emailColumnPosition);
  
            //Line below for testing
  
            // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
            //testing line ends here
            setSheetCellBackground(reportSheetCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(currentCell, LIGHT_RED_HEX_CODE);
            setSheetCellBackground(mainSheetHeaderCell, LIGHT_RED_HEX_CODE);
            insertCommentToSheetCell(reportSheetCell, INVALID_EMAIL_COMMENT);
            insertHeaderComment(mainSheetHeaderCell, "Invalid Email/Emails found");
            setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, row);
            }
  
          });
  
        }
  
        
        }
      );
  
      reportSummaryCommentsBinding.push("Success: ran check for invalid emails");
      
    } else {
      reportSummaryCommentsBinding.push("Success: ran check for invalid emails, but did not find column");
    }
    
  }
  
  try {
    function programsAndAgenciesTemplateCheck() {
      const REPORT_SHEET_NAME = 'Agencies/Programs Report';
      const AGENCIES_AND_PROGRAMS_TEMPLATE_NAME = 'Agencies/Programs Import Template';
      const sheet = SpreadsheetApp.getActiveSheet();
      const data = sheet.getDataRange();
      const values = data.getValues();
      const columnHeaders = values[0].map(header => header.trim());
      //Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here
      const reportSummaryColumnPosition = columnHeaders.length + 2;
      const reportSummaryComments = [];
      let reportSheet = createReportSheet(ss, values, REPORT_SHEET_NAME);
  
      sheet.setName(AGENCIES_AND_PROGRAMS_TEMPLATE_NAME);
      
  
      ss.setActiveSheet(getSheet(AGENCIES_AND_PROGRAMS_TEMPLATE_NAME));
  
      
  
  
      setFrozenRows(sheet, 1);
      setFrozenRows(reportSheet, 1);
  
     try {
            removeWhiteSpaceFromCells(sheet, values, reportSummaryComments);
          } catch (err) {
            Logger.log(err);
            reportSummaryComments.push("Failed: remove white space from cells");
            throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
      
      clearSheetSummaryColumn(sheet, reportSummaryColumnPosition);
      clearSheetSummaryColumn(reportSheet, reportSummaryColumnPosition);
      setErrorColumnHeaderInMainSheet(sheet, reportSummaryColumnPosition);
  
      try{
        checkFirstFourColumnsForBlanks(sheet, reportSheet, reportSummaryComments, reportSummaryColumnPosition);
          } catch(err) {
              Logger.log(err);
              reportSummaryComments.push("Failed: check first four columns for missing values");
              throw new Error(`Check not ran for missing values within first four columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
  
      try {
      convertStatesToTwoLetterCode(sheet, reportSheet, columnHeaders, reportSummaryComments);
    } catch (err) {
      Logger.log(err);
      reportSummaryComments.push("Failed: check not ran for converting states to two-letter code");
      throw new Error(`Check not ran for converting states to two-letter code: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
    try {
      checkForInvalidEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
    } catch(err) {
        Logger.log(err);
        reportSummaryComments.push("Failed: check not ran for invalid emails");
        throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
  
      try {
        formatAdditionalContactEmails(sheet, reportSheet, columnHeaders, reportSummaryComments);
      } catch (err) {
          Logger.log(err);
          reportSummaryComments.push("Failed: did not remove whitespace and empty values from Additional Contact Emails");
          throw new Error(`Did not remove whitespace and empty values from Additional Contact Emails. Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
  
      try {
      checkForInvalidAdditionalContactEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
      } catch(err) {
          Logger.log(err);
          reportSummaryComments.push("Failed: check not ran for invalid additional contact emails");
          throw new Error(`Check not ran for invalid additional contact emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        } 
    
    try {
      checkForInvalidPostalCodes(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
    } catch(err) {
      Logger.log(err);
      reportSummaryComments.push("Failed: check not ran for invalid postal codes");
      throw new Error(`Check not ran for invalid postal codes. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    try {
      checkForInvalidNumbers(sheet, reportSheet,columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
    } catch(err) {
      Logger.log(err);
      reportSummaryComments.push("Failed: check not ran for invalid phone numbers");
      throw new Error(`Check not ran for invalid home or mobile phone numbers. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
    
    try {
      setCommentsOnReportCell(reportSheet, reportSummaryComments, reportSummaryColumnPosition);
    } catch(err) {
      Logger.log(err);
      throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
     SpreadsheetApp.getUi().alert("Agencies/Programs Import Check Complete");
    }
  
  } catch (err) {
      Logger.log(err);
      throw new Error(`An error occured the the Agencies/Programs import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
  