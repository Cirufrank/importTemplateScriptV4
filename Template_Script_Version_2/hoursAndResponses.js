function checkFirstFiveColumnsForBlanks(sheetBinding, reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding) {
    let rowStartPosition = 2;
    let columnStartPosition = 1;
    let secondColumnPosition = 2;
    let thirdColumnPositon = 3;
    let fourthColumnPosition = 4;
    let fifthColumnPosition = 5;
    let maxRows = sheetBinding.getMaxRows();
    let totalColumsToCheck = 5;
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
      let item5 = currentRow[4];
      let item5CurrentCell = getSheetCell(sheetBinding, cellRow, fifthColumnPosition);
      let item5ReportCell = getSheetCell(reportSheetBinding, cellRow, fifthColumnPosition);
      let item5HeaderCell = getSheetCell(sheetBinding, headerRowPosition, fifthColumnPosition);
  
      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0 || item5.length !== 0) {
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
              case 4:
                setSheetCellBackground(item5CurrentCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item5ReportCell, LIGHT_RED_HEX_CODE);
                setSheetCellBackground(item5HeaderCell, LIGHT_RED_HEX_CODE);
                insertCommentToSheetCell(item5ReportCell, VALUES_MISSING_COMMENT);
                setErrorColumns(sheetBinding, reportSheetBinding, reportSummaryColumnPositionBinding, cellRow);
                insertHeaderComment(item5HeaderCell, VALUES_MISSING_COMMENT);
                break;
            }
          }
        });
      }
    }
  
  reportSummaryCommentsBinding.push("Success: checked for missing values in first five columns");
  
  }
  
  try {
    function hoursAndResponsesCheck() {
      const REPORT_SHEET_NAME = 'Responses and Hours Report';
      const HOURS_AND_RESPONSES_TEMPLATE_NAME = 'Responses and Hours Import Template';
      const sheet = SpreadsheetApp.getActiveSheet();
      const data = sheet.getDataRange();
      const values = data.getValues();
      const columnHeaders = values[0].map(header => header.trim());
      //Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here
      const reportSummaryColumnPosition = columnHeaders.length + 2;
      const reportSummaryComments = [];
      let reportSheet = createReportSheet(ss, values, REPORT_SHEET_NAME);
  
      sheet.setName(HOURS_AND_RESPONSES_TEMPLATE_NAME);
      
  
      ss.setActiveSheet(getSheet(HOURS_AND_RESPONSES_TEMPLATE_NAME));
  
  
      setFrozenRows(sheet, 1);
      setFrozenRows(reportSheet, 1);
  
      try {
              removeWhiteSpaceFromCells(sheet, values, reportSummaryComments);
            } catch (err) {
              Logger.log(err);
              reportSummaryComments.push("Failed: remove white space from cells");
              throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
            }
  
            try {
            removeFormattingFromSheetCells(sheet, values, reportSummaryComments);
            } catch (err) {
              Logger.log(err);
              reportSummaryComments.push("Failed: remove formatting from cells");
              throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
            }
      
            clearSheetSummaryColumn(sheet, reportSummaryColumnPosition);
            clearSheetSummaryColumn(reportSheet, reportSummaryColumnPosition);
            setErrorColumnHeaderInMainSheet(sheet, reportSummaryColumnPosition);
  
            try {
            formatDateServedColumn(sheet, reportSheet, columnHeaders, reportSummaryComments);
          } catch(err) {
            Logger.log(err);
            reportSummaryComments.push("Failed: did not format date served column");
            throw new Error(`Check not ran for formatting of the date served column. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
      
          try {
            checkForInvalidEmails(sheet, reportSheet, columnHeaders, reportSummaryComments, reportSummaryColumnPosition);
          } catch(err) {
              Logger.log(err);
              reportSummaryComments.push("Failed: check not ran for invalid emails");
              throw new Error(`Check not ran for invalid emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the email column is titled "Email" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
            } 
  
            try{
        checkFirstFiveColumnsForBlanks(sheet, reportSheet, reportSummaryComments, reportSummaryColumnPosition);
          } catch(err) {
              Logger.log(err);
              reportSummaryComments.push("Failed: check first five columns for missing values");
              throw new Error(`Check not ran for missing values within first five columns. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first five colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for five]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
  
            try {
            setCommentsOnReportCell(reportSheet, reportSummaryComments, reportSummaryColumnPosition);
          } catch(err) {
            Logger.log(err);
            throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
          }
  
          SpreadsheetApp.getUi().alert("Responses and Hours Import Check Complete");
  
    }
  
  } catch (err) {
      Logger.log(err);
      throw new Error(`An error occured the the responses and hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  