/*

Goals Accomplished:

Formats "Date Served column (relies on "Date Served" spelling
checks for invalid email addresses
Check for missing values in first threee (required columns)
Inserts the individual hours records for us

Things to keep in mind:
Wanted the word "Individual" to be inserted during the time the required columsn are being checked so that the code doesn't have to be re-creatd with a idential function (also speeds up performance)

The error column is an extra column down so when no "Hours Type" column is provided it does not clash with the errors colum (which is typcially inserted after the last row of values from the start (which would be before the indv values are created)

Think about if it's be worth it do abstract the functionality out of the other function*/





function formatDateServedColumn(sheetBinding, reportSheetBinding, columnsHeadersBinding, reportSummaryCommentsBinding) {
    let row = 1;
    let dateServedColumn = getColumnRange('Date Served', sheetBinding, columnsHeadersBinding);
  
    if (dateServedColumn) {
      let columnPosition = dateServedColumn.getColumn();
      let reportSheetHeaderCell = getSheetCell(reportSheetBinding, row, columnPosition);
  
      setColumnToYYYYMMDDFormat(dateServedColumn);
      setSheetCellBackground(reportSheetHeaderCell, LIGHT_GREEN_HEX_CODE);
      insertCommentToSheetCell(reportSheetHeaderCell, DATE_FORMATTED_YYYY_MM_DD_COMMENT);
      reportSummaryCommentsBinding.push("Success: formatted date served column");
    }
  
  }
  
  function setIndividualHoursHeaderCellVal(mainSheetBinding, indvColumnPosition) {
    let headerRow = 1;
    getSheetCell(mainSheetBinding, headerRow, indvColumnPosition).setValue('Hours Type');
  }
  
  function checkFirstThreeColumnsForBlanksAndAddIndvColumn(sheetBinding, reportSheetBinding, reportSummaryCommentsBinding, reportSummaryColumnPositionBinding, indvColumnPosition) {
    let rowStartPosition = 2;
    let columnStartPosition = 1;
    let middleColumnPosition = 2;
    let thirdColumnPositon = 3;
    let maxRows = sheetBinding.getMaxRows();
    let totalColumsToCheck = 3;
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
      let item1IndvHoursCell = getSheetCell(sheetBinding, cellRow, indvColumnPosition);
      let item2 = currentRow[1];
      let item2CurrentCell = getSheetCell(sheetBinding, cellRow, middleColumnPosition);
      let item2ReportCurrentCell = getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
      let item2HeaderCell = getSheetCell(sheetBinding, headerRowPosition, middleColumnPosition);
      let item2IndvHoursCell = getSheetCell(sheetBinding, cellRow, indvColumnPosition);
      let item3 = currentRow[2];
      let item3CurrentCell = getSheetCell(sheetBinding, cellRow, thirdColumnPositon);
      let item3ReportCell = getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      let item3HeaderCell = getSheetCell(reportSheetBinding, headerRowPosition, thirdColumnPositon);
      let item3IndvHoursCell = getSheetCell(sheetBinding, cellRow, indvColumnPosition);
  
      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
        item1IndvHoursCell.setValue('Individual');
        item2IndvHoursCell.setValue('Individual');
        item3IndvHoursCell.setValue('Individual');
        setIndividualHoursHeaderCellVal(sheetBinding, indvColumnPosition);
  
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
                insertHeaderComment(item3HeaderCell, VALUES_MISSING_COMMENT);
                
                break;
            }
          }
        });
      }
    }
  
    reportSummaryCommentsBinding.push("Success: checked for missing values in first three columns and added the 'Hours Type' column with a record of 'Individual' anywhere that row with records were found");
  
  }
  
  
  
  try {
  
    function checkIndvImportTemplate() {
        const REPORT_SHEET_NAME = 'Indv Hours Report';
        const INDV_HOURS_TEMPLATE_NAME = 'Individual hours Import Template';
        const INDV_HOURS_TYPE_COLUMN_POSITION = 5;
        const sheet = SpreadsheetApp.getActiveSheet();
        const data = sheet.getDataRange();
        const values = data.getValues();
        const columnHeaders = values[0].map(header => header.trim());
        //Consider values[0].map(header => (header[0] + header.substr(1)).trim()) to help avoid errors here
        const reportSummaryColumnPosition = columnHeaders.length + 2;
        const reportSummaryComments = [];
        let reportSheet = createReportSheet(ss, values, REPORT_SHEET_NAME);
  
        sheet.setName(INDV_HOURS_TEMPLATE_NAME);
        
  
        ss.setActiveSheet(getSheet(INDV_HOURS_TEMPLATE_NAME));
      
  
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
  
        try {
          checkFirstThreeColumnsForBlanksAndAddIndvColumn(sheet, reportSheet, reportSummaryComments, reportSummaryColumnPosition, INDV_HOURS_TYPE_COLUMN_POSITION);
        } catch(err) {
          Logger.log(err);
          reportSummaryComments.push("Failed: check first three columns for missing values");
          throw new Error(`Check not ran for missing values within first three columns and 'Hour Type' column not filled with the value 'Individual'. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first three colums (i.e. "First Name, Last Name, and Email is running a User Import Check), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
  
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
  
       try {
        setCommentsOnReportCell(reportSheet, reportSummaryComments, reportSummaryColumnPosition);
      } catch(err) {
        Logger.log(err);
        throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
      }
  
  
    }
  } catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the individual hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  
  
  