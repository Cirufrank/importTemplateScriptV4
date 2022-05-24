/*

Goals Accomplished:

Formats "Date Served column (relies on "Date Served" spelling
checks for invalid email addresses
Check for missing values in first threee (required columns)
Inserts the individual hours records for us

Things to keep in mind:
Wanted the word "Individual" to be inserted during the time the required columsn are being checked so that the code doesn't have to be re-creatd with a idential function (also speeds up performance)

The error column is an extra column down so when no "Hours Type" column is provided it does not clash with the errors colum (which is typcially inserted after the last row of values from the start (which would be before the indv values are created)

Think about if it's be worth it do abstract the functionality out of the other function

Things to keep in mind:

Reliant on 'Date Served being spelled correctly for format to run'

Anything in column 5 will be overriden

re-wrote some function due to slightly custom needs with indv hours template checker

*/

class IndividualHoursTemplate extends Template {
  constructor() {
    super();
    this._foundFirstThreeColumnsAndRanCheckForMissingValuesComment = "Success: checked for missing values in first three columns and added the 'Hours Type' column with a record of 'Individual' anywhere that row with records were found";
    this._failedCheckFirstThreeColumnsForBlanksAndAddIndvColumnMessage = "Failed: check first three columns for missing values and add 'Individual' hours type column";
    this._theWordIndividual = 'Individual';
    this._individualHoursTypeColumnPosition = 5;
  }
  get failedCheckFirstThreeColumnsForBlanksAndAddIndvColumnMessage() {
    return this._failedCheckFirstThreeColumnsForBlanksAndAddIndvColumnMessage;
  }

    setIndividualHoursHeaderCellVal() {
      let headerRow = 1;
      this.getSheetCell(this._sheet, headerRow, this._individualHoursTypeColumnPosition).setValue('Hours Type');
    }
    checkFirstThreeColumnsForBlanksAndAddIndvColumn(reportSheetBinding) {
      let rowStartPosition = 2;
      let columnStartPosition = 1;
      let middleColumnPosition = 2;
      let thirdColumnPositon = 3;
      let maxRows = this._sheet.getMaxRows();
      let totalColumsToCheck = 3;
      let range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
      let values = range.getValues();
    
      for (let row = 0; row < values.length; row += 1) {
        let headerRowPosition = 1;
        let cellRow = row + 2;
        let currentRow = values[row];
        let item1 = currentRow[0];
        let item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
        let item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
        let item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
        let item1IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
        let item2 = currentRow[1];
        let item2CurrentCell = this.getSheetCell(this._sheet, cellRow, middleColumnPosition);
        let item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
        let item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, middleColumnPosition);
        let item2IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
        let item3 = currentRow[2];
        let item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
        let item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
        let item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
        let item3IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
    
        if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
          item1IndvHoursCell.setValue(this._theWordIndividual);
          item2IndvHoursCell.setValue(this._theWordIndividual);
          item3IndvHoursCell.setValue(this._theWordIndividual);
          this.setIndividualHoursHeaderCellVal();
    
          currentRow.forEach((val, index) => {
            if (val === "") {
              switch (index) {
                case 0:
                  this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding,cellRow);
                  this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
          
                  break;
                case 1:
                  this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                  
                  break;
                case 2:
                  this.setSheetCellBackground(item3CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item3ReportCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item3HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item3ReportCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item3HeaderCell, this._valuesMissingComment);
                  
                  break;
              }
            }
          });
        }
      }
    
      this._reportSummaryComments.push(this._foundFirstThreeColumnsAndRanCheckForMissingValuesComment);
    
    }
  
  }

  
  try {
  
    function checkIndvImportTemplate() {
        const individualHoursTemplate = new IndividualHoursTemplate();
        let reportSheet = individualHoursTemplate.createReportSheet();
  
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

      SpreadsheetApp.getUi().alert("Individual Hours Import Check Complete");
  
  
    }
  } catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the individual hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }
  