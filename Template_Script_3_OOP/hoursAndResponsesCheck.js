class HoursAndResponsesTemplate extends IndividualHoursTemplate {
    constructor() {
      super();
      this._foundFirstFiveColumnsAndCheckedForMissingValuesComment = "Success: checked for missing values in first five columns";
      this._failedFirstFiveColumnsMissingValuesCheck = "Failed: check first five columns for missing values";
    }
    get failedFirstFiveColumnsMissingValuesCheck() {
      return this._failedFirstFiveColumnsMissingValuesCheck;
    }
    checkFirstFiveColumnsForBlanks(reportSheetBinding) {
      let rowStartPosition = 2;
      let columnStartPosition = 1;
      let secondColumnPosition = 2;
      let thirdColumnPositon = 3;
      let fourthColumnPosition = 4;
      let fifthColumnPosition = 5;
      let maxRows = this._sheet.getMaxRows();
      let totalColumsToCheck = 5;
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
        let item2 = currentRow[1];
        let item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
        let item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
        let item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);
        let item3 = currentRow[2];
        let item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
        let item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
        let item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
        let item4 = currentRow[3];
        let item4CurrentCell = this.getSheetCell(this._sheet, cellRow, fourthColumnPosition);
        let item4ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fourthColumnPosition);
        let item4HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fourthColumnPosition);
        let item5 = currentRow[4];
        let item5CurrentCell = this.getSheetCell(this._sheet, cellRow, fifthColumnPosition);
        let item5ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fifthColumnPosition);
        let item5HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fifthColumnPosition);
  
        if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0 || item5.length !== 0) {
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
                case 3:
                  this.setSheetCellBackground(item4CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item4ReportCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item4HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item4ReportCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item4HeaderCell, this._valuesMissingComment);
                  break;
                case 4:
                  this.setSheetCellBackground(item5CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item5ReportCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item5HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item5ReportCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item5HeaderCell, this._valuesMissingComment);
                  break;
              }
            }
          });
        }
      }
  
    this._reportSummaryComments.push(this._foundFirstFiveColumnsAndCheckedForMissingValuesComment);
  
    }
  
  }
  
  
  try {
    function hoursAndResponsesCheck() {
      let responsesAndHoursTemplate = new HoursAndResponsesTemplate();
      let reportSheet = responsesAndHoursTemplate.createReportSheet();
  
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
            responsesAndHoursTemplate.formatDateServedColumn(reportSheet);
          } catch(err) {
            Logger.log(err);
            responsesAndHoursTemplate.reportSummaryComments = responsesAndHoursTemplate.failedDidNotFormatDateServedColumnMessage;
            throw new Error(`Check not ran for formatting of the date served column. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
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
  
          SpreadsheetApp.getUi().alert("Responses and Hours Import Check Complete");
  
    }
  
  } catch (err) {
      Logger.log(err);
      throw new Error(`An error occured the the responses and hours import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }
  
  