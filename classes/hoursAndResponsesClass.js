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
    const emailIndex = 0;
    const opportunityIndex = 1;
    const programIndex = 2;
    const dateIndex = 3;
    const hoursIndex = 4;
    const rowStartPosition = 2;
    const columnStartPosition = 1;
    const secondColumnPosition = 2;
    const thirdColumnPositon = 3;
    const fourthColumnPosition = 4;
    const fifthColumnPosition = 5;
    const maxRows = this._sheet.getMaxRows();
    const totalColumsToCheck = 5;
    const range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
    const values = range.getValues();

    for (let row = 0; row < values.length; row += 1) {
      const zeroRangeAndFirstRow = 2;
      const headerRowPosition = 1;
      const cellRow = row + zeroRangeAndFirstRow;
      const currentRow = values[row];
      //item 1 represents an email address of the user that needs the hours added
      const item1 = currentRow[emailIndex];
      const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
      const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
      const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
      //item 2 represents a opportunity title
      const item2 = currentRow[opportunityIndex];
      const item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
      const item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
      const item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);
      //item 3 represents a program name
      const item3 = currentRow[programIndex];
      const item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
      const item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      const item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
      //item 4 represents a hours date served
      const item4 = currentRow[dateIndex];
      const item4CurrentCell = this.getSheetCell(this._sheet, cellRow, fourthColumnPosition);
      const item4ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fourthColumnPosition);
      const item4HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fourthColumnPosition);
      //item 5 represents a hours entry
      const item5 = currentRow[hoursIndex];
      const item5CurrentCell = this.getSheetCell(this._sheet, cellRow, fifthColumnPosition);
      const item5ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fifthColumnPosition);
      const item5HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fifthColumnPosition);
      //if any of the required fields are filled with a value then all of them need to be
      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0 || item5.length !== 0) {
        //iterates through the current row of required values
        currentRow.forEach((val, index) => {
          //if the current value is empty then takes the nexessary actions to flag it
          if (val === "") {
            switch (index) {
              //missing email
              case 0:
                this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding,cellRow);
                this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
                break;
              //missing opportunity title
              case 1:
                this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                break;
              //missing program name
              case 2:
                this.setSheetCellBackground(item3CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item3ReportCell, this._lightRedHexCode);
                this.setSheetCellBackground(item3HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item3ReportCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item3HeaderCell, this._valuesMissingComment);
                break;
              //missing date served entry
              case 3:
                this.setSheetCellBackground(item4CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4ReportCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item4ReportCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item4HeaderCell, this._valuesMissingComment);
                break;
              //missing hours entry
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