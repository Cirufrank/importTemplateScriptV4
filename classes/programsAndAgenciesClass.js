class ProgramsAndAgenciesTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._foundFirstFourColumnsAndCheckedThemForMissingValuesComment = "Success: checked for missing values in first four columns";
    this._failedFirstFourColumnsMissingValuesCheckMessage = "Failed: check first four columns for missing values";
  }
  get failedFirstFourColumnsMissingValuesCheckMessage() {
    return this._failedFirstFourColumnsMissingValuesCheckMessage;
  }
  checkFirstFourColumnsForBlanks(reportSheetBinding) {
    const programIndex = 0;
    const managerFNIndex = 1;
    const managerLNIndex = 2;
    const managerEmailIndex = 3;
    const rowStartPosition = 2;
    const columnStartPosition = 1;
    const secondColumnPosition = 2;
    const thirdColumnPositon = 3;
    const fourthColumnPosition = 4;
    const maxRows = this._sheet.getMaxRows();
    const totalColumsToCheck = 4;
    const range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
    const values = range.getValues();

    for (let row = 0; row < values.length; row += 1) {
      const zeroRangeAndFirstRow = 2;
      const headerRowPosition = 1;
      const cellRow = row + zeroRangeAndFirstRow;
      const currentRow = values[row];
      //item 1 represents a program name
      const item1 = currentRow[programIndex];
      const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
      const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
      const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
      //item 2 represents a program manager's first name
      const item2 = currentRow[managerFNIndex];
      const item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
      const item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
      const item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);
      //item 3 represents a program manager's last name
      const item3 = currentRow[managerLNIndex];
      const item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
      const item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      const item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
      //item 4 represents a program manager's email address
      const item4 = currentRow[managerEmailIndex];
      const item4CurrentCell = this.getSheetCell(this._sheet, cellRow, fourthColumnPosition);
      const item4ReportCell = this.getSheetCell(reportSheetBinding, cellRow, fourthColumnPosition);
      const item4HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, fourthColumnPosition);
      //if any of the required values are present, then the others msut be entered, so this cehck to see if any of the required values are present so any missing required values, in that case, can be flagged
      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0 || item4.length !== 0) {
        //iterated over each required value within the current row
        currentRow.forEach((val, index) => {
          //checks to see if the value is empty
          if (val === "") {
            switch (index) {
              //missing program name
              case 0:
                this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
                break;
              //missing program manager's first name
              case 1:
                this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                break;
              //missing program manager's last name
              case 2:
                this.setSheetCellBackground(item3CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item3ReportCell, this._lightRedHexCode);
                this.setSheetCellBackground(item3HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item3ReportCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item3HeaderCell, this._valuesMissingComment);
                break;
              //missing program manager's email address
              case 3:
                this.setSheetCellBackground(item4CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4ReportCell, this._lightRedHexCode);
                this.setSheetCellBackground(item4HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item4ReportCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item4HeaderCell, this._valuesMissingComment);
                break;
            }
          }
        });
      }
    }

  this._reportSummaryComments.push(this._foundFirstFourColumnsAndCheckedThemForMissingValuesComment);

  }
} 