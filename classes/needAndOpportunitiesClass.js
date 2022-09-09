class NeedsAndOpportunitiesTemplate extends UsersNeedsAndAgenciesTemplate {
    constructor() {
      super();
      this._foundFirstTwoColumnsAndCheckedThemForMissingValuesComment = "Success: checked for missing values in first two columns";
      this._failedFirstTwoColumnsMissingValuesCheckMessage = "Failed: check first two columns for missing values";
    }
    checkFirstTwoColumnsForBlanks(reportSheetBinding) {
      const agencyIndex = 0;
      const needIndex = 1;
      const rowStartPosition = 2;
      const firstColumnPosition = 1;
      const secondColumnPosition = 2;
      const maxRows = this._sheet.getMaxRows();
      const totalColumsToCheck = 2;
      const range = this._sheet.getRange(rowStartPosition, firstColumnPosition, maxRows, totalColumsToCheck);
      const values = range.getValues();
  
      for (let row = 0; row < values.length; row += 1) {
        const zeroRangeAndFirstRow = 2;
        const headerRowPosition = 1;
        const cellRow = row + zeroRangeAndFirstRow;
        const currentRow = values[row];
        //item 1 represents the agency/program name
        const item1 = currentRow[agencyIndex];
        const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, firstColumnPosition);
        const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, firstColumnPosition);
        const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, firstColumnPosition);
        //item 2 represents the need/opportunity title
        const item2 = currentRow[needIndex];
        const item2CurrentCell = this.getSheetCell(this._sheet, cellRow, secondColumnPosition);
        const item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, secondColumnPosition);
        const item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, secondColumnPosition);
        //if a program name or opportunity title is present then the other required value must be as well
        if (item1.length !== 0 || item2.length !== 0) {
          //iterates over the required values within the first row
          currentRow.forEach((val, index) => {
            //checks to see if the value is empty
            if (val === "") {
              switch (index) {
                //missing program/agency name
                case 0:
                  this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
                  break;
                //misssing opportunity/need title
                case 1:
                  this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                  break;
              }
            }
          });
        }
      }
  
    this._reportSummaryComments.push(this._foundFirstTwoColumnsAndCheckedThemForMissingValuesComment);
  
    }
  }
  