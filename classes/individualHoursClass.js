////////////////////////////////////////////////////////////
//Superclass of Hours/Responses Template
////////////////////////////////////////////////////////////

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
    const headerRow = 1;
    this.getSheetCell(this._sheet, headerRow, this._individualHoursTypeColumnPosition).setValue('Hours Type');
  }
  checkForHourTypeColumn() {
    if (this.getSheetCell(this._sheet, 1, this._individualHoursTypeColumnPosition).getValue()[0].toLowerCase().trim() !== 'h') {
      this._sheet.insertColumns(this._individualHoursTypeColumnPosition);
      this._reportSummaryColumnPosition += 1;
    };
  }

////////////////////////////////////////////////////////////
//Runs the checks for missing values when there are other 
//values present within the first three columns, and 
//also adds a 'Hours Type' column at the end of the sheet with 
//the word "Individual" in the cell for every row that has 
//values within in
////////////////////////////////////////////////////////////

    checkFirstThreeColumnsForBlanksAndAddIndvColumn(reportSheetBinding) {
      const rowStartPosition = 2;
      const columnStartPosition = 1;
      const middleColumnPosition = 2;
      const thirdColumnPositon = 3;
      const maxRows = this._sheet.getMaxRows();
      const totalColumsToCheck = 3;
      const range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
      const values = range.getValues();
    
      for (let row = 0; row < values.length; row += 1) {
        const zeroRangeAndFirstRow = 2;
        const emailIndex = 0;
        const dateIndex = 1;
        const hoursIndex = 2;
        const headerRowPosition = 1;
        const cellRow = row + zeroRangeAndFirstRow;
        const currentRow = values[row];
        //item 1 represents an email value
        const item1 = currentRow[emailIndex];
        const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
        const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
        const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
        const item1IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
        //item 2 represents a date value
        const item2 = currentRow[dateIndex];
        const item2CurrentCell = this.getSheetCell(this._sheet, cellRow, middleColumnPosition);
        const item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
        const item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, middleColumnPosition);
        const item2IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
        //item 3 represents a hours value
        const item3 = currentRow[hoursIndex];
        const item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
        const item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
        const item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);
        const item3IndvHoursCell = this.getSheetCell(this._sheet, cellRow, this._individualHoursTypeColumnPosition);
        //if any of the email, date, or hours values contain a value, then they all must within that row (this checks for that)
        if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
          item1IndvHoursCell.setValue(this._theWordIndividual);
          item2IndvHoursCell.setValue(this._theWordIndividual);
          item3IndvHoursCell.setValue(this._theWordIndividual);
          this.setIndividualHoursHeaderCellVal();
          //iterates through each email, date, and hours value within the current row, and if the current value is empty, based on the index the necessary background highlighting and error records will be set
          currentRow.forEach((val, index) => {
            if (val === "") {
              switch (index) {
                //empty email value
                case 0:
                  this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding,cellRow);
                  this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
          
                  break;
                //empty date value
                case 1:
                  this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                  this.setErrorColumns(reportSheetBinding, cellRow);
                  this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                  
                  break;
                //empty hours value
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
