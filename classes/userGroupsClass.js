class UserGroupsTempalte extends Template {
  constructor() {
    super();
    this._allowedEmailDomainsOptions = ['allowed email domains', 'domains', 'allowed domains', 'allowed email domain'];
    this._foundFirstColumnAndCheckedForMissingValuesComment = 'Success: Checked first column for missing values';
    this._invalidDomainNameComment = 'Invalid domain name/names found';
    this._invalidDomainNameFoundHeaderComment = 'Invalid domain name/names found';
    this._foundAllowedEmailDomainsColumnAndRanCheckForInvalidAllowedEmailDomainsComment = 'Success: ran check for invalid allowed email domains';
    this._didNotFindAllowedEmailDomainsColumnButRanCheckForInvalidAllowedEmailDomainsComment = 'Success: ran check for invalid allowed email domains, but did not find column';
    this._failedFirstColumnMissingValuesCheckMessage = "Failed: check first column for missing values";
    this._failedCheckNotRanForInvalidAllowedEmailDomains = 'Failed: check not ran for invalid allowed email domains';
  }
  checkFirstColumnForBlanks(reportSheetBinding) {
    const rowStartPosition = 2;
    const columnStartPosition = 1;
    const maxRows = this._sheet.getMaxRows();
    const totalColumsToCheck = 4;
    const range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
    const values = range.getValues();

    for (let row = 0; row < values.length; row += 1) {
      const zeroRangeAndFirstRow = 2;
      const groupTitleIndex = 0;
      const descriptionIndex = 1;
      const emailDomainsIndex = 2;
      const userGroupMembersIndex = 3;
      const headerRowPosition = 1;
      const cellRow = row + zeroRangeAndFirstRow;
      const currentRow = values[row];
      //this is the only required value ar erepresents a User Group title
      const item1 = currentRow[groupTitleIndex];
      const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
      const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
      const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
      //this represents an optional value, the User Group description
      const item2 = currentRow[descriptionIndex];
      //this represents an optional value, the User Group allowed email domains
      const item3 = currentRow[emailDomainsIndex];
      //this represents an optional value, the User Group memebr emails
      const item4 = currentRow[userGroupMembersIndex];
      //this checks to see if there are any values within the optional columns, and if so, makes sure the require User Group title is present
      if (item2.length !== 0  || item3.length !== 0 || item4.length !== 0) {
        //if there is a missing User Group title take the necessary error flagging actions
          if (item1 === "") {
                this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
          }
        }
      }

  this._reportSummaryComments.push(this._foundFirstColumnAndCheckedForMissingValuesComment);

  }
  validateDomainName(domainName) {
    return String(domainName)
      .toLowerCase()
      .match(
        /^((?:(?:(?:\w[\.\-\+]?)*)\w)+)((?:(?:(?:\w[\.\-\+]?){0,62})\w)+)\.(\w{2,6})$/);
  }
  checkForInvalidAllowedDomainNames(reportSheetBinding) {
    const headerRow = 1;
    //Runs for every column with a valid header implying that it contains allowed email domains
    this._allowedEmailDomainsOptions.forEach((headerTitle) => {
      const domainNameColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (domainNameColumnRange) {
        const domainNameColumnValues = this.getValues(domainNameColumnRange);
        const domainNameColumnPosition = domainNameColumnRange.getColumn();
        //Goes through each value, checks to see if the allowed email domain(s) provided are valid or a value is not present, and if not, flags the record as an error
        domainNameColumnValues.forEach((domainNameArray, index) => {
          const zeroRangeAndFirstRow = 2;
          const row = index + zeroRangeAndFirstRow;
          //If the domain name array is equal to a string, this implies that only one domain name was present within the cell
          if (typeof domainNameArray === 'string') {
            const currentDomainName = domainNameArray;
            if (currentDomainName !== "" && !this.validateDomainName(currentDomainName)) {
              const currentCell = this.getSheetCell(this._sheet, row, domainNameColumnPosition);
              const reportSheetCell = this.getSheetCell(reportSheetBinding, row, domainNameColumnPosition);
              const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, domainNameColumnPosition);

              //Line below for testing

              // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
              //testing line ends here
              this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
              this.setSheetCellBackground(currentCell, this._lightRedHexCode);
              this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
              this.insertCommentToSheetCell(reportSheetCell, this._invalidDomainNameComment);
              this.insertHeaderComment(mainSheetHeaderCell, this._invalidDomainNameFoundHeaderComment);
              this.setErrorColumns(reportSheetBinding, row);
            }
          //if the domain name array is not equal to a string, then there are multiple commas-separated domain names to check and the below methods will be used to do so
          } else {
            //domain name array holds all of the rows' domain name values
              domainNameArray.forEach((domainNameRowArray) => {
                //This create and array of allowed domain names within the current row and each domain name is validated
                domainNameRowArray.split(",").forEach(domainName => {
                const currentDomainName = domainName.trim();
                if (currentDomainName !== "" && !this.validateDomainName(currentDomainName)) {
                  const currentCell = this.getSheetCell(this._sheet, row, domainNameColumnPosition);
                  const reportSheetCell = this.getSheetCell(reportSheetBinding, row, domainNameColumnPosition);
                  const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, domainNameColumnPosition);

                  this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
                  this.setSheetCellBackground(currentCell, this._lightRedHexCode);
                  this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
                  this.insertCommentToSheetCell(reportSheetCell,this._invalidDomainNameComment);
                  this.insertHeaderComment(mainSheetHeaderCell, this._invalidDomainNameFoundHeaderComment);
                  this.setErrorColumns(reportSheetBinding, row);
                }
              });
            });
          }
        });
       }
     });
   this._reportSummaryComments.push(this._foundAllowedEmailDomainsColumnAndRanCheckForInvalidAllowedEmailDomainsComment);
    
  }
  get failedFirstColumnMissingValuesCheckMessage() {
    return this._failedFirstColumnMissingValuesCheckMessage = "Failed: check first column for missing values";
  }
  get failedCheckNotRanForInvalidAllowedEmailDomains() {
    return this._failedCheckNotRanForInvalidAllowedEmailDomains = 'Failed: check not ran for invalid allowed email domains';
  }
    

}
