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
    let rowStartPosition = 2;
    let columnStartPosition = 1;
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
      let item2 = currentRow[1];
      let item3 = currentRow[2];

      if (item2.length !== 0  || item3.length !== 0) {
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
    let headerRow = 1;
    this._allowedEmailDomainsOptions.forEach((headerTitle) => {
      let domainNameColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (domainNameColumnRange) {
        let domainNameColumnValues = this.getValues(domainNameColumnRange);
        let domainNameColumnPosition = domainNameColumnRange.getColumn();

        domainNameColumnValues.forEach((domainNameArray, index) => {
          let row = index + 2;

          if (typeof domainNameArray === 'string') {
            let currentDomainName = domainNameArray;
            if (currentDomainName !== "" && !this.validateDomainName(currentDomainName)) {
              let currentCell = this.getSheetCell(this._sheet, row, domainNameColumnPosition);
              let reportSheetCell = this.getSheetCell(reportSheetBinding, row, domainNameColumnPosition);
              let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, domainNameColumnPosition);

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

          } else {
              domainNameArray.forEach((domainNameRowArray) => {
                domainNameRowArray.split(",").forEach(domainName => {
                let currentDomainName = domainName.trim();
                if (currentDomainName !== "" && !this.validateDomainName(currentDomainName)) {
                  let currentCell = this.getSheetCell(this._sheet, row, domainNameColumnPosition);
                  let reportSheetCell = this.getSheetCell(reportSheetBinding, row, domainNameColumnPosition);
                  let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, domainNameColumnPosition);

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


try {
  function userGroupsTemplateCheck() {
    const userGroupsTemplate = new UserGroupsTempalte();
    let reportSheet = userGroupsTemplate.createReportSheet();

    userGroupsTemplate.sheet.setName(userGroupsTemplate.templateName);
    

   userGroupsTemplate.ss.setActiveSheet(userGroupsTemplate.getSheet(userGroupsTemplate.templateName));

    


    userGroupsTemplate.setFrozenRows(userGroupsTemplate.sheet, 1);
    userGroupsTemplate.setFrozenRows(reportSheet, 1);

   try {
          userGroupsTemplate.removeWhiteSpaceFromCells();
        } catch (err) {
          Logger.log(err);
          userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedRemovedWhiteSpaceFromCellsComment;
          throw new Error(`White space not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

        try {
          userGroupsTemplate.removeFormattingFromSheetCells();
        } catch (err) {
          Logger.log(err);
          userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedRemovedFormattingFromCellsComment;
          throw new Error(`Formatting not removed from cells. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
        }
    
    userGroupsTemplate.clearSheetSummaryColumn(userGroupsTemplate.sheet);
    userGroupsTemplate.clearSheetSummaryColumn(reportSheet);
    userGroupsTemplate.setErrorColumnHeaderInMainSheet(userGroupsTemplate.sheet);

    try{
      userGroupsTemplate.checkFirstColumnForBlanks(reportSheet);
        } catch(err) {
            Logger.log(err);
            userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedFirstColumnMissingValuesCheckMessage;
            throw new Error(`Check not ran for missing values within first column. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

    try {
      userGroupsTemplate.formatCommaSeparatedLists(reportSheet);
    } catch (err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = programsAndAgenciesTemplate.failedRemoveWhiteSpaceAndMissingValuesFromCommaSeparatedListsMessage;
        throw new Error(`Did not remove whitespace and empty values from comma-separated list(s). Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    userGroupsTemplate.checkForInvalidCommaSeparatedEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedCheckNotRanForInvalidCommaSeparatedEmail;
        throw new Error(`Check not ran for invalid User Group Members emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the user group members column is titled "User Group Members (emails)" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 

    try {
  userGroupsTemplate.checkForInvalidAllowedDomainNames(reportSheet);
  } catch(err) {
      Logger.log(err);
      userGroupsTemplate.reportSummaryComments = userGroupsTemplate.failedCheckNotRanForInvalidAllowedEmailDomains;
      throw new Error(`Check not ran for invalid allowed domain names. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the allowed domain names column is titled "Allowed Email Domains" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

  try {
    userGroupsTemplate.setCommentsOnReportCell(reportSheet);
  } catch(err) {
    Logger.log(err);
    throw new Error(`Report sheet cell comment not added for summary of checks. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

   SpreadsheetApp.getUi().alert("User Groups Import Check Complete");
  }

} catch (err) {
    Logger.log(err);
    throw new Error(`An error occured the the User Groups import template check did not successfully run. Reason: ${err.name}: ${err.message}. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
  }

  
