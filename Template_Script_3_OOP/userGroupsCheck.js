class UserGroupsTempalte extends Template {
  constructor() {
    super();
    this._foundFirstColumnAndCheckedForMissingValuesComment = 'Success: Checked first column for missing values';
    this._foundUserGroupMembersColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment = "Removed whitespace and empty values from User Group Members Emails on main sheet";
    this._foundUserGroupMembersColumnAndRemovedWhiteSpaceAndEmptyValuesComment = "Success: removed whitespace and empty values from Additional Contact Emails";
    this._didNotFindUserGroupsMembersColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment = "Success: removed whitespace and empty values from Additional Contact Emails, but did note find column";
    this._invalidEmailsFoundHeaderComment = "Invalid Email/Emails found"; //duplicated
    this._invalidDomainNameComment = 'Invalid domain name/names found';
    this._invalidDomainNameFoundHeaderComment = 'Invalid domain name/names found';
    this._foundUserGroupMembersColumnAndRanCheckForInvalidEmailsComment = "Success: ran check for invalid emails";
    this._didNotFindUserGroupMembersColumnButRanCheckForInvalidEmailsComment = "Success: ran check for invalid emails, but did not find column";
    this._foundAllowedEmailDomainsColumnAndRanCheckForInvalidAllowedEmailDomainsComment = 'Success: ran check for invalid allowed email domains';
    this._didNotFindAllowedEmailDomainsColumnButRanCheckForInvalidAllowedEmailDomainsComment = 'Success: ran check for invalid allowed email domains, but did not find column';
    this._foundAllowedEmailDomainsColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment = 'Removed whitespace and empty values from allowed email domains on main sheet';
    this._foundALlowedEmailDomainsColumnAndRemovedWhiteSpaceAndEmptyValuesComment = 'Success: removed whitespace and empty values from allowed emails domains';
    this._didNotFindAllowedEmailDomainsColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment = 'Success: removed whitespace and empty values from allowed emails domains, but did not find column';
    this._failedFirstColumnMissingValuesCheckMessage = "Failed: check first column for missing values";
    this._failedRemoveWhiteSpaceAndMissingValuesFromUserGroupMembersEmailsMessage = "Failed: did not remove whitespace and empty values from Additional Contact Emails";
    this._failedRemoveWhiteSpaceAndMissingValuesFromAllowedEmailDomainsMessage = "Failed: did not remove whitespace and empty values from allowed email domains";
    this._failedCheckNotRanForInvalidUserGroupMembersEmails = "Failed: check not ran for invalid additional contact emails";
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
  formatUserGroupMemberEmails(reportSheetBinding) {
    let headerRow = 1;
    let emailColumnRange = this.getColumnRange('User Group Members (emails)', this._sheet);

    if (emailColumnRange) {
    let emailColumnValues = this.getValues(emailColumnRange);
    let emailColumnPosition = emailColumnRange.getColumn();

    emailColumnValues.forEach((emailString, index) => {
      let row = index + 2;
      let emailArray = String(emailString).split(",");


        if (emailArray.length > 1) {
          let emailArrayNew = emailArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(", ");
          let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
          let reportSheetHeaderCell = this.getSheetCell(reportSheetBinding, headerRow, emailColumnPosition);
          let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
          let removedWhiteSpaceAndBlanksValue = emailArrayNew;
          currentCell.setValue(removedWhiteSpaceAndBlanksValue);

            this.insertHeaderComment(reportSheetHeaderCell, this._foundUserGroupMembersColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment);
            this.setSheetCellBackground(reportSheetHeaderCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
        }

        });

        this._reportSummaryComments.push(this._foundUserGroupMembersColumnAndRemovedWhiteSpaceAndEmptyValuesComment);

        } else {
      this._reportSummaryComments.push(this._didNotFindUserGroupsMembersColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment);
    }
    
  }
  checkForInvalidUserGroupMemberEmails(reportSheetBinding) {
    let headerRow = 1;
    let emailColumnRange = this.getColumnRange('User Group Members (emails)', this._sheet);

    if (emailColumnRange) {
      let emailColumnValues = this.getValues(emailColumnRange);
      let emailColumnPosition = emailColumnRange.getColumn();

    emailColumnValues.forEach((emailArray, index) => {
      let row = index + 2;

        if (typeof emailArray === 'string') {
          let currentEmail = emailArray;
          if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
            let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
            let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
            let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

            //Line below for testing

            // SpreadsheetApp.getUi().alert(`Invalid Email! ${currentEmail} ${reportSheetCell.getA1Notation()}`);
            //testing line ends here
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailsFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
          }

        } else {
            emailArray.forEach((emailRowArray) => {
              emailRowArray.split(",").forEach(email => {
                let currentEmail = email.trim();
                if (currentEmail !== "" && !this.validateEmail(currentEmail)) {
                let currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);
                let reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
                let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);

                this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
                this.setSheetCellBackground(currentCell, this._lightRedHexCode);
                this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(reportSheetCell, this._invalidEmailComment);
                this.insertHeaderComment(mainSheetHeaderCell, this._invalidEmailsFoundHeaderComment);
                this.setErrorColumns(reportSheetBinding, row);
                }
              });
          });

        }

        
        }
      );

      this._reportSummaryComments.push(this._foundUserGroupMembersColumnAndRanCheckForInvalidEmailsComment);
      
    } else {
      this._reportSummaryComments.push(this._didNotFindUserGroupMembersColumnButRanCheckForInvalidEmailsComment);
    }
    
  }
  formatAllowedEmailsDomainNames(reportSheetBinding) {
    let headerRow = 1;
    let allowedEmailDomainsColumnRange = this.getColumnRange('Allowed Email Domains', this._sheet);

    if (allowedEmailDomainsColumnRange) {
    let allowedEmailDomainsColumnValues = this.getValues(allowedEmailDomainsColumnRange);
    let allowedEmailDomainsColumnPosition = allowedEmailDomainsColumnRange.getColumn();

    allowedEmailDomainsColumnValues.forEach((domainString, index) => {
      let row = index + 2;
      let domainStringArray = String(domainString).split(",");


        if (domainStringArray.length > 1) {
          let domainStringArrayNew = domainStringArray.filter((val) => val.trim().length > 0).map((val) => val.trim()).join(", ");
          let currentCell = this.getSheetCell(this._sheet, row, allowedEmailDomainsColumnPosition);
          let reportSheetHeaderCell = this.getSheetCell(reportSheetBinding, headerRow, allowedEmailDomainsColumnPosition);
          let mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, allowedEmailDomainsColumnPosition);
          let removedWhiteSpaceAndBlanksValue = domainStringArrayNew;
          currentCell.setValue(removedWhiteSpaceAndBlanksValue);

            this.insertHeaderComment(reportSheetHeaderCell, this._foundAllowedEmailDomainsColumnAndRemovedWhiteSpaceAndEmptyValuesOnMainSheetHeaderComment);
            this.setSheetCellBackground(reportSheetHeaderCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
        }

        });

        this._reportSummaryComments.push(this._foundALlowedEmailDomainsColumnAndRemovedWhiteSpaceAndEmptyValuesComment);

        } else {
      this._reportSummaryComments.push(this._didNotFindAllowedEmailDomainsColumnButRanTheFunctionToRemoveWhiteSpaceAndEmptyValuesComment);
    }
    
  }
  validateDomainName(domainName) {
    return String(domainName)
      .toLowerCase()
      .match(
        /^((?:(?:(?:\w[\.\-\+]?)*)\w)+)((?:(?:(?:\w[\.\-\+]?){0,62})\w)+)\.(\w{2,6})$/);
  }
  checkForInvalidAllowedDomainNames(reportSheetBinding) {
    let headerRow = 1;
    let domainNameColumnRange = this.getColumnRange('Allowed Email Domains', this._sheet);

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
        }
      );

      this._reportSummaryComments.push(this._foundAllowedEmailDomainsColumnAndRanCheckForInvalidAllowedEmailDomainsComment);
      
    } else {
      this._reportSummaryComments.push(this._didNotFindAllowedEmailDomainsColumnButRanCheckForInvalidAllowedEmailDomainsComment);
    }
    
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
            userGroupsTemplate.reportSummaryComments = userGroupsTemplate._failedFirstColumnMissingValuesCheckMessage;
            throw new Error(`Check not ran for missing values within first column. Reason: ${err.name}: ${err.message}.  Please revert sheet to previous version, ensure the correct values are within the first four colums (i.e. "First Name, Last Name, and Email" is running a User Import Check [this would be for first three columns, but this checker is for four]), and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
        }

    try {
      userGroupsTemplate.formatUserGroupMemberEmails(reportSheet);
    } catch (err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = userGroupsTemplate._failedRemoveWhiteSpaceAndMissingValuesFromUserGroupMembersEmailsMessage;
        throw new Error(`Did not remove whitespace and empty values from User Group members emails. Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
    userGroupsTemplate.checkForInvalidUserGroupMemberEmails(reportSheet);
    } catch(err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = userGroupsTemplate._failedCheckNotRanForInvalidUserGroupMembersEmails;
        throw new Error(`Check not ran for invalid User Group Members emails. Reason: ${err.name}: ${err.message} at line. Please revert sheet to previous version, ensure the user group members column is titled "User Group Members (emails)" within its header column, and try again. If this test does not work, record this error message, revert sheet to previous version, and contact developer to fix.`);
      } 
      try {
      userGroupsTemplate.formatAllowedEmailsDomainNames(reportSheet);
    } catch (err) {
        Logger.log(err);
        userGroupsTemplate.reportSummaryComments = userGroupsTemplate._failedRemoveWhiteSpaceAndMissingValuesFromAllowedEmailDomainsMessage;
        throw new Error(`Did not remove whitespace and empty values from User Group members emails. Reason: ${err.name}: ${err.message} at line. Please record this error message, revert sheet to previous version, and contact developer to fix.`);
    }

    try {
  userGroupsTemplate.checkForInvalidAllowedDomainNames(reportSheet);
  } catch(err) {
      Logger.log(err);
      userGroupsTemplate.reportSummaryComments = userGroupsTemplate._failedCheckNotRanForInvalidAllowedEmailDomains;
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

  
