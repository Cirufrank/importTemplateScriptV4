class UserTemplate extends UsersNeedsAndAgenciesTemplate {
  constructor() {
    super();
    this._duplicateEmailComment = 'Duplicate Email';
    this._duplicateEmailsFoundMessage = "Duplicate Email/Emails Found";
    this._genderExpandedComment = 'Gender option expanded';
    this._invalidGenderOptionComment = 'Invalid Gender Option';
    this._genderOptionsAbbreviationToFullObject = {
      "F":"Female",
      "M":"Male",
      "N/A":"",
    }
    this._genderOptionsAbbreviations = Object.keys(this._genderOptionsAbbreviationToFullObject);
    this._fullGenderOptions = ['female', 'male','prefer not to say', 'other', 'f','m'];
    this._foundEmailColumnAndCheckedForDuplicatesComment = "Success: checked emails column for duplicates";
    this._foundFirstThreeColumnsAndCheckedForMissingValuesComment = "Success: checked for missing values in first three columns";
    this._foundUserDateAddedAndBirthdayColumnsAndFormattedThemComment = "Success: formatted birthday column and user date added column";
    this._didNotFindUserDateAddedColumnButDidFindBirthdayColumnAndFormattedItComment = "Success: formatted birthday column and did not find user date added column";
    this._didNotFindBirthdayColumnButDidFindUserDateAddedColumnAndFormattedItComment = "Success: formatted user date added column and did not find birthday column";
    this._invalidGenderOptionsFoundHeaderComment = "Invalid Gender Option/Options found";
    this._foundGenderOptionsColumnAndRanCheckSuccessfullyComment = "Success: ran check for invalid gender options";
    this._didNotFindGenderOptionsColumnAndRanCheckSuccessfullyComment = "Success: ran check for invalid gender options, but did not find column";
    this._failedEmailDuplicatesCheckMessage = "Failed: check emails column for duplicates";
    this._failedFirstThreeColumnsBlankCheck = "Failed: check first three columns for missing values";
    this._failedFormatUserDateAddedAndBirthdayColumnsCheckMessage = "Failed: did not format user date added and birthday columns";
    this._failedInvalidGenderOptionsCheckMessage = "Failed: check not ran for invalid gender options";

  }
  get failedEmailDuplicatesCheckMessage() {
    return this._failedEmailDuplicatesCheckMessage;
  }
  get failedFirstThreeColumnsBlankCheck() {
    return this._failedFirstThreeColumnsBlankCheck;
  }
  get failedFormatUserDateAddedAndBirthdayColumnsCheckMessage() {
    return this._failedFormatUserDateAddedAndBirthdayColumnsCheckMessage;
  }
  get failedInvalidGenderOptionsCheckMessage() {
    return this._failedInvalidGenderOptionsCheckMessage;
  }
  checkForDuplicateEmails(reportSheetBinding) { 
    let headerRow = 1;
    //Runs for every column with a valid header implying that it contains emails
    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      const emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        const reportSheetEmailColumnRange = this.getColumnRange(headerTitle, reportSheetBinding);
        const emailColumnPosition = emailColumnRange.getColumn();
        const beginningOfEmailColumnRange = emailColumnRange.getA1Notation().slice(0,2); //0,2 represents beginning cell for column range i.e.C1
        const endOfEmailColumnRange = emailColumnRange.getA1Notation().slice(3); // slice 3 represents the end of the column range excluding the comma
        //Uses App Script's API to set a conditional foramt rule that will give all duplicate emails found a background of light red on the main sheet
        const duplicateEmailsRuleMainSheet = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
        .setBackground(this._lightRedHexCode)
        .setRanges([emailColumnRange])
        .build();
        //grabs the conditional format ruls for the active sheet
        const rules = this._sheet.getConditionalFormatRules();
        //pushes this new rule that has been build
        rules.push(duplicateEmailsRuleMainSheet);
        //sets the sheet's conditional format rules to the new rules that now have our added rule
        this._sheet.setConditionalFormatRules(rules);
        //uses App Script's API to set a conditional foramt rule that will give all duplicate emails found a background of light red on the report sheet
        const duplicateEmailsRuleReportSheet = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=COUNTIF(${beginningOfEmailColumnRange}:${endOfEmailColumnRange}, ${beginningOfEmailColumnRange})>1`)
        .setBackground(this._lightRedHexCode)
        .setRanges([reportSheetEmailColumnRange])
        .build();
        //grabs the conditional format rules for the active sheet
        const rules2 = reportSheetBinding.getConditionalFormatRules();
        //pushes this new rule that has been build
        rules2.push(duplicateEmailsRuleReportSheet);
        //sets the sheet's conditional format rules to the rules that now have our added rule
        reportSheetBinding.setConditionalFormatRules(rules2);

        //Sorts the report sheet and main sheet by their emails ascending
        reportSheetBinding.sort(emailColumnPosition, true);
        this._sheet.sort(emailColumnPosition, true);
    
      }
    });
    
    this._reportSummaryComments.push(this._foundEmailColumnAndCheckedForDuplicatesComment);

  }
  checkForLightRedHexCode(cell) {
    return cell.getBackground() === this._lightRedHexCode;
  }
  //It's best to use a custom conditional format rule to highlight duplicate emails, since this optimizes runtime, however, App Script does not allow us to set custom methods for conditional format rules so we are not able to makes sure both emails are highlighted
  //Do to this, in order to continue to allow users to sort the error column for all records, we must go back through after finding duplicate emails and set the background to light red for our other emails that are not yet highlighted and then set error columns for both duplicate email (since we are not able to do this at the time of highlighting throguh the conditional format rule)
  checkForDuplicateEmailsAndSetErrors(reportSheetBinding) {
    const headerRow = 1;
    //Runs for every column with a valid header implying that it contains emails
    this._emailColumnHeaderOptions.forEach((headerTitle) => {
      const emailColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (emailColumnRange) {
        const emailColumnValues = this.getValues(emailColumnRange);
        const emailColumnPosition = emailColumnRange.getColumn();
        //Goes through each value, checks to see if its background has been highlighted red and a value is present, and if so, sets the cell under its background color to red and sets its error comments
        emailColumnValues.forEach((email, index) => {
          const zeroRangeAndFirstRow = 2;
          const additionalRow = 1;
          const currentEmail = String(email).trim();
          const row = index + zeroRangeAndFirstRow;
          const nextRow = row + additionalRow;
          const currentCell = this.getSheetCell(this._sheet, row, emailColumnPosition);

          if (currentEmail.length > 0 && this.checkForLightRedHexCode(currentCell)) {
            // let nextCell = this.getSheetCell(this._sheet, nextRow, emailColumnPosition);
            const reportSheetCell = this.getSheetCell(reportSheetBinding, row, emailColumnPosition);
            const nextReportSheetCell = this.getSheetCell(reportSheetBinding, nextRow, emailColumnPosition)
            const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, emailColumnPosition);
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(nextReportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._duplicateEmailComment);
            this.insertCommentToSheetCell(nextReportSheetCell, this._duplicateEmailComment);
            this.insertHeaderComment(mainSheetHeaderCell, this._duplicateEmailsFoundMessage);
            this.setErrorColumns(reportSheetBinding, row);
            this.setErrorColumns(reportSheetBinding, nextRow);
            }
          });
        }
      });

      this._reportSummaryComments.push(this._ranCheckForDuplicateEmailsAndSetComments);
  }
  //Makes sure if there is at least one value within the first name, last name, or email columsn that no other columns within the row are empty
  checkFirstThreeColumnsForBlanks(reportSheetBinding) {
    const firstNameIndex = 0;
    const lastNameIndex = 1;
    const emailIndex = 2;
    const rowStartPosition = 2;
    const columnStartPosition = 1;
    const middleColumnPosition = 2;
    const thirdColumnPositon = 3;
    const maxRows = this._sheet.getMaxRows();
    const totalColumsToCheck = 3;
    const range = this._sheet.getRange(rowStartPosition, columnStartPosition, maxRows, totalColumsToCheck);
    const values = range.getValues();

    for (let row = 0; row < values.length; row += 1) {
      const headerRowPosition = 1;
      const cellRow = row + 2;
      const currentRow = values[row];
      //item 1 represents a first name value
      const item1 = currentRow[firstNameIndex];
      const item1CurrentCell = this.getSheetCell(this._sheet, cellRow, columnStartPosition);
      const item1ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, columnStartPosition);
      const item1HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, columnStartPosition);
      //item 2 represents a last name value
      const item2 = currentRow[lastNameIndex];
      const item2CurrentCell = this.getSheetCell(this._sheet, cellRow, middleColumnPosition);
      const item2ReportCurrentCell = this.getSheetCell(reportSheetBinding, cellRow, middleColumnPosition);
      const item2HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, middleColumnPosition);
      //item 3 represents an email value
      const item3 = currentRow[emailIndex];
      const item3CurrentCell = this.getSheetCell(this._sheet, cellRow, thirdColumnPositon);
      const item3ReportCell = this.getSheetCell(reportSheetBinding, cellRow, thirdColumnPositon);
      const item3HeaderCell = this.getSheetCell(this._sheet, headerRowPosition, thirdColumnPositon);

      if (item1.length !== 0 || item2.length !== 0  || item3.length !== 0) {
        currentRow.forEach((val, index) => {
          if (val === "") {
            switch (index) {
              //empty first name
              case 0:
                this.setSheetCellBackground(item1CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item1HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item1ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item1HeaderCell, this._valuesMissingComment);
                break;
              case 1:
              //empty last name
                this.setSheetCellBackground(item2CurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2ReportCurrentCell, this._lightRedHexCode);
                this.setSheetCellBackground(item2HeaderCell, this._lightRedHexCode);
                this.insertCommentToSheetCell(item2ReportCurrentCell, this._valuesMissingComment);
                this.setErrorColumns(reportSheetBinding, cellRow);
                this.insertHeaderComment(item2HeaderCell, this._valuesMissingComment);
                break;
              case 2:
              //empty email
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

  this._reportSummaryComments.push(this._foundFirstThreeColumnsAndCheckedForMissingValuesComment);

  }
  validateGenderOptions(reportSheetBinding) {
    //Runs for every column with a valid header implying that it contains gender options
    this._genderOptionsHeaderOptions.forEach((headerTitle) => {
      const genderOptionColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (genderOptionColumnRange) {
        const genderOptionColumnRangeValues = this.getValues(genderOptionColumnRange);

        const genderOptionColumnRangePoition = genderOptionColumnRange.getColumn();
        //Goes through each value, checks to see if the gender option is valid or a value is not present, and if not, flags the record as an error
        genderOptionColumnRangeValues.forEach((val, index) => {
          const currentGenderOption = String(val).toLowerCase();
          
          if (!this._fullGenderOptions.includes(currentGenderOption) && currentGenderOption.length !== 0) {
            const zeroRangeAndFirstRow = 2;
            const headerRow = 1;
            const row = index + zeroRangeAndFirstRow;
            const currentCell = this.getSheetCell(this._sheet, row, genderOptionColumnRangePoition);
            const currentReportCell = this.getSheetCell(reportSheetBinding, row, genderOptionColumnRangePoition);
            const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, genderOptionColumnRangePoition);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._invalidGenderOptionComment);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidGenderOptionsFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);

          }
        });
       }
      });

    this._reportSummaryComments.push(this._foundGenderOptionsColumnAndRanCheckSuccessfullyComment);
  }
} 
