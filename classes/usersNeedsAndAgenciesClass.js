////////////////////////////////////////////////////////////
//Superclass of Agency/Programs, Needs/Opportunities, and Users Templates
////////////////////////////////////////////////////////////

class UsersNeedsAndAgenciesTemplate extends Template {
  constructor() {
    super();
    this._stateHeaderOptions = ['state','state (ex: nh)', 'state (e.g. tn)','need state', 'opportunity state'];
    this._plainTextNumberFormat = `@STRING@`;
    this._phoneHeaderOptions = ['phone number', 'phone','mobile', 'cell phone', 'cell phone numbers','number','user phone','user phone number','home phone number','mobile phone number', 'mobile phone', 'cell phone number', 'mobile phone numbers', 'user phone numbers', 'user phone'];
    this._zipHeaderOptions = ['zipcode', 'zip code', 'postal','postal code', 'zip', 'zip codes', 'postal codes', 'user zip', 'user zip codes', 'user postal codes', 'user postal code', 'need zip', 'opportunity zip', 'opportunity zip code', 'need zip code', 'need postal code', 'opportunity postal code'];
    this._generalURLColumnOptions = ['website url', 'main site', 'webpage'];
    this._facebookLinkColumnOptions = ['facebook link', 'facebook', 'facebook page', 'fb', 'fb page', 'fb link', 'face book', 'face book page'];
    this._twitterColumnOptions = ['twitter link', 'twitter', 'twitter page', 'twitter url'];
    this._linkedInColumnOptions = ['linkedin link', 'linked in link', 'linked in', 'linkedin'];
    this._instagramColumnOptions = ['instagram link', 'instagram page', 'instagram'];
    this._youTubeColumnoptions = ['agency video (youtube or vimeo url)', 'youtube', 'youtube link', 'youtube video'];
    this._invalidLinkComment = 'Invalid link';
    this._invalidLinkHeaderComment = 'Invalid link/links found';
    this._stateTwoLetterCodeComment = 'State converted to two letter code'; 
    this._invalidPostalCodeComment = 'Invalid postal code'; 
    this._invalidPhoneNumberComment = 'Invalid phone number'; 
    this._invalidStateComment = 'Invalid State'; 
    this._invalidStateFoundHeaderComment = "Invalid State/States found";
    this._invalidPhoneNumbersFoundHeaderComment =  "Invalid Phone Number/Numbers Found";
    this._invalidPostalCodeFoundHeaderComment = "Invalid Postal Code/Codes Found";
    this._stateColumnFoundAndConversionFunctionRanComment = "Success: ran check to convert states to two-digit format";
    this._stateColumnNotFoundAndConversionRanComment = "Success: ran check to convert states to two-digit format, but did not find column";
    this._stateColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid states";
    this._stateColumnNotFoundAndValidationCheckRanComment = "Success: ran check for invalid states, but did not find column";
    this._phoneNumberColumnFoundAndValidationCheckRanComment = "Success: ran check for invalid phone numbers";
    this._postalCodeColumnFoundAndValidationCheckRan = "Success: ran check for invalid postal codes";
    this._postalCodeColumnNotFoundAndValidationCheckRan = "Success: ran check for invalid postal codes, but did not find column";
    this._generalLinkFoundAndValidationCheckRan = "Success: ran check for invalid general Program/Agency URL";
    this._twitterLinkFoundAndValidationCheckRan = "Success: ran check for invalid Twitter Program/Agency URL";
    this._facebookLinkFoundAndValidationCheckRan = "Success: ran check for invalid Facebook Program/Agency URL";
    this._instagramLinkFoundAndValidationCheckRan = "Success: ran check for invalid Instagram Program/Agency URL";
    this._youtubeLinkFoundAndValidationCheckRan = "Success: ran check for invalid YouTube Program/Agency URL";
    this._linkedInLinkFoundAndValidationCheckRan = "Success: ran check for invalid LinkedIn Program/Agency URL";
    this._failedCheckNotRanForGeneralURL = "Failed: check not ran for general URL";
    this._failedCheckNotRanForTwitterLink = "Failed: check not ran for Twitter Link";
    this._failedCheckNotRanForFacebookLink = "Failed: check not ran for Facebook Link";
    this._failedCheckNotRanForInstagramLink = "Failed: check not ran for Instagram Link";
    this._failedCheckNotRanForYouTubeLink = "Failed: check not ran for YouTube Link";
    this._failedCheckNotRanForLinkedInLink = "Failed: check not ran for LinkedIn Link";
    this._failedCheckNotRanForConvertingStatesToTwoLetterCodes = "Failed: check not ran for converting states to two-letter code";
    this._failedInvalidStatesCheckMessage = "Failed: check not ran for invalid states";
    this._failedInvalidPostalCodeCheck = "Failed: check not ran for invalid postal codes";
    this._failedInvalidPhoneNumbersCheck = "Failed: check not ran for invalid phone numbers";
    this._usStateToAbbreviation = {
      "Alabama": "AL",
      "Alaska": "AK",
      "Arizona": "AZ",
      "Arkansas": "AR",
      "California": "CA",
      "Colorado": "CO",
      "Connecticut": "CT",
      "Delaware": "DE",
      "Florida": "FL",
      "Georgia": "GA",
      "Hawaii": "HI",
      "Idaho": "ID",
      "Illinois": "IL",
      "Indiana": "IN",
      "Iowa": "IA",
      "Kansas": "KS",
      "Kentucky": "KY",
      "Louisiana": "LA",
      "Maine": "ME",
      "Maryland": "MD",
      "Massachusetts": "MA",
      "Michigan": "MI",
      "Minnesota": "MN",
      "Mississippi": "MS",
      "Missouri": "MO",
      "Montana": "MT",
      "Nebraska": "NE",
      "Nevada": "NV",
      "New Hampshire": "NH",
      "New Jersey": "NJ",
      "New Mexico": "NM",
      "New York": "NY",
      "North Carolina": "NC",
      "North Dakota": "ND",
      "Ohio": "OH",
      "Oklahoma": "OK",
      "Oregon": "OR",
      "Pennsylvania": "PA",
      "Rhode Island": "RI",
      "South Carolina": "SC",
      "South Dakota": "SD",
      "Tennessee": "TN",
      "Texas": "TX",
      "Utah": "UT",
      "Vermont": "VT",
      "Virginia": "VA",
      "Washington": "WA",
      "West Virginia": "WV",
      "Wisconsin": "WI",
      "Wyoming": "WY",
      "District of Columbia": "DC",
      "American Samoa": "AS",
      "Guam": "GU",
      "Northern Mariana Islands": "MP",
      "Puerto Rico": "PR",
      "United States Minor Outlying Islands": "UM",
      "U.S. Virgin Islands": "VI",
    }
    this._usFullStateNames = Object.keys(this._usStateToAbbreviation); 
    this._usStateAbbreviations = Object.values(this._usStateToAbbreviation);
  }
  get failedCheckNotRanForConvertingStatesToTwoLetterCodes() {
    return this._failedCheckNotRanForConvertingStatesToTwoLetterCodes;
  }
  get failedInvalidStatesCheckMessage() {
    return this._failedInvalidStatesCheckMessage;
  }
  get failedInvalidPostalCodeCheck() {
    return this._failedInvalidPostalCodeCheck;
  }
  get failedInvalidPhoneNumbersCheckk() {
    return this._failedInvalidPhoneNumbersCheck;
  }
  get failedCheckNotRanForGeneralURL() {
    return this._failedCheckNotRanForGeneralURL;
  }
  get failedCheckNotRanForInstagramLink() {
    return this._failedCheckNotRanForInstagramLink;
  }
  get failedCheckNotRanForLinkedInLink() {
    return this._failedCheckNotRanForLinkedInLink;
  }
  get failedCheckNotRanForTwitterLink() {
    return this._failedCheckNotRanForTwitterLink;
  }
  get failedCheckNotRanForFacebookLink() {
    return this._failedCheckNotRanForFacebookLink;
  }
  get failedCheckNotRanForYouTubeLink() {
    return this._failedCheckNotRanForYouTubeLink;
  }
  convertStatesToTwoLetterCode(reportSheetBinding) {
    //Runs for every column with a valid header implying that it contains states
    this._stateHeaderOptions.forEach((headerTitle) => {
      const stateColumnRange = this.getColumnRange(headerTitle, this._sheet);
      //Checks to see if a state's length is greater that the state abbreviation length, and if so, capitalizes its first letter and sets the rest of it to lower case so it can be mapped to an abbreviation
      if (stateColumnRange) {
        const stateAbbrvLength = 2;
        const stateColumnRangeValues = this.getValues(stateColumnRange).map(val => {
        const currentState = String(val);

        if (currentState.length > stateAbbrvLength) {
          return (currentState[0].toUpperCase() + currentState.slice(1).toLowerCase());
        } else {
          return currentState;
          }
        });

        const stateColumnRangePosition = stateColumnRange.getColumn();
        //Goes through all state values after they've been mapped and checks to see if a value can be mapped to a state's full-length name, if so, this value is replaced with the state's abbreviation
        stateColumnRangeValues.forEach((val, index) => {
          const zeroRangeAndFirstRow = 2;
          const currentState = String(val);
          const row = index + zeroRangeAndFirstRow;
          const currentCell = this.getSheetCell(this._sheet, row, stateColumnRangePosition);
          const currentReportCell = this.getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
          if (currentState.length > stateAbbrvLength && this._usFullStateNames.includes(currentState)) {
            currentCell.setValue(this._usStateToAbbreviation[currentState]);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._stateTwoLetterCodeComment);

          }
        });

    this._reportSummaryComments.push(this._stateColumnFoundAndConversionFunctionRanComment);

    }
   }); 
    
  }
  validateStates(reportSheetBinding) {
    //Runs for every column with a valid header implying that it contains states
    this._stateHeaderOptions.forEach((headerTitle) => {
      const stateColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (stateColumnRange) {
        const stateColumnRangeValues = this.getValues(stateColumnRange);

        const stateColumnRangePosition = stateColumnRange.getColumn();
        //Goes through each value, checks to see if it's not included within valid state abbreviations and a value is present, and if so, flags this as an error
        stateColumnRangeValues.forEach((val, index) => {
          const currentState = String(val).toUpperCase();
          if (!this._usStateAbbreviations.includes(currentState) && currentState.length !== 0) {
            const zeroRangeAndFirstRow = 2;
            const headerRow = 1;
            const row = index + zeroRangeAndFirstRow;
            const currentCell = this.getSheetCell(this._sheet, row, stateColumnRangePosition);
            const currentReportCell = this.getSheetCell(reportSheetBinding, row, stateColumnRangePosition);
            const mainSheetHeaderCell = this.getSheetCell(this._sheet, headerRow, stateColumnRangePosition);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentReportCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(currentReportCell, this._invalidStateComment);
            this.setSheetCellBackground(mainSheetHeaderCell, this._lightRedHexCode);
            this.insertHeaderComment(mainSheetHeaderCell, this._invalidStateFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding,row);
          }
        });

       this._reportSummaryComments.push(this._stateColumnFoundAndValidationCheckRanComment);
      }
    });
  }
  validatePhoneNumbers(number) { 
    return String(number)
      .toLowerCase()
      .match(
    /^\s*(?:\+?(\d{1,3}))?[-. (]*(\d{3})[-. )]*(\d{3})[-. ]*(\d{4})(?: *x(\d+))?\s*$/
      );
  }

  validatePostalCode(postalCode) { 
    return /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(postalCode);
  }
  validateGeneralURL(url) {
    return String(url)
    .toLowerCase()
    .match(
      /^(?:(?:https?|http):\/\/)?(?:(?!(?:10|127)(?:\.\d{1,3}){3})(?!(?:169\.254|192\.168)(?:\.\d{1,3}){2})(?!172\.(?:1[6-9]|2\d|3[0-1])(?:\.\d{1,3}){2})(?:[1-9]\d?|1\d\d|2[01]\d|22[0-3])(?:\.(?:1?\d{1,2}|2[0-4]\d|25[0-5])){2}(?:\.(?:[1-9]\d?|1\d\d|2[0-4]\d|25[0-4]))|(?:(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)(?:\.(?:[a-z\u00a1-\uffff0-9]-*)*[a-z\u00a1-\uffff0-9]+)*(?:\.(?:[a-z\u00a1-\uffff]{2,})))(?::\d{2,5})?(?:\/\S*)?$/
    );
  }
  validateTwitterLink(twitterLink) {
    return String(twitterLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*twitter.com\/.+/
      );
  }
  validateInstagramLink(instagramLink) {
    return String(instagramLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*instagram.com\/.+/
      );
  }
  validateYouTubeLink(youTubeLink) {
    return String(youTubeLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*youtube.com\/.+/
      );
  }
  validateLinkedInLink(linkedInLink) {
    return String(linkedInLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*linkedin.com\/.+/
      );
  }
  validateFaceBookLink(faceBookLink) {
    return String(faceBookLink)
      .toLowerCase()
      .match(
        /https:\/\/[www.]*facebook.com\/.+/
      );
  }
  //This method can be used to validate each type of url by passing in the report sheet, and the type of url to validate
  checkForInvalidURL(reportSheetBinding, urlType) {
    let headerOptions;
    let validationFunction;
    let reportSummaryCommentToUse;

    switch (urlType) {
      case 'general':
        headerOptions = this._generalURLColumnOptions;
        validationFunction = this.validateGeneralURL;
        reportSummaryCommentToUse = this._generalLinkFoundAndValidationCheckRan;
        break;
      case 'youtube':
        headerOptions = this._youTubeColumnoptions;
        validationFunction = this.validateYouTubeLink;
        reportSummaryCommentToUse = this._youtubeLinkFoundAndValidationCheckRan;
        break;
      case 'twitter':
        headerOptions = this._twitterColumnOptions;
        validationFunction = this.validateTwitterLink;
        reportSummaryCommentToUse = this._twitterLinkFoundAndValidationCheckRan;
        break;
      case 'linkedin':
        headerOptions = this._linkedInColumnOptions;
        validationFunction = this.validateLinkedInLink;
        reportSummaryCommentToUse = this._linkedInLinkFoundAndValidationCheckRan;
        break;
      case 'instagram':
        headerOptions = this._instagramColumnOptions;
        validationFunction = this.validateInstagramLink;
        reportSummaryCommentToUse = this._instagramLinkFoundAndValidationCheckRan;
        break;
      case 'facebook':
        headerOptions = this._facebookLinkColumnOptions;
        validationFunction = this.validateFaceBookLink;
        reportSummaryCommentToUse = this._facebookLinkFoundAndValidationCheckRan;
    }

    const headerRow = 1;
    //Runs for every columns with a valid header implying that it contains the url type being specified
    headerOptions.forEach((headerTitle) => {
      const linkColumn = this.getColumnRange(headerTitle, this._sheet);
      if (linkColumn) {
        const linkColumnRangeValues = this.getValues(linkColumn);
        const linkColumnRangePosition = linkColumn.getColumn();
        const mainSheetLinkHeaderCell = this.getSheetCell(this._sheet, headerRow, linkColumnRangePosition);
        //Goes through each link value, checks to see if it's valid or empty, and if not, flags the error
        linkColumnRangeValues.forEach((link, index) => {
          const zeroRangeAndFirstRow = 2;
          const currentLink = String(link);
          const row = index + zeroRangeAndFirstRow;

          if (currentLink !== "" && !validationFunction(currentLink)) {
            const currentCell = this.getSheetCell(this._sheet, row, linkColumnRangePosition);
            const reportSheetCell = this.getSheetCell(reportSheetBinding, row, linkColumnRangePosition);

            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetLinkHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidLinkComment);
            this.insertHeaderComment(mainSheetLinkHeaderCell,this._invalidLinkHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          });
        }
      });

    this._reportSummaryComments.push(reportSummaryCommentToUse);
  }
  setColumnToPlainTextNumerFormat(columnRangeBinding) {
    columnRangeBinding.setNumberFormat(this._plainTextNumberFormat);
  }

  checkForInvalidNumbers(reportSheetBinding) {
    const headerRow = 1;
    //Runs for every columns with a valid header implying that it contains the phone numbers
    this._phoneHeaderOptions.forEach((headerTitle) => {
      const phoneNumberRange = this.getColumnRange(headerTitle, this._sheet);
      if (phoneNumberRange) {
        const phoneNumberRangeValues = this.getValues(phoneNumberRange);
        const phoneNumberRangePosition = phoneNumberRange.getColumn();
        const mainSheetPhoneHeaderCell = this.getSheetCell(this._sheet, headerRow, phoneNumberRangePosition);
        //Goes through each phone number value, checks to see if it's valid or empty, and if not, flags the error
        phoneNumberRangeValues.forEach((number, index) => {
          const zeroRangeAndFirstRow = 2;
          const currentNumber= String(number);
          const row = index + zeroRangeAndFirstRow;

          if (currentNumber !== "" && !this.validatePhoneNumbers(currentNumber)) {
            const currentCell = this.getSheetCell(this._sheet, row, phoneNumberRangePosition);
            const reportSheetCell = this.getSheetCell(reportSheetBinding, row, phoneNumberRangePosition);

            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetPhoneHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidPhoneNumberComment);
            this.insertHeaderComment(mainSheetPhoneHeaderCell,this._invalidPhoneNumbersFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          });
      }
    });

    this._reportSummaryComments.push(this._phoneNumberColumnFoundAndValidationCheckRanComment);
  }
  checkForInvalidPostalCodes(reportSheetBinding) {
    const headerRow = 1;
    //Runs for every columns with a valid header implying that it contains postal codes
    this._zipHeaderOptions.forEach((headerTitle) => {
      const postalCodeColumnRange = this.getColumnRange(headerTitle, this._sheet);
      if (postalCodeColumnRange) {
        this.setColumnToPlainTextNumerFormat(postalCodeColumnRange)
        const postalCodeColumnRangeValues = this.getValues(postalCodeColumnRange);
        const postalCodeColumnRangePosition = postalCodeColumnRange.getColumn();
        const mainSheetPostalHeaderCell = this.getSheetCell(this._sheet, headerRow, postalCodeColumnRangePosition);
        //Goes throguh each phone number value, checks to see if it's valid or empty, and if not, flags the error
        postalCodeColumnRangeValues.forEach((code, index) => {
          const zeroRangeAndFirstRow = 2;
          const currentCode = String(code);
          const row = index + zeroRangeAndFirstRow;

          if (currentCode !== "" && !this.validatePostalCode(currentCode)) {
            const currentCell = this.getSheetCell(this._sheet, row, postalCodeColumnRangePosition);
            const reportSheetCell = this.getSheetCell(reportSheetBinding, row, postalCodeColumnRangePosition);
            this.setSheetCellBackground(reportSheetCell, this._lightRedHexCode);
            this.setSheetCellBackground(currentCell, this._lightRedHexCode);
            this.setSheetCellBackground(mainSheetPostalHeaderCell, this._lightRedHexCode);
            this.insertCommentToSheetCell(reportSheetCell, this._invalidPostalCodeComment);
            this.insertHeaderComment(mainSheetPostalHeaderCell, this._invalidPostalCodeFoundHeaderComment);
            this.setErrorColumns(reportSheetBinding, row);
            }
          });
         }
       });
    this._reportSummaryComments.push(this._postalCodeColumnFoundAndValidationCheckRan);
  } 

}


