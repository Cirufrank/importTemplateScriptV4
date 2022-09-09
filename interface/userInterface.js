////////////////////////////////////////////////////////////
//Creates the "Template Check" menu and buttons within Spreadsheet UI
////////////////////////////////////////////////////////////

function onOpen() {
    const ui = SpreadsheetApp.getUi()
    ui.createMenu('Import Teplate Checker')
      .addItem('Remove Header Cell Comments', 'removeHeaderCellComments')
      .addItem(`Remove Cell Comments`, 'removeCellComments')
      .addItem('Script Termination Page', 'openURL')
      .addItem('Check User Import Template', 'checkUserImportTemplate')
      .addItem('Check Individual Hours Import Template', 'checkIndvImportTemplate')
      .addItem('Check Responses and Hours Template', 'hoursAndResponsesCheck')
      .addItem('Check Need/Opportunity Responses Import Template', 'needAndOpportunityResponsesCheck')
      .addItem('Check Agencies/Programs Template', 'programsAndAgenciesTemplateCheck')
      .addItem('Check Needs/Opportunities Template', 'needsAndOpportunitiesTemplateCheck')
      .addItem('Check User Groups Template', 'userGroupsTemplateCheck')
      .addToUi();
  }