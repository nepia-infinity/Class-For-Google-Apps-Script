function test_class() {
  const sheetUrl = "https://docs.google.com/spreadsheets/d/1GVHNPWBRh1SWltEDMv9eONWsyg9P81OtXhgh-dssa9A/edit?gid=0#gid=0";
  try {
    const utils = new SpreadsheetUtils(sheetUrl);
    const columnData = utils.extractColumnData("title", "url");
    console.log(columnData['url']);

    const selectedColumnIndex = utils.selectHeaderIndex("title", "url");
    const filtered = utils.getFilteredValues('ASD');

  }catch(error){
    console.warn(error);
  }
}
