function getOneDimensionalSpreadsheetData(sheet, range) {
  const data = sheet.getRange(range).getValues();
  let oneDimensionalData = [];
  for (const entry of data){
    if (entry[0] == "")
      continue;

    oneDimensionalData.push(entry[0]);
  }

  return oneDimensionalData;
}

function flattenResponse(response){
  let flattenedResponse = [];
  const itemResponses = response.getItemResponses();

  for (const itemResponse of itemResponses) {
    flattenedResponse.push(itemResponse.getResponse());
  }

  return flattenedResponse;
}

function getFormResponses(form) {
  const responses = form.getResponses();
  let allResponses = [];
  for (const response of responses) {
    allResponses.push(flattenResponse(response));
  }

  return allResponses;
}

function createDropdown(range, directory){
  const sourceRange = directory.getRange(range); 
  const validation = SpreadsheetApp.newDataValidation()
    .requireValueInRange(sourceRange)
    .build();
  return validation;
}
