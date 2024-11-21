function handleFormSubmission(e) {
  const directory = SpreadsheetApp.openById(DIRECTORY_SHEET_ID);

  const memberSheet = directory.getSheetByName("Members");

  const members = getOneDimensionalSpreadsheetData(directory, DIRECTORY_MEMBERS_RANGE);
  const chapters = getOneDimensionalSpreadsheetData(directory, DIRECTORY_CHAPTERS_RANGE);

  const chapterDropdown = createDropdown(DIRECTORY_CHAPTERS_RANGE, directory);
  const teamDropdown = createDropdown(DIRECTORY_TEAMS_RANGE, directory);
  const gradeDropdown = createDropdown(DIRECTORY_GRADES_RANGE, directory);

  const response = flattenResponse(e.response);
  handleResponse(
    response, 
    memberSheet, 
    members, 
    chapters,
    chapterDropdown,
    teamDropdown,
    gradeDropdown
  );
}

function handleResponse(response, memberSheet, members, chapters, chapterDropdown, teamDropdown, gradeDropdown){
  const parsedResponse = new FormResponse(response, chapters);
  
  const existingDataRow = members.indexOf(parsedResponse.name);

  if (existingDataRow != -1){
    fillData(parsedResponse.data(), existingDataRow + 2, memberSheet, chapterDropdown, teamDropdown, gradeDropdown);
    return; 
  } 

  const lastRow = memberSheet.getLastRow();
  const teamData = memberSheet.getRange(1, TEAM_COLUMN, lastRow).getValues();
  const chapterData = memberSheet.getRange(1, CHAPTER_COLUMN, lastRow).getValues();

  if (parsedResponse.team == "") {
    fillData(parsedResponse.data(), lastRow + 1, memberSheet, chapterDropdown, teamDropdown, gradeDropdown);
    return;
  }

  let insertRow = 0;
  let lastTeamRow = 0; 

  for (let i = 0; i < lastRow; i++) {
    if (teamData[i][0] == parsedResponse.team) {
      lastTeamRow = i + 1; 
      if (chapterData[i][0] == parsedResponse.chapter) {
        insertRow = i + 1; 
      }
    }
  } 

  if (insertRow == 0) {
    insertRow = lastTeamRow;
  }

  if (insertRow == 0) {
    insertRow = lastRow;
  }

  memberSheet.insertRowAfter(insertRow);
  fillData(parsedResponse.data(), insertRow + 1, memberSheet, chapterDropdown, teamDropdown, gradeDropdown);
}

function fillData(data, row, memberSheet, chapterDropdown, teamDropdown, gradeDropdown){
  const range = memberSheet.getRange(row, 1, 1, data.length);

  var currentData = range.getValues()[0];
  var backgrounds = range.getBackgrounds()[0];
  var dataValidations = range.getDataValidations()[0];

  for (let i = 0; i < currentData.length; i++){
    if (currentData[i] != "")
      continue;

    currentData[i] = data[i];

    if (data[i] == "")
      backgrounds[i] = MISSING_DATA_COLOR;

    if (i == CHAPTER_COLUMN - 1)
      dataValidations[i] = chapterDropdown;
    else if (i == TEAM_COLUMN - 1)
      dataValidations[i] = teamDropdown;
    else if (i == GRADE_COLUMN - 1)
      dataValidations[i] = gradeDropdown;
  }

  range.setValues([currentData])
    .setBackgrounds([backgrounds])
    .setDataValidations([dataValidations]);
}
