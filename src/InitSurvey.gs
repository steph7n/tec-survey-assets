/**
 * InitSurvey.gs
 *
 * Handles survey initiation and creation of the Yearly Config file
 * from a template, including:
 * - Prompting for surveyYear, surveyStartDate, surveyEndDate, maxOpenEndedLength
 * - Copying the Yearly Config template
 * - Writing configuration values
 * - Generating student secret codes in StudentDB
 * - Storing the current Yearly Config file ID in the Survey Core Database Config sheet
 */

function initiateSurvey() {
  const ui = SpreadsheetApp.getUi();
  const coreFile = DriveApp.getFileById(SpreadsheetApp.getActive().getId());
  const coreFolder = coreFile.getParents().next(); // assumes exactly 1 folder

  const TEMPLATE_ID = "1n_Sm0-fLeHswgVOy5EUCs_1PcrZruw1hPVPnAvckIEQ";

  // ===== Helper: Cancel handling =====
  function checkCancel(result) {
    if (result.getSelectedButton() === ui.Button.CANCEL) {
      ui.alert("Survey initiation cancelled. No changes were made.");
      throw new Error("User cancelled the process.");
    }
  }

  // ===== Prompt 1: surveyYear =====
  let surveyYear;
  while (true) {
    const result = ui.prompt(
      "Survey Initiation",
      "Enter surveyYear (4 digits, between 2025 and 2100):",
      ui.ButtonSet.OK_CANCEL
    );
    checkCancel(result);

    const value = result.getResponseText().trim();

    if (/^\d{4}$/.test(value)) {
      const yearNum = Number(value);
      if (yearNum >= 2025 && yearNum <= 2100) {
        surveyYear = yearNum;
        break;
      }
    }

    ui.alert("Invalid surveyYear.\nPlease enter a number between 2025 and 2100.");
  }

  // ===== Helper: date validation =====
  function isValidDateInYear(str, year) {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(str)) return false;

    const [y, m, d] = str.split("-").map(Number);
    if (y !== year) return false;

    const date = new Date(y, m - 1, d);

    // Check that reconstructed date matches components (catches invalid dates)
    return (
      date.getFullYear() === y &&
      date.getMonth() === m - 1 &&
      date.getDate() === d
    );
  }

  // ===== Prompt 2: surveyStartDate =====
  let surveyStartDate;
  while (true) {
    const result = ui.prompt(
      "Survey Initiation",
      `Enter surveyStartDate (YYYY-MM-DD) within year ${surveyYear}:`,
      ui.ButtonSet.OK_CANCEL
    );
    checkCancel(result);

    const value = result.getResponseText().trim();

    if (isValidDateInYear(value, surveyYear)) {
      surveyStartDate = value;
      break;
    }

    ui.alert(`Invalid surveyStartDate.\nPlease enter a valid date in ${surveyYear}.`);
  }

  // ===== Prompt 3: surveyEndDate =====
  let surveyEndDate;
  while (true) {
    const result = ui.prompt(
      "Survey Initiation",
      `Enter surveyEndDate (YYYY-MM-DD) within year ${surveyYear}, at least 1 day AFTER ${surveyStartDate}:`,
      ui.ButtonSet.OK_CANCEL
    );
    checkCancel(result);

    const value = result.getResponseText().trim();

    if (isValidDateInYear(value, surveyYear)) {
      const start = new Date(surveyStartDate + "T00:00:00");
      const end = new Date(value + "T00:00:00");
      if (end.getTime() >= start.getTime() + 24 * 60 * 60 * 1000) {
        surveyEndDate = value;
        break;
      }
    }

    ui.alert(
      `Invalid surveyEndDate.\nMust be a valid date in ${surveyYear} and at least 1 day after ${surveyStartDate}.`
    );
  }

  // ===== Prompt 4: maxOpenEndedLength =====
  let maxOpenEndedLength;
  while (true) {
    const result = ui.prompt(
      "Survey Initiation",
      "Enter maxOpenEndedLength (integer between 10 and 1000):",
      ui.ButtonSet.OK_CANCEL
    );
    checkCancel(result);

    const value = result.getResponseText().trim();

    if (/^\d+$/.test(value)) {
      const num = Number(value);
      if (num >= 10 && num <= 1000) {
        maxOpenEndedLength = num;
        break;
      }
    }

    ui.alert("Invalid maxOpenEndedLength.\nEnter an integer between 10 and 1000.");
  }

  // ======================================
  // CAPTURE TEMPLATE SNAPSHOT FOR StudentDB & FacultyDB
  // ======================================
  const templateSpreadsheet = SpreadsheetApp.openById(TEMPLATE_ID);

  let templateStudentValues = null;
  const templateStudentSheet = templateSpreadsheet.getSheetByName("StudentDB");
  if (templateStudentSheet) {
    const tStuLastRow = templateStudentSheet.getLastRow();
    const tStuLastCol = templateStudentSheet.getLastColumn();
    if (tStuLastRow > 0 && tStuLastCol > 0) {
      templateStudentValues = templateStudentSheet
        .getRange(1, 1, tStuLastRow, tStuLastCol)
        .getValues();
    }
  }

  let templateFacultyValues = null;
  let templateFacultyValidations = null;
  const templateFacultySheet = templateSpreadsheet.getSheetByName("FacultyDB");
  if (templateFacultySheet) {
    const tFacLastRow = templateFacultySheet.getLastRow();
    const tFacLastCol = templateFacultySheet.getLastColumn();
    if (tFacLastRow > 0 && tFacLastCol > 0) {
      const tFacRange = templateFacultySheet.getRange(1, 1, tFacLastRow, tFacLastCol);
      templateFacultyValues = tFacRange.getValues();
      templateFacultyValidations = tFacRange.getDataValidations();
    }
  }

  // ======================================
  // CREATE NEW YEARLY CONFIG FILE
  // ======================================
  const newFileName = `Survey Config ${surveyYear}`;
  const templateFile = DriveApp.getFileById(TEMPLATE_ID);

  const newFile = templateFile.makeCopy(newFileName, coreFolder);
  const newSpreadsheet = SpreadsheetApp.openById(newFile.getId());

  // ======================================
  // UPDATE Survey Core Database: store this new Yearly Config File ID
  // ======================================
  const coreSS = SpreadsheetApp.getActive();
  const coreConfigSheet = coreSS.getSheetByName("Config");
  if (!coreConfigSheet) {
    throw new Error("Survey Core Database is missing a 'Config' sheet for storing currentConfigFile.");
  }
  coreConfigSheet.getRange("B1").setValue(newFile.getId());

  const studentSheet = newSpreadsheet.getSheetByName("StudentDB");
  const facultySheet = newSpreadsheet.getSheetByName("FacultyDB");

  // ======================================
  // APPLY SNAPSHOT TO NEW StudentDB
  // ======================================
  if (studentSheet && templateStudentValues) {
    const sRows = templateStudentValues.length;
    const sCols = templateStudentValues[0].length;
    studentSheet.getRange(1, 1, sRows, sCols).setValues(templateStudentValues);
  }

  // ======================================
  // APPLY SNAPSHOT TO NEW FacultyDB (values + validation)
  // ======================================
  if (facultySheet && templateFacultyValues && templateFacultyValidations) {
    const fRows = templateFacultyValues.length;
    const fCols = templateFacultyValues[0].length;
    const fRangeNew = facultySheet.getRange(1, 1, fRows, fCols);
    fRangeNew.setValues(templateFacultyValues);
    fRangeNew.setDataValidations(templateFacultyValidations);
  }

  // ======================================
  // WRITE VALUES INTO THE CONFIG SHEET
  // ======================================
  const configSheet = newSpreadsheet.getSheetByName("Config");
  if (!configSheet) {
    ui.alert("ERROR: The template is missing a sheet named 'Config'.");
    throw new Error("Missing Config sheet in template.");
  }

  const keys = ["surveyYear", "surveyStartDate", "surveyEndDate", "maxOpenEndedLength"];
  const values = [surveyYear, surveyStartDate, surveyEndDate, maxOpenEndedLength];

  const lastRow = configSheet.getLastRow();
  const data = configSheet.getRange(1, 1, lastRow, 2).getValues(); // col A & B

  keys.forEach((key, i) => {
    const rowIndex = data.findIndex(row => row[0] === key);
    if (rowIndex === -1) {
      ui.alert(`ERROR: Key '${key}' not found in Config sheet.`);
      throw new Error(`Key '${key}' not found.`);
    }
    configSheet.getRange(rowIndex + 1, 2).setValue(values[i]); // Column B
  });

  // ======================================
  // GENERATE SECRET CODES IN StudentDB (Column G)
  // ======================================
  if (studentSheet) {
    const studentValues = studentSheet.getRange(2, 1, studentSheet.getLastRow() - 1, 7).getValues();
    // studentValues[row][0] = A column, [6] = G column
    for (let i = 0; i < studentValues.length; i++) {
      const row = studentValues[i];
      if (row[0]) { // Column A not empty
        if (!row[6]) { // Only generate if Column G is blank
          const code = [...Array(6)].map(_ => "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789".charAt(Math.floor(Math.random() * 36))).join("");
          studentSheet.getRange(i + 2, 7).setValue(code);
        }
      }
    }
  }


  ui.alert(
    `Survey initialization complete.\n\nCreated: ${newFileName}\nLocation: Same folder as Survey Core Database`
  );
}

