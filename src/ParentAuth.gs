/**
 * ParentAuth.gs
 *
 * Handles parent login via secretCode, validating against the
 * StudentDB sheet in the current Yearly Config file.
 */

function parentLogin(secretCode) {
  secretCode = (secretCode || "").trim().toUpperCase();

  // Basic format validation: must be exactly 6 alphanumeric characters
  if (!/^[A-Z0-9]{6}$/.test(secretCode)) {
    return {
      success: false,
      message: "Invalid secret code format.",
    };
  }

  const yearlyConfigFileId = getCurrentConfigFileId();
  if (!yearlyConfigFileId) {
    return {
      success: false,
      message:
        "Configuration error: currentConfigFile is not set in Survey Core Database.",
    };
  }

  const ss = SpreadsheetApp.openById(yearlyConfigFileId);
  const sheet = ss.getSheetByName("StudentDB");
  if (!sheet) {
    return {
      success: false,
      message:
        "Configuration error: StudentDB sheet not found in the Yearly Config file.",
    };
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    return {
      success: false,
      message: "No student records found. Please contact the school.",
    };
  }

  const header = values[0];
  const codeColIdx = header.indexOf("secretCode");
  if (codeColIdx === -1) {
    return {
      success: false,
      message:
        "Configuration error: 'secretCode' column not found in StudentDB.",
    };
  }

  const nameIdx = header.indexOf("studentName");
  const emailIdx = header.indexOf("studentEmail");
  const deptIdx = header.indexOf("department");

  for (let i = 1; i < values.length; i++) {
    const rowCode = String(values[i][codeColIdx]).trim().toUpperCase();
    if (rowCode === secretCode) {
      return {
        success: true,
        studentName: nameIdx >= 0 ? values[i][nameIdx] : "",
        studentEmail: emailIdx >= 0 ? values[i][emailIdx] : "",
        department: deptIdx >= 0 ? values[i][deptIdx] : "",
      };
    }
  }

  return {
    success: false,
    message:
      "Secret code is invalid. Please confirm the secret code with your child's homeroom teacher",
  };
}

