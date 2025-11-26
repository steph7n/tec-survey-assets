/**
 * FacultyAuth.gs
 *
 * Handles Google account–based login for faculty and students,
 * including:
 * - Superadmin detection (owner of the Survey Core Database)
 * - Active faculty check (FacultyDB in Yearly Config)
 * - Active student check (StudentDB in Yearly Config)
 * Used primarily by loginFaculty.html.
 */

/**
 * Returns the superadmin email, defined as the owner of the Survey Core DB.
 * Falls back to a hard-coded email if needed (optional).
 */
function getSuperadminEmail() {
  try {
    // If your script is standalone and SURVEY_CORE_DB_ID is a constant:
    //   const coreSs = SpreadsheetApp.openById(SURVEY_CORE_DB_ID);
    //
    // If the web app is container-bound to the Survey Core DB, you can also do:
    //   const coreSs = SpreadsheetApp.getActiveSpreadsheet();

    const coreSs = SpreadsheetApp.getActiveSpreadsheet();
    const ownerEmail = coreSs.getOwner().getEmail();
    return ownerEmail ? ownerEmail.toLowerCase() : "";
  } catch (e) {
    Logger.log("getSuperadminEmail error: " + e);

    // Optional: fallback to your old hard-coded superadmin if something breaks
    // return "infosys@tabgha.education";

    return "";
  }
}

/**
 * Faculty/Student login validator for loginFaculty.html
 * Returns structured JSON for frontend.
 *
 * Possible result.status:
 * - "superadmin"
 * - "faculty"
 * - "student"
 * - "inactive"
 * - "notLoggedIn"
 */

function checkFacultyLoginStatus() {
  const email = Session.getActiveUser().getEmail();

  // Not logged in at all
  if (!email) {
    return {
      loggedIn: false,
      email: "",
      validDomain: false,
      status: "notLoggedIn",
    };
  }

  const domainValid = email.toLowerCase().endsWith("@tabgha.education");

  // Logged in but wrong domain
  if (!domainValid) {
    return {
      loggedIn: true,
      email: email,
      validDomain: false,
      status: "notLoggedIn",
    };
  }

  // -----------------------------
  // A. Superadmin (dynamic owner of Survey Core DB)
  // -----------------------------
  const superadminEmail = getSuperadminEmail();
  if (superadminEmail && email.toLowerCase() === superadminEmail) {
    return {
      loggedIn: true,
      email: email,
      validDomain: true,
      status: "superadmin",
    };
  }

  // -----------------------------
  // B. Active faculty?
  // -----------------------------
  try {
    if (isActiveFaculty(email)) {
      return {
        loggedIn: true,
        email: email,
        validDomain: true,
        status: "faculty",
      };
    }
  } catch (err) {
    Logger.log("Faculty check error: " + err);
  }

  // -----------------------------
  // C. Active student?
  // -----------------------------
  try {
    if (isActiveStudent(email)) {
      return {
        loggedIn: true,
        email: email,
        validDomain: true,
        status: "student",
      };
    }
  } catch (err) {
    Logger.log("Student check error: " + err);
  }

  // -----------------------------
  // D. Domain OK but NOT found
  // -----------------------------
  return {
    loggedIn: true,
    email: email,
    validDomain: true,
    status: "inactive",
  };
}

function isActiveFaculty(email) {
  const configFileId = getCurrentConfigFileId();
  const ss = SpreadsheetApp.openById(configFileId);
  const sheet = ss.getSheetByName("FacultyDB");
  if (!sheet) throw new Error("FacultyDB sheet missing in Yearly Config");

  const values = sheet.getDataRange().getValues();
  const header = values.shift();

  const emailIndex = header.indexOf("facultyEmail");
  if (emailIndex < 0) throw new Error("facultyEmail column missing");

  const target = email.trim().toLowerCase();

  return values.some(row => {
    const cell = (row[emailIndex] || "").toString().trim().toLowerCase();
    return cell === target;
  });
}

function isActiveStudent(email) {
  const configFileId = getCurrentConfigFileId();
  const ss = SpreadsheetApp.openById(configFileId);
  const sheet = ss.getSheetByName("StudentDB");
  if (!sheet) {
    throw new Error("StudentDB sheet missing in Yearly Config");
  }

  const values = sheet.getDataRange().getValues();
  if (values.length < 2) {
    // Only header or empty
    return false;
  }

  const header = values[0];
  const emailIndex = header.indexOf("studentEmail");
  if (emailIndex === -1) {
    throw new Error("studentEmail column missing in StudentDB");
  }

  const target = email.trim().toLowerCase();

  for (let i = 1; i < values.length; i++) {
    const cell = (values[i][emailIndex] || "").toString().trim().toLowerCase();
    if (cell === target) {
      return true;
    }
  }

  return false;
}

