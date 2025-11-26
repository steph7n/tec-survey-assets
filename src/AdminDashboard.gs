/**
 * AdminDashboard.gs
 *
 * Backend helpers for admin.html (superadmin landing page).
 */

/**
 * Returns summary info for the Superadmin dashboard.
 * Called from admin.html via google.script.run.getSuperadminDashboardData().
 */
function getSuperadminDashboardData() {
  ensureSuperadmin_();
  // If this script is container-bound to Survey Core DB:
  var coreDbId = SpreadsheetApp.getActiveSpreadsheet().getId();

  var currentConfigFileId = getCurrentConfigFileId(); // from Utils.gs
  var debugMode = false;
  var surveyYear = "";
  var lockStatus = "";
  var surveyStartDate = "";
  var surveyEndDate = "";
  var surveyStatus = "NOT STARTED";
  var surveyWindowText = "";
  var stats = {
    overall: 0,
    parents: 0,
    students: 0,
    faculty: 0,
    totalResponses: 0,
    validated: 0,
    errors: 0,
  };

  if (currentConfigFileId) {
    var ss = SpreadsheetApp.openById(currentConfigFileId);
    var configSheet = ss.getSheetByName("Config");
    if (!configSheet) {
      throw new Error("Config sheet missing in Yearly Config file.");
    }

    var cfg = getConfigAsObject(configSheet); // { surveyYear: "...", debugMode: "...", ... }

    debugMode = cfg.debugMode === true || String(cfg.debugMode).toUpperCase() === "TRUE";
    surveyYear = cfg.surveyYear || "";
    lockStatus = cfg.lockStatus || "";
    surveyStartDate = cfg.surveyStartDate || "";
    surveyEndDate = cfg.surveyEndDate || "";

    if (surveyStartDate && surveyEndDate) {
      surveyWindowText = surveyStartDate + " – " + surveyEndDate;
    }

    if (lockStatus === "OPEN") {
      surveyStatus = "ACTIVE";
    } else if (lockStatus === "CLOSED") {
      surveyStatus = "CLOSED";
    }

    // TODO (later): compute real stats from ResponseDB
    // For now, we leave defaults in `stats`.
  }

  return {
    debugMode: debugMode,
    loggedInEmail: Session.getActiveUser().getEmail() || "",
    surveyYear: surveyYear,
    surveyStatus: surveyStatus,
    surveyWindowText: surveyWindowText,
    configFileName: surveyYear
      ? surveyYear + "_Tabgha_Survey_Config"
      : "Yearly Config",
    configFileUrl: currentConfigFileId
      ? "https://docs.google.com/spreadsheets/d/" + currentConfigFileId + "/edit"
      : "",
    coreDbId: coreDbId,
    currentConfigId: currentConfigFileId || "",
    lockStatus: lockStatus,
    stats: stats,
  };
}

/**
 * Toggles lockStatus between OPEN and CLOSED in the current Yearly Config.
 * Called from admin.html via google.script.run.toggleSurveyOpenClosed().
 */
function toggleSurveyOpenClosed() {
  ensureSuperadmin_();
  var currentConfigFileId = getCurrentConfigFileId();
  if (!currentConfigFileId) {
    throw new Error("No current Yearly Config file found (currentConfigFile is empty).");
  }

  var ss = SpreadsheetApp.openById(currentConfigFileId);
  var configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    throw new Error("Config sheet missing in Yearly Config file.");
  }

  // Get key-value pairs
  var range = configSheet.getDataRange();
  var values = range.getValues(); // assuming no header, just key/value rows

  var current = "";
  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0]).trim();
    if (!key) continue;
    if (key.toLowerCase() === "lockstatus") {
      current = String(values[i][1] || "").toUpperCase();
      break;
    }
  }

  var next = current === "OPEN" ? "CLOSED" : "OPEN";

  // Write back new lockStatus
  for (var j = 0; j < values.length; j++) {
    var key2 = String(values[j][0]).trim();
    if (!key2) continue;
    if (key2.toLowerCase() === "lockstatus") {
      values[j][1] = next;
      break;
    }
  }
  range.setValues(values);

  return "Survey lockStatus changed from " + (current || "(empty)") + " to " + next + ".";
}

function ensureSuperadmin_() {
  var email = Session.getActiveUser().getEmail() || "";
  var superEmail = getSuperadminEmail();
  if (!email || !superEmail || email.toLowerCase() !== superEmail.toLowerCase()) {
    throw new Error("Unauthorized: Superadmin access only.");
  }
}