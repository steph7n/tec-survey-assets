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
  var coreDbId = SURVEY_CORE_DB_ID;

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
      surveyWindowText = formatSurveyWindow_(surveyStartDate, surveyEndDate);
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

function formatSurveyWindow_(startValue, endValue) {
  if (!startValue || !endValue) return "";

  var tz = Session.getScriptTimeZone();

  // Ensure we have Date objects
  var startDate = startValue instanceof Date ? startValue : new Date(startValue);
  var endDate = endValue instanceof Date ? endValue : new Date(endValue);

  // Fallback: if parsing failed, just concatenate the raw values
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
    return String(startValue) + " - " + String(endValue);
  }

  var datePattern = "dd MMM yyyy"; // e.g. 30 Nov 2025
  var startStr = Utilities.formatDate(startDate, tz, datePattern);
  var endStr = Utilities.formatDate(endDate, tz, datePattern);

  // Long timezone name, e.g. "Western Indonesia Time"
  var tzLabel = Utilities.formatDate(startDate, tz, "zzzz");

  return startStr + " - " + endStr + " (" + tzLabel + ")";
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
