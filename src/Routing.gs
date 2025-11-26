/**
 * Routing.gs
 *
 * Handles HTTP entry point (doGet) and HTML template loading
 * for different pages based on the 'page' query parameter.
 * Also exposes the base web app URL to HTML templates.
*/

function doGet(e) {
  const page = e && e.parameter && e.parameter.page ? e.parameter.page : "splash";

  const allowedPages = [
    "splash",
    "loginParent",
    "loginFaculty",
    "surveyHome",
    "survey",
    "thankyou",
    "admin"
  ];

  const fileToLoad = allowedPages.includes(page) ? page : "splash";

  // Create template and inject the web app URL
  const template = HtmlService.createTemplateFromFile(fileToLoad);
  template.baseUrl = ScriptApp.getService().getUrl();

  return template
    .evaluate()
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .setTitle("Tabgha Education Center School Survey");
}

/**
 * Returns the base URL of this web app deployment (the /exec URL).
 * Used by client-side code to build correct navigation links.
 */
function getWebAppUrl_() {
  return ScriptApp.getService().getUrl();
}

