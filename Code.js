//Test
/**
 * Serve the Dashboard, or one of its subordinate pages,
 * such as 'Issue RMA'.
 *
 * @note This project can serve several different pages.
 *       Keeping them in the same project makes it easy
 *       for them to share helper scripts.
 *
 * @param e Event passed to doGet, with querystring.
 *          If present, the 'page' field specifies a 
 *          sub-ordinate page.
 * @returns Text (HTML) to be served.
 */

 //change this to see if there are any changes available
 //change again
function doGet(e)
{
  Logger.log("GET");
  Logger.log(JSON.stringify(e));
  
  const debugging = false;
  const defaultPage = 'DashboardTemplate';
  var selectedPage;
  
  if (!e.parameter.page)
  {
    // No specific page requested, so serve the Dashboard.
    selectedPage = defaultPage;
  }
  else
  {
    selectedPage = e.parameter['page'];
  }

  var template;
  
  try
  {
    template = HtmlService.createTemplateFromFile(selectedPage);
  }
  catch (error)
  {
    // Probably non-existent page requested.
    Logger.log(JSON.stringify(error));
    template = HtmlService.createTemplateFromFile(defaultPage);
  }
  
  // Attach parameters so each form may pre-populate some fields.
  template.httpGetParams = e.parameter;
  
  if (debugging)
  {
    Logger.log(template.getCode());
  }
  return template.evaluate();
}



/**
 * Use this to debug (single-step) doGet.
 */
function testDoGet()
{
  var e = {"contextPath":"","parameter":{"page":"RmaMainFormz"},"queryString":"page=RmaMainFormz","contentLength":-1,"parameters":{"page":["RmaMainFormz"]}};
  doGet(e);
}



/**
 * Get the URL for the Google Apps Script running as a WebApp.
 */
function getScriptUrl()
{
  var url = ScriptApp.getService().getUrl();
  Logger.log("GSU " + url);
  return url;
}



/**
 * Supply one of this project's html files.
 *
 * @note  This allows stylesheets or javascript functions to be shared 
 *        by multiple script or html files.
 *
 * @return  Content of an html file.
 */
function include(filename)
{
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}



/**
 * Get property (key-value pair) from script's store.
 *
 * @param key
 *
 * @return The value of the specified key, or null in failure cases.
 */
function getProperty(key)
{
  var scriptProperties = PropertiesService.getScriptProperties();
  try
  {
    return scriptProperties.getProperty(key);
  }
  catch (error)
  {
    Logger.log(error);
    return null;
  }
}



function testGetProperty()
{
  var answer;
  answer = getProperty("rubbish");
  Logger.log(answer === null ? "Pass" : "Fail");
  answer = getProperty("PpsUserDropboxAccessToken");
  Logger.log(answer === null ? "Fail" : "Pass");
  answer = getProperty("AsanaAccessToken");
  Logger.log(answer === null ? "Fail" : "Pass");
}



function isNullOrWhitespace(str)
{
  return str === null || (/^\s*$/).test(str);
}



/**
 * Handle POST requests.  Other PPS GAS projects use POST 
 * requests to access this project's Asana, Dropbox, Slides
 * and Sheets services.
 *
 * @param e Event object containing POST data etc.
 *
 * @return Text (JSON) with results of the service function.
 */
function doPost(e)
{
  Logger.log("Post received.");
  Logger.log(e.postData.contents);  

  var response = {};
  var result;
  var message;
  
  try
  {
    const postData = JSON.parse(e.postData.contents);
    
    if ("createTask" in postData)
    {
      result = Service_createTask(postData.createTask);
      message = "OK";
    }
    response.result = result;
    response.message = message;
  }
  catch (error)
  {
    response.error_message = error.message;
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);  
}



function Service_createTask(taskInfo)
{
  if (!("assignee" in taskInfo) ||
      !("title" in taskInfo) ||
      !("section" in taskInfo) ||
      !("notes" in taskInfo))
  {
    // Insufficient info
    Logger.log("Missing fields in createTask info.");
    return null;
  }
  
  // FIXME convert text assignee etc. into gid here?
  // FIXME allow no assignee.
  
  return Asana_createTask(taskInfo.title, taskInfo.notes, taskInfo.section, taskInfo.assignee);
}




/**
 * Create standard project
 * Three things to do when the button is clicked
 * Find and replace the content, and save as a new google slide
 * Copy the template folder, change the folder name, add the google slide
 * Copy the Asana template and change the project name
 */
function userClicked(userInfo)
{
  //Get the template google slide ID
  const templateId = "1q7zyEP7aCuq-OoN7iWPkwZhzV2b0waZzYLCON4-IjtA";
  var newId = null;
  
    //Generate a new google slide - In my own google drive - Authorisation needed?
    const newName = userInfo.customerName + " (" + userInfo.distributorName + ") " + userInfo.projectName;
    var newFile = cloneFile(templateId, newName);
    newId = newFile.getId();
    Logger.log("Created " + newName + ", id: " + newId);
    
    var ss = SlidesApp.openById(newId);
    ss.replaceAllText("{{Customer}}",userInfo.customerName);
    ss.replaceAllText("{{Distributor}}",userInfo.distributorName);
    ss.replaceAllText("{{Project}}",userInfo.projectName);
    ss.replaceAllText("{{ContactName}}",userInfo.contactName);
    ss.replaceAllText("{{ContactNumber}}",userInfo.contactNumber);
    ss.replaceAllText("{{ContactEmail}}",userInfo.contactEmail);
    ss.replaceAllText("{{ContactAddress}}",userInfo.contactEmail);
    ss.replaceAllText("{{SensorType}}",userInfo.sensorType);
    ss.replaceAllText("{{FSR}}",userInfo.FSR);
    ss.replaceAllText("{{ScanRate}}",userInfo.scanRate);
    ss.replaceAllText("{{InterfaceType}}",userInfo.interfaceType);
    ss.saveAndClose();
  
    //Create dropbox folder
    destinationPath = "/+ Customer Projects Global/" + newName;
    Dropbox_copyFolder(
    '/+ Customer Projects Global/+ Templates & Reference Info/Template, Existing Design Project Folders',
    destinationPath);
  
    //Copy the new slide into the Dropbox 
    Dropbox_createShortcut(
    "https://docs.google.com/presentation/d/" + newId,
    destinationPath + "/Google Slide 2.0.url");
  
    //Duplicate the CSP template 
    Asana_duplicateProject(
    "csp",
    newName);

  
}
