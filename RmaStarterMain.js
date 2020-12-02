// Use for retrieving secrets, such as access tokens.
function extractTaskId(linkOrNumber)
{
  //https://app.asana.com/0/134017349624135/1176874649688496
  var trimmed = linkOrNumber.trim();
  var taskId = null;
  if (/^\d+$/.test(trimmed))
  {
    // All numeric
    return trimmed;
  }
  
  // Probably full http link.  Task id is final part.
  var parts = trimmed.split("/");
  var finalPart = parts[parts.length - 1];
  if (/^\d+$/.test(finalPart))
  {
    // All numeric
    return finalPart;
  }

  // Failed
  return null;  
}



/**
 * Get Tech Support task info from Asana.
 *
 * @note  Invoked by the onChange handler of the form's
 *        Tech Support Task field.
 *
 * @param  linkOrNumber  Either the full URL of the Asana
 *                       task, or just the task id.
 *
 * @return  A dictionary with info from the customer's 
 *          original submission, or null in failure cases.
 */
function getTechSupportTask(linkOrNumber)
{
  var trimmed = linkOrNumber.trim();
  var taskId = null;
  if (/^\d+$/.test(trimmed))
  {
    // All numeric - assume this is the task id on its own.
    taskId = trimmed;
  }
  else
  {
    // Probably full http link, such as:
    //   https://app.asana.com/0/134017349624135/1176874649688496
    // Task id is final part.
    var parts = trimmed.split("/");
    var finalPart = parts[parts.length - 1];
    if (/^\d+$/.test(finalPart))
        taskId = finalPart;
  }
  if (taskId === null)
  {
    Logger.log("Can't find task id");
    return null;
  }
  
  var accessToken = getProperty("AsanaAccessToken");
  var taskJson = Asana_getTask(taskId, accessToken);
  if (taskJson === null)
  {
    Logger.log("Failed to get task.");
    return null;
  }

  // Cache task info to avoid re-fetching later.
  var cache = CacheService.getUserCache();
  cache.put(taskId, JSON.stringify(taskJson));

  var summary = {};
  // Don't bother with custom_fields since we can get Product
  // and Nature of Issue from the notes field.
  addNotes(summary, taskJson.data.notes);
  Logger.log("getTechSupportTask returning:");
  Logger.log(summary);
  return summary;
}



function testGetTechSupportTask()
{
  const summary = getTechSupportTask("1176874649688496");
  Logger.log("FINAL RESULT");
  Logger.log(summary);
}
             


// Find value for field in the task 'notes'.
function findField(fieldName, entireString)
{
  // The notes string is something like 
  // "Company Name:\nAcme\n\nName:\nJim Smith\n...etc."
  var re = new RegExp(fieldName + ":\\n(.*)\\n");
  var matches = re.exec(entireString);
  if (matches === null || matches.length < 2)
  {
    // Failed to find it.
    Logger.log("Failed to find " + fieldName);
    // Try a startsWith-style match
    re = new RegExp("\\n\\n" + fieldName + "[^:]+:\\n(\\w+)\\n");
    matches = re.exec(entireString);
    if (matches === null || matches.length < 2)
    {
      // Failed to find it.
      Logger.log("Failed to find field starting with " + fieldName);
      return "";
    }
  }
  // Found it.
  return matches[1];
}



function testFindField()
{
  const notesString = "Your Name:\nJim Smith\n\nEmail:\nmalcolm.hannah@pressureprofile.com\n\nCompany Name:\nAcme\n\nProduct:\nCustom System\n\nSerial Number(s):\n12345\n\nNature of the Issue:\nCan't do something I want to in Chameleon\n\nCan you replicate the problem?:\nIt's intermittent\n\nPlease describe the conditions when this issue occurs or how you can replicate the problem.   The more information you can provide, the quicker it will be for us to diagnose and fix the issue.:\nHello\n\nIf your issue is related to system performance or stability, please describe the applied pressure and duration, temperature, moisture, wireless RF noise environment, etc.:\nAgain\n\n———————————————\nThis task was submitted through Technical Support Request\nhttps://form.asana.com/?hash=59d9c4fdb0d5d9532f5ac5bf7635fb02eb6d0395c1db0ada32d4edf0683bbaf4&id=1175836475924952";
  findField("Please describe", notesString);
}



/**
 * Extract fields from task 'notes'.
 *
 * @summary  Dictionary to receive extracted notes.
 * @notesString  The task notes as one big string.
 */
function addNotes(summary, notesString)
{
  summary['companyName'] = findField("Company Name", notesString);
  summary['contactName'] = findField("Your Name", notesString);
  summary['contactEmail'] = findField("Email", notesString);
  summary['distributor'] = findField("Distributor", notesString);
  summary['product'] = findField("Product", notesString);
  summary['serialNumbers'] = findField("Serial Number\\(s\\)", notesString);
  summary['issueNature'] = findField("Nature of the Issue", notesString);
  summary['issueDescription'] = findField("Please describe", notesString);
}



/**
 * Decide which presentation to append RMA slides to.
 *
 * @param product  String from Tech Support task drop-down list.
 * @param projectSlides  Link found in original project folder.
 *
 * @return id of presentation to update.
 */
function deckToUpdate(product, projectSlides)
{
  const productLower = product.trim().toLowerCase();
  var slidesId = null;

  if (productLower.startsWith("tactileglove"))
  {
    return "1-jFboSTzBN5fEme2h-bcP8VJYRza4H72e-Oz8RH4f3w";
  }
  
  if (productLower.startsWith("fingertps"))
  {
    return "1zGiZFtEhCaUiqRln_yVYgA0YuBBmmUnIkpMf4YIf2qQ";
  }
  
  if (productLower.includes("spray"))
  {
    return "1JzXJ2H75VuNFxG05NE73sYx1WBfGkyNagssI723g0X8";
  }
  
  // Not a standard product, so use Slides link in original folder.
  if (!isNullOrWhitespace(projectSlides))
  {
    // URL is something like:
    // https://docs.google.com/open?id=1LzNOASc4orflxbU9Fb_rqWbDvntgJFL_shVn-AQHRe4
    const linkParts = projectSlides.split("?id=");
    if (linkParts.length == 2)
    {
      return linkParts[1];
    }
  }
  
  // We didn't find any valid deck to append to.
  return null;
}



/**
 * Invoked when Submit is clicked.
 *
 * @param formObject  The HTML form, as a JS object with members for
 *                    each of the field items with a 'name' attribute.
 *
 * @return  Text for the HTML page to display to the user.
 */
function folderSelected(formObject)
{
  var results = {};

  var optionAppendSlides = true;
  // Set next line true when debugging info passed from form.
  const debugFormOnly = false;
  
  Logger.log("RMA STARTING");
  Logger.log(formObject);
  
  if (debugFormOnly)
  {
    results.error = "Debug form only";
    return results;
  }
  
  // The form has either the full task URL, or just the id.  We need both.
  const taskId = extractTaskId(formObject.tech_support_link);
  if (taskId === null)
  {
    return "Currently this only supports the Tech Support Task route.";
  }
  const techSupportLink = Asana_getSupportUrl(taskId);
  results['supportTask'] = {'id': taskId, 'link': techSupportLink};

  var description;
  var product;
  var contactName;

  // Retrieve task info cached earlier.
  var cache = CacheService.getUserCache();
  taskJsonText = cache.get(taskId);
  if (taskJsonText !== null)
  {
    let taskJson = JSON.parse(taskJsonText);
    description = findField("error message etc", taskJson.data.notes);
    product = findField("Product", taskJson.data.notes);
    contactName = findField("\\nName", taskJson.data.notes);
  }
  else
  {
    // Not found in cache, so refetch.
    let summary = getTechSupportTask(taskId);
    description = summary.issueDescription;
    product = summary.product;
    contactName = summary.contactName;
  }

  // The form displays shortened folders but we need the full path.
  const selectedFolder = toggleFolderPrefix(formObject.selectedFolder);
  
  // Allocate and populate new row in RMA spreadsheet.  Assigns next free RMA number.
  const rmaNumber = Sheets_newRow(formObject.customer, description, selectedFolder, taskId);
  results['rmaNumber'] = rmaNumber;

  // Copy RMA folder-tree template into 'RMA nnn' in 
  // the project's original Dropbox folder.
  const newRmaFolder = selectedFolder + "/RMA " + rmaNumber;
  Logger.log(newRmaFolder);

  const templateFolder = '/+ Customer Projects Global/+ Templates & Reference Info/Template, RMA Folders';
  var createdFolder = Dropbox_copyFolder(templateFolder, newRmaFolder);
  if (createdFolder === null)
  {
    Logger.log("Failed to copy RMA template folder");
    return "Failed to copy RMA template folder";
  }
  Logger.log(createdFolder);
  
  results['rmaFolder'] = makeFolderCollection(newRmaFolder);
  results['commsFolder'] = makeFolderCollection(newRmaFolder + "/1. Customer Communication");
  results['photosFolder'] = makeFolderCollection(newRmaFolder + "/2. Incoming Photos");

  // Kick off the creation of an Asana project for this RMA.
  // This is asynchronous, typically completing within one minute.
  const newProjectName = "RMA " + rmaNumber + " " + formObject.customer;
  var duplicationJob = Asana_duplicateProject("rma", newProjectName);
  // We'll do other stuff while Asana servers duplicate the template.
  
  // Copy Asana Tech Support attachments into the new RMA folder.
  results['downloadedAttachments'] = downloadAsanaAttachments(
    results.supportTask.id,
    results.commsFolder.compliant);

  if (optionAppendSlides)
  {
    const presId = deckToUpdate(product, formObject.slidesLink);
    if (presId !== null)
    {
      const appended = Slides_appendRmaInfo(
        presId, rmaNumber, results.rmaFolder.compliant,
        results.rmaFolder.link, techSupportLink);
      if (appended)
      {
        // store link in return results
        // FIXME results['rmaSlide'] = slidelink
      }
    }
  }
  
  // In the new RMA folder, create a shortcut to the Asana Tech Support task
  var shortcutDestination = results.commsFolder.compliant + "/Tech Support Link.url";
  Dropbox_createShortcut(techSupportLink, shortcutDestination);
  
  // Wait (10 x 2 seconds) for the project duplication to complete.
  duplicationJob = Asana_waitForJob(duplicationJob.data.gid, 10, 2000);
  if (duplicationJob.data.status !== "succeeded")
  {
    Logger.log("Duplication failed or took too long");
  }
  var rmaProject = {};
  rmaProject.id = duplicationJob.data.new_project.gid;
  rmaProject.link = Asana_getProjectUrl(rmaProject.id);
  rmaProject.status = duplicationJob.data.status;
  results['rmaProject'] = rmaProject;
  customizeRmaProject(results);

  if (formObject.emailEnabler === "on")
  {
    sendEmail(
      formObject.email_recipient,
      contactName,
      rmaProject.id,
      rmaNumber,
      formObject.returnLocation);
    results.emailSent = true;
  }
  else
  {
    results.emailSent = false;
  }

  return results;  
}



function customizeRmaProject(results)
{
  const newProjectId = results.rmaProject.id;
  
  // By default, Asana creates new projects in the same Team as the template.
  // Move it to the Tech Support & RMA team.
  if (!Asana_moveProjectToRma(newProjectId))
  {
    Logger.log("Failed to move new project to RMA team.");
  }
  
  const searchTerms = ["received photos", "document resolution", 
                       "replicate", "notify customer"];
  const foundTasks = Asana_findProjectTasks(newProjectId, searchTerms);
  for (let i = 0; i < searchTerms.length; i++)
  {
    let foundTask = foundTasks[i];
    if (foundTask === undefined)
      continue;
    
    let items = [];
 
    items.push({text: "Tech Support Record", link: results.supportTask.link});
    
    switch (i)
    {
      case 0:
        // Add link to 'incoming photos' folder.
        items.push({text: "Incoming Photos Folder", link: results.photosFolder.link});
        break;
      case 1:
        // FIXME Add link to Google Slides
        break;
      case 2:
        items.push({text: "Customer Communications Folder", link: results.commsFolder.link});        
        break;
      case 3:
        // FIXME add customer contact details
        break;
      default:
        break;
    }
    
    Asana_setTaskDescription(
      foundTask,
      items);
  }
}



function makeFolderCollection(compliantPath)
{
  const windowsPath = Dropbox_pathAsWindows(compliantPath);
  const pathLink = Dropbox_createLink(compliantPath);
  const collection =
  {
    'compliant': compliantPath, 
    'windows': windowsPath,
    'link': pathLink
  };
  return collection;
}



function downloadAsanaAttachments(taskId, destinationFolder)
{
  const attachments = Asana_getTaskAttachments(taskId);
  if (attachments === null)
  {
    Logger.log("Failed to find task attachment info.");
    return null;
  }
  
  // attachments is an array of dictionaries, one per attachment.
  var downloadedAttachments = [];
  for (const attachment of attachments)
  {
    let destination = destinationFolder + "/" + attachment.name;
    if (Dropbox_storeFromUrl(attachment.download_url, destination))
    {
      downloadedAttachments.push(destination);
    }
  }
  
  return downloadedAttachments;
}




function testsendEmail()
{
  sendEmail(
      "malcolm.hannah+pps@gmail.com",
      "Malcolm Hannah",
      "1176696000640487",
      "123",
      "UK");
}



function sendEmail(
      recipient,
      contactName,
      rmaProjectId,
      rmaNumber,
      returnOffice)
{
  const subject = "PPS Return Merchandise Authorization";
  const bodyPlain = makeBodyText(contactName, rmaProjectId, rmaNumber, returnOffice);
  /*
    Note: we could also do fancy HTML stuff, in options.htmlBody
    const bodyHtml = "<html><body>This is the <b>fancy</b> text!</body></html>";
  */
  
  var replyAddress = "daniel.park@pressureprofile.com";
  if (returnOffice === "UK")
  {
    replyAddress = "dayi.zhang@pressureprofile.com";
  }
    
  const options = 
  {
    name: 'PPS Returns Team',
    replyTo: replyAddress
  }
    
  MailApp.sendEmail(
    recipient,
    subject,
    bodyPlain,
    options);
}



function testMakeBodyText()
{
  Logger.log(makeBodyText("Jim Smith", "1234567890", "987", "LA"));
}



function makeBodyText(contactName, rmaProjectId, rmaNumber, returnOffice)
{
  formUrl = "https://app.asana.com/0/" + rmaProjectId + "/form";
  var bodyLines = [
    "Hi " + contactName,
    "",
    "Further to your discussions with the PPS tech support team, we require the system to be returned to PPS for evaluation.",
    "",
    "Please complete this short form so we can ensure we have the correct contact and return address.",
    formUrl,
    "",
    "Please ship the complete system, in the appropriate protective packaging to:"];

  if (returnOffice === "UK")
  {
    bodyLines = bodyLines.concat([
      "Dayi Zhang",
      "RMA " + rmaNumber,
      "Inovo", 
      "121 George Street",
      "Glasgow",
      "G1 1RD",
      "UK"]);
  }
  else
  {
    bodyLines = bodyLines.concat([
      "Daniel Park",
      "RMA " + rmaNumber,  
      "Medical Tactile, Inc.",
      "5500 W Rosecrans Ave.",
      "Hawthorne, CA 90250",
      "USA"]);
  }
  
  bodyLines = bodyLines.concat([
    "",
    "Once we receive the package we will evaluate the system and get back to you as soon as possible.",
    "",
    "Thanks,"]);

  if (returnOffice === "UK")
  {
    bodyLines = bodyLines.concat([
      "Dayi"]);
  }
  else
  {
    bodyLines = bodyLines.concat([
      "Daniel"]);
  }
  
  return bodyLines.join("\n");
}
