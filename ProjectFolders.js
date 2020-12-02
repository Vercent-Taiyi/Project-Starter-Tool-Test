function isDistributor(candidate)
{
  distributors = ["Super Tooling", "SysCom", "PPS UK", 
                  "PPS KR", "PPS Korea", "WiseTouch"];
  
  var rx = new RegExp("\\s+", "g");
  
  for (const dist of distributors)
  {
    var distNoSpace = dist.replace(rx, "");
    var candidateNoSpace = candidate.replace(rx, "");
    if (distNoSpace.toLowerCase() === candidateNoSpace.toLowerCase())
    {
      return true;
    }
  }
  
  // Input parameter does not match any known distributor.
  return false;
}

function testIsDistributor()
{
  testOneDistributor("PPS UK", true);
  testOneDistributor("pps uk", true);
  testOneDistributor("PpSuK", true);
  testOneDistributor("pps   kr", true);
  testOneDistributor("PPSkorea", true);
  testOneDistributor("wise touch", true);
  testOneDistributor("supertoo ling", true);
  testOneDistributor("sys coM", true);
  testOneDistributor("tooling", false);
  testOneDistributor("sys", false);
}



function testOneDistributor(candidate, expectedResult)
{
  var message = "Testing " + candidate + ". Expecting " + expectedResult.toString() + ". ";
  
  if (isDistributor(candidate) != expectedResult)
  {
    message += "FAIL";
  }
  else
  {
    message += "PASS";
  }
  
  Logger.log(message);
}



/**
 * Search a folder (such as "ABC") for all projects for a company
 * (such as "Acme").
 *
 * @param accessToken  PPS User Dropbox API access token.
 * @param parentFolder  Folder to search (such as "ABC").
 * @param target  Company or distributor of interest (such as "Acme").
 *
 * @return  Array of lower-case full paths to folders whose names
 *          include the target company.
 */
function findCompanyFolders(accessToken, parentFolder, target)
{
  var companyFolders = [];
  
  Logger.log("Looking for " + target + " in " + parentFolder);
  
  response = Dropbox_listFolder(parentFolder, accessToken);
  
  if (response === null)
    return companyFolders;

  // Convert response (string) to JSON
  const parentListing = JSON.parse(response);

  for (const item of parentListing.entries)
  {
    if (item[".tag"] === "folder")
    {
      var lowerName = item.name.toLowerCase();
      if (lowerName.includes(target))
      {
        companyFolders.push(item.path_lower);
        Logger.log("CANDIDATE: " + item.path_lower);
      }
    }
  }
  
  return companyFolders;
}



/**
 * Find folder such as "ABC" or "PQRS" likely to contain info about
 * a target company or distributor.
 *
 * @param folderListing  Dropbox list_folder response (as JSON object)
 * @param target  The thing we're searching for, e.g. company name.
 *
 * @return Full lower-case path to folder that may contain target.
 */
function findLetterFolder(folderListing, target)
{
  var firstLetter = target.charAt(0);
  var firstLetterUpper = firstLetter.toUpperCase();
  
  for (const item of folderListing.entries)
  {
    if (item[".tag"] === "folder" &&
        item.name.length <= 4)
    {
      var upperName = item.name.toUpperCase();
      if (upperName === item.name)
      {
        // This looks like "ABC" or "PQRS" etc.
        if (upperName.includes(firstLetterUpper))
        {
          return item.path_lower;
        }
      }
    }
  }
}



/**
 * Find folder such as "Syscom" or "Verily" which are at the same
 * level as the letter folders such as "ABC".
 *
 * @param folderListing  Dropbox list_folder response (as JSON object)
 * @param target  The thing we're searching for, e.g. company name.
 *
 * @return Full lower-case path to folder that may contain target.
 */
function findDirectFolder(folderListing, target)
{
  for (const item of folderListing.entries)
  {
    if (item[".tag"] === "folder")
    {
      let lowerName = item.name.toLowerCase();
      if (lowerName.includes(target))
      {
        return item.path_lower;
      }
    }
  }
  
  // Didn't find it.
  return null;
}



/**
 * Strip or apply the standard shipped projects path prefix.
 *
 * @param folder  Path which may or may not include the prefix.
 *
 * @return  The specified path without the prefix, if it was originally
 *          included; otherwise the specified path with the prefix.
 */
function toggleFolderPrefix(folder)
{
  // Convert to lower case for consistency.
  folder = folder.toLowerCase();

  var prefix = "/+ customer projects global/z. shipped projects/";
  if (folder.startsWith(prefix))
  {
    // Strip prefix
    return folder.substring(prefix.length, folder.length);
  }
  
  // Apply prefix

  if (folder.startsWith("/"))
  {
    // Avoid double separator
    prefix = prefix.slice(0, prefix.length - 1);
  }
  return prefix + folder;
}



/**
 * Search the Shipped Projects archive on Dropbox for the named 
 * company and distributor.
 *
 * @note  This is invoked from script on some template HTML pages.
 *
 * @param company  Name of company.
 * @param distributor  Name of distributor.  Ignored if null.
 * @param stripPrefix  Option to return paths without the 
 *                     "/+ customer projects global/z. shipped projects"
 *                     prefix.
 *
 * @return  Array of lower-case paths to plausible folders, or null.
 */
function findShippedProjects(company, distributor, stripPrefix = false)
{
  var accessToken = getProperty("PpsUserDropboxAccessToken");
  var companyLower = company.toLowerCase();
  
  var response = Dropbox_listFolder(
    "/+ customer projects global/z. shipped projects",
    accessToken);
  
  if (response === null)
  {
    Logger.log("Dropbox_listFolder failed.");
    return null;
  }
  
  // Convert response (string) to JSON
  const folderListing = JSON.parse(response);
  
  // Some companies, e.g. Verily, get their own folder at
  // the same level as "ABC" etc.
  var parentFolder = findDirectFolder(folderListing, companyLower);

  if (parentFolder == null)
  {
    // We didn't find a company-specific top-level folder, so look
    // for it in one of the lettered folders.
    parentFolder = findLetterFolder(folderListing, companyLower);
  }

  if (parentFolder == null)
  {
    // This would only happen if someone moves or renames the "ABC" etc. folders.
    Logger.log("Didn't find dedicated or letter folder for " + company);
    return null;
  }

  // Now we have the relevant "ABC" etc. folder, look inside it.
  companyFolders = findCompanyFolders(accessToken, parentFolder, companyLower)
  if (companyFolders.length == 0)
  {
    // Try again with distributor
    if (distributor === undefined || isNullOrWhitespace(distributor))
    {
      Logger.log("No distributor to search for.");
      return null;
    }
    var distributorLower = distributor.toLowerCase();
    
    // Syscom gets its own folder at the same level as 'PQRS'
    parentFolder = findDirectFolder(folderListing, distributorLower);
    if (parentFolder == null)
    {
      // Find 'PQRS' for 'PPS UK', for example.    
      parentFolder = findLetterFolder(folderListing, distributorLower);
    }
    
    if (parentFolder == null)
    {
      Logger.log("Didn't find dedicated or letter folder for " +  distributor);
      return null;
    }
    
    // We now hope to find the company name associated with the distributor, for
    // example: "PPSUK(IBM)" in the "PQRS" folder.
    companyFolders = findCompanyFolders(accessToken, parentFolder, companyLower);
    if (companyFolders.length == 0)
    {
      Logger.log("No folders for " + company + " in " + parentFolder);
      return null;
    }
  }
  
  if (stripPrefix)
  {
    for (let i = 0; i < companyFolders.length; i++)
    {
      companyFolders[i] = toggleFolderPrefix(companyFolders[i]);
    }
  }
  
  Logger.log("fsp returning:");
  Logger.log(companyFolders);
  return companyFolders;
}



function testFSP()
{
  let projectFolders = findShippedProjects("kyushu", "syscom", true);
  for (const folder of projectFolders)
  {
    Logger.log(folder);
  }
}



// Note: to test one company+distributor combination, add them to the start of the
// arrays, then change "companies.length" to "1" in the for loop.
function testFindShippedProjects(company, distributor)
{
  const companies = ["Kyushu Univ", "Medtronic", "Chiaro", "IBM", "Verily", "P&G"];
  const distributors = ["Syscom", null, "PPS UK", "PPS UK", "PPS Korea", "Super tooling"];
  var i;
  
  for (i = 0; i < 1/*companies.length*/; i++)
  {
    let company = companies[i];
    let distributor = distributors[i];
    
    let projectFolders = findShippedProjects(company, distributor, true);
    if (projectFolders === null)
    {
      Logger.log("No matches for " + company);
      continue;
    }
    
    for (const folder of projectFolders)
    {
      Logger.log(folder);
    }
  }
}



function ProjectFolders_fromSerialNumber(serialNumber)
{
  try
  {
  var candidates = Dropbox_searchSerialNumber(serialNumber);
  if (candidates === null)
    return null;

  // candidates is an array of files and folders which mention the 
  // serial number.  We want just the containing project folder.
  var folders = [];
  for (const candidate of candidates)
  {
    // Example: "/+ customer projects global/z. shipped projects/mno/tto engineering 2014-12-30 - 881 tactilehead/3. test plan, procedure, data/sensor performance/sn6085_specs-02.18.2015.csv"
    let parts = candidate.split("/");
    if (parts.length > 4)
    {
      let partsToKeep = parts.slice(1, 5);
      let projectFolder = partsToKeep.join("/");
      if (folders.includes(projectFolder))
      {
        // Ignore this duplicate.
      }
      else
      {        
        folders.push(projectFolder);
      }
    }
  }
  
  return folders;
  }
  catch (e)
  {
    Logger.log(e.message);
  }   
}



function testFromSerialNumber()
{
  var t = ProjectFolders_fromSerialNumber("SN9538");
  Logger.log(t);
}



/**
 * @param folder  path may be short or long
 */
function ProjectFolders_getInfo(folder)
{
  var info = {};
  
  var path = null;
  const compliantPath = Dropbox_compliantPath(folder);
  const toggledPath = toggleFolderPrefix(compliantPath);
  if (toggledPath.length > compliantPath.length)
  {
    // We were passed a short path.  Use the full path.
    path = toggledPath;
  }
  else
  {
    // We were passed a full path, so use it.
    path = compliantPath;
  }
  
  info['path'] = path;
  info['shortPath'] = toggleFolderPrefix(path);
  info['link'] = Dropbox_createLink(path);

  var shortcuts = [];
  
  const jsonText = Dropbox_listFolder(path);
  if (jsonText !== null)
  {
    const listing = JSON.parse(jsonText);
    for (const item of listing.entries)
    {
      if (item[".tag"] !== "file" || 
          !item.path_lower.endsWith(".url"))
      {
        // Skip non-shortcut
        continue;
      }

      let url = Dropbox_urlFromShortcut(item.path_lower);
      if (url === null)
        continue;
      
      let shortcutInfo = {};
      
      if (url.includes("google.com/presentation"))
      {
        // Get generic URL (e.g. no direct slide)
        shortcutInfo['url'] = Slides_getCleanUrl(url);
        shortcutInfo['display'] = 'Slides';
      }
      else
      {
        // FIXME recognise more categories
        shortcutInfo['url'] = url;
        shortcutInfo['display'] = 'Unknown';
      }
      
      shortcuts.push(shortcutInfo);
    }
  }
  
  info['shortcuts'] = shortcuts;
  
  return info;
}
