/**
 * Use Dropbox API to copy a folder (and its contents).
 *
 * @param source  Path (from Dropbox root) to folder to copy.
 * @param destination  Path (from Dropbox root) to new folder.
 * @param accessToken  Dropbox-issued Access Token for PPS User.
 *                     Optional: if omitted, this function reads
 *                     it from the Project Properties.
 */
function Dropbox_copyFolder(source, destination, accessToken)
{
  if (accessToken === undefined)
  {
    accessToken = getProperty("PpsUserDropboxAccessToken");
  }
  
  const url = "https://api.dropboxapi.com/2/files/copy_v2"

  var dropboxParameters = 
  {
    'from_path': Dropbox_compliantPath(source),
    'to_path': Dropbox_compliantPath(destination)
  };
    
  var options = 
  {
    "method" : "post",
    "headers" : {"Authorization": "Bearer " + accessToken},
    "contentType" : "application/json",
    "muteHttpExceptions" : true,
    "payload" : JSON.stringify(dropboxParameters)
  };
  
  Logger.log(url);
  Logger.log(options);

  try
  {
    var result = UrlFetchApp.fetch(url, options);
    var responseText = result.getContentText();
    Logger.log(responseText);
    
    var responseJson = JSON.parse(responseText);
    /*
    Response is a dictionary.  If successful, it contains metadata like this:
        "metadata":
        {
            ".tag": "folder",
            "name": "Acme Tactile Dynamite",
            "path_lower": "/+ customer projects global/acme tactile dynamite",
            "path_display": "/+ Customer Projects Global/Acme Tactile Dynamite",
            "parent_shared_folder_id": "1481237362",
            "id": "id:M_W1dmORk2AAAAAAAAHUbg",
            "sharing_info":
            {
                "read_only": false,
                "parent_shared_folder_id": "1481237362",
                "traverse_only": false,
                "no_access": false
            }
        }
    If unsuccessful, it contains error information like this:
        "error_summary": "from_lookup/not_found/...",
        "error":
        {
            ".tag": "from_lookup",
            "from_lookup":
            {
                ".tag": "not_found"
            }
        }
    */
    
    if (responseJson["metadata"] != null)
    {
      // Success.
      Logger.log("Created " + responseJson.metadata.path_lower);
      return responseJson.metadata.path_lower;
    }
    
    if (responseJson.error_summary != null)
    {
      // Fail.
      Logger.log("Error copying folder: " + responseJson.error_summary);
      return null;
    }
  } 
  catch (e)
  {
    Logger.log(e);
    return null;
  }  
}



function testCopyFolder()
{
  Dropbox_copyFolder(
    '/+ Customer Projects Global/+ Templates & Reference Info/Template, RMA Folders',
    '/+ Customer Projects Global/Z. Shipped Projects/PQRS/PPS UK (Malcolm Hannah Limited)/RMA 5678');
}



/**
 * Non-recursively list contents of a given folder.
 *
 * @param path  Path to folder of interest.
 * @param accessToken  Optional Dropbox-issued API Access Token.  If 
 *                     omitted, this function reads one for PPS User 
 *                     from the Project Properties.
 * @return  JSON response as string if successful; otherwise null.
 */
function Dropbox_listFolder(path, accessToken)
{
  if (typeof accessToken === 'undefined')
  {
    accessToken = getProperty("PpsUserDropboxAccessToken");
  }
  
  const url = "https://api.dropboxapi.com/2/files/list_folder";
  
  var dropboxParameters = 
      {
        "path" : Dropbox_compliantPath(path),
        "recursive" : false
      };
    
  var fetchOptions = 
      {
        "method" : "post",
        "headers" : {"Authorization": "Bearer " + accessToken},
        "contentType" : "application/json",
        "muteHttpExceptions" : false,
        "payload" : JSON.stringify(dropboxParameters)
      };
  
  Logger.log(url);
  Logger.log(fetchOptions);

  try
  {
    var result = UrlFetchApp.fetch(url, fetchOptions);
    var reqReturn = result.getContentText();
    return reqReturn;
  } 
  catch (e)
  {
    Logger.log(e);
    return null;
  }
}



function testListFolder()
{
  var response = Dropbox_listFolder("/deliberately non-existent");
  Logger.log(response === null ? "Pass" : "Fail");
  
  response = Dropbox_listFolder("/software firmware controlled builds");
  Logger.log(response !== null ? "Pass" : "Fail");  
}





/**
 * Download one file.  Currently limited to text files.
 *
 * @param path  Path to file of interest.
 * @param accessToken  Optional Dropbox-issued API Access Token.  If 
 *                     omitted, this function reads one for PPS User 
 *                     from the Project Properties.
 * @return  Entire file as a single string, or null in failure cases.
 */
function Dropbox_download(path, accessToken)
{
  if (typeof accessToken === 'undefined')
  {
    accessToken = getProperty("PpsUserDropboxAccessToken");
  }
  
  const url = "https://content.dropboxapi.com/2/files/download";
  
  var dropboxParameters = 
      {
        "path" : Dropbox_compliantPath(path),
      };
    
  var fetchOptions = 
      {
        "method" : "post",
        "headers" : 
        {
          "Authorization": "Bearer " + accessToken,
          "Dropbox-API-Arg": JSON.stringify(dropboxParameters)
        },
        "contentType" : "text/plain",
        "muteHttpExceptions" : true
      };
  
  Logger.log(url);
  Logger.log(fetchOptions);

  try
  {
    var result = UrlFetchApp.fetch(url, fetchOptions);
    var reqReturn = result.getContentText();
    return reqReturn;
  } 
  catch (e)
  {
    Logger.log(e);
    return null;
  }
}


function testDownload()
{
  var t = Dropbox_download("/+ Customer Projects Global/Z. Shipped Projects/PQRS/P&G FingerTPS/Google Slide.url");
  if (t != null)
    Logger.log(t);
}

/**
 * Convert path to API-acceptable form.
 *
 * @inputPath  Path to convert. 
 * @return Dropbox-root-relative path with Unix-style separators.
 */
function Dropbox_compliantPath(inputPath)
{
  // Replace Windows path separators, if any.
  inputPath = inputPath.replace(/\\/g,"/");
  
  // Strip leading material such as "C:/Users/Jim" before
  // Dropbox root, if present.
  var segments = inputPath.split("/");
  for (let i = 0; i < segments.length; i++)
  {
    if (segments[i].startsWith("Dropbox") &&
        i + 1 < segments.length)
    {
      // segments[i] is "Dropbox (PPS)" or "Dropbox (work)" etc.
      // Take everything after this.
      inputPath = segments.slice(i + 1).join("/");
      break;
    }
  }
  
  // Dropbox API requires paths relative to the root.
  if (!inputPath.startsWith("/"))
  {
    inputPath = "/" + inputPath;
  }
  
  return inputPath;
}



function testDropbox_compliantPath()
{
  Logger.log(Dropbox_compliantPath(
    "C:\\Users\\PPS User\\Dropbox (PPS)\\+ Customer Projects Global\\Z. Shipped Projects\\PQRS\\P&G FingerTPS"));
  Logger.log(Dropbox_compliantPath(
    "Dropbox (PPS)/+ Customer Projects Global/Z. Shipped Projects/PQRS/P&G FingerTPS"));
  Logger.log(Dropbox_compliantPath(
    "/+ Customer Projects Global/Z. Shipped Projects/PQRS/P&G FingerTPS"));
  
  // Confirm that our compliant path is acceptable to Dropbox.
  const path = Dropbox_compliantPath("+ Customer Projects Global\\Z. Shipped Projects\\PQRS\\P&G FingerTPS");
  Logger.log(path);
  response = Dropbox_listFolder(path);
  Logger.log(response);
}



/**
 * Search Dropbox for any instance of given serial number.
 *
 * @param serialNumber  The number to search for (as a string).
 *                      This function removes any 'SN' prefix.
 * 
 * @return Array of strings, corresponding to search 'hits'.
 */
function Dropbox_searchSerialNumber(serialNumber)
{
  // Strip off any "SN" prefix.
  serialNumber = serialNumber.toLowerCase();
  if (serialNumber.startsWith('sn'))
  {
    serialNumber = serialNumber.slice(2, serialNumber.length);
  }
  
  var resultText = searchDropbox(serialNumber);
  if (resultText === null)
  {
    Logger.log("No search results for " + serialNumber);
    return null;
  }
 
  const resultJson = JSON.parse(resultText);
  
  // The result is a dictionary containing an item called "matches"
  // or an item called "error_summary".
  
  if ("error_summary" in resultJson)
  {
    Logger.log("Searching for " + serialNumber + 
               " gave Dropbox error: " + resultJson.error_summary);
    return;
  }
  
  if (!("matches" in resultJson) || resultJson.matches.length < 1)
  {
    Logger.log("Unexpected search results for " + serialNumber);
    return null;
  }

  // Matches is an array of dictionaries, each containing an item called
  // 'metadata' which is itself a dictionary (containing a metadata dictionary!)
  
  // Only collect instances like "SN1234".
  const serialText = "sn" + serialNumber;
  var hits = [];
  for (const match of resultJson.matches)
  {
    let item = match.metadata.metadata.path_lower;
    
    if (item.includes(serialText))
    {
      hits.push(item);
    }
  }
 
  return hits;
}



function testSearchSerialNumber()
{
  var x = Dropbox_searchSerialNumber("6085");
  Logger.log(x);
}


// Use the Dropbox search_v2 API to find relevant folders.
// Note: search_v2 has no non-recursive option so may find
// sub-folders within each project's main folder; so use
// the list_folder API instead.
function searchDropbox(searchTerm, accessToken)
{
  if (accessToken === undefined)
  {
    accessToken = getProperty("PpsUserDropboxAccessToken");
  }

  var url = "https://api.dropboxapi.com/2/files/search_v2";
  
  var dropboxParameters = 
      {
        "query" : searchTerm,
        "options" : 
        {
          "path" : "/+ customer projects global/z. shipped projects"
        }
      };
    
  var options = 
      {
        "method" : "post",
        "headers" : {"Authorization": "Bearer " + accessToken},
        "contentType" : "application/json",
        "muteHttpExceptions" : true,
        "payload" : JSON.stringify(dropboxParameters)
      };
  
  try
  {
    var result = UrlFetchApp.fetch(url, options);
    var reqReturn = result.getContentText();
    return reqReturn;
  } 
  catch (e)
  {
    Logger.log(e);
    return "Fail";
  }  
}






function runFilesApi(path, parameters)
{
  const baseUrl = "https://api.dropboxapi.com/2/files/";
  const url = baseUrl + path;
  const accessToken = getProperty("PpsUserDropboxAccessToken");
  
  var fetchOptions = 
      {
        "method" : "post",
        "headers" : {"Authorization": "Bearer " + accessToken},
        "contentType" : "application/json",
        "payload" : JSON.stringify(parameters),
        "muteHttpExceptions" : true
      };
  
  Logger.log(url);
  Logger.log(fetchOptions);

  try
  {
    var result = UrlFetchApp.fetch(url, fetchOptions);
    var responseText = result.getContentText();
    Logger.log("Response: " + responseText);    
    var responseJson = JSON.parse(responseText);
    if ("error_summary" in responseJson)
    {
      Logger.log(responseJson.error_summary);
      return null;
    }
    return responseJson;
  }
  catch (e)
  {
    Logger.log(e);
    return null;
  }
}






/**
 * Download a file straight to Dropbox.
 *
 * @param source  Url of file to download.
 * @param destination  Name (including root-relative path) to 
 *                     store downloaded file as.
 */
function Dropbox_storeFromUrl(source, destination)
{
  Logger.log("S: " + source, ", D: " + destination);
  var asyncJobId = null;
  
  var dropboxParameters = 
      {
        "path" : Dropbox_compliantPath(destination),
        "url" : source
      };
    
  var responseJson = runFilesApi("save_url", dropboxParameters);

  // Typically the download takes a while to complete.
  // Check the progress a few times, if necessary.
  for (let i = 0; i < 5; i++)
  {
    if (responseJson === null)
    {
      return false;
    }

    if (!(".tag" in responseJson))
    {
      Logger.log("Error: no '.tag' entry in response.");
      return false;
    }
    
    if (responseJson[".tag"] === "complete")
    {
      Logger.log("Complete");
      if ("size" in responseJson)
      {
        Logger.log(responseJson.size + " bytes stored.");
      }
      return true;
    }
    
    Logger.log("Status: " + responseJson[".tag"]);

    if (i === 0)
    {
      // First iteration.  Our JSON is the response to the initial save_url
      // call, which includes an id to check progress of async task.
      if (!(responseJson[".tag"] === "async_job_id"))
      {
        Logger.log("Unexpected status tag: " + responseJson[".tag"]);
        return false;
      }

      dropboxParameters = { "async_job_id": responseJson.async_job_id };
    }
    
    Logger.log("Waiting for completion");
    Utilities.sleep(500);
    
    responseJson = runFilesApi("save_url/check_job_status", dropboxParameters);
  }
  
  // Check result of latest status query.
  if (responseJson !== null &&
      ".tag" in responseJson &&
      responseJson[".tag"] === "complete")
  {
    return true;
  }
  
  return false;
}



function testStoreFromUrl()
{
  const source = "https://images.squarespace-cdn.com/content/v1/5b6af56a96d455ed3c06b55a/1555101150208-M9KYALBJDJBH8TIJC0CS/ke17ZwdGBToddI8pDm48kLkXF2pIyv_F2eUT9F60jBl7gQa3H78H3Y0txjaiv_0fDoOvxcdMmMKkDsyUqMSsMWxHk725yiiHCCLfrh8O1z4YTzHvnKhyp6Da-NYroOW3ZGjoBKy3azqku80C789l0iyqMbMesKd95J-X4EagrgU9L3Sa3U8cogeb0tjXbfawd0urKshkc5MgdBeJmALQKw/pressure-profile-team-photo-company.jpg?format=640w";
  const destination = "+ Customer Projects Global/Z. Shipped Projects/PQRS/PPS UK (Malcolm Hannah Limited)/test-download/group-photo.jpg";
  var response = Dropbox_storeFromUrl(source, destination);
  if (response === null)
  {
    Logger.log("Failed.");
    return;
  }
}






/**
 * Store a string as a text file.
 *
 * @param textToStore  A string, possibly multi-line, of text to be 
 *                     stored on Dropbox as a file.
 * @param destination  Desired name (including path) of new file.
 *
 * @return  ??Entire file as a single string, or null in failure cases.
 */
function Dropbox_storeText(textToStore, destination)
{
  const accessToken = getProperty("PpsUserDropboxAccessToken");
  
  const url = "https://content.dropboxapi.com/2/files/upload";
  
  var dropboxParameters = 
      {
        "path" : Dropbox_compliantPath(destination),
        "mode" : "overwrite",
        "mute" : true
      };
    
  var fetchOptions = 
      {
        "method" : "post",
        "headers" : 
        {
          "Authorization": "Bearer " + accessToken,
          "Dropbox-API-Arg": JSON.stringify(dropboxParameters)
        },
        "contentType" : "application/octet-stream",
        "muteHttpExceptions" : true,
        "payload" : textToStore
      };
  
  Logger.log(url);
  Logger.log(fetchOptions);

  try
  {
    var result = UrlFetchApp.fetch(url, fetchOptions);
    var reqReturn = result.getContentText();
    return reqReturn;
  } 
  catch (e)
  {
    Logger.log(e);
    return null;
  }
}


function testStoreText()
{
  var t = Dropbox_storeText(
    "Hello\nWorld",
    "Dropbox (PPS)\\+ Customer Projects Global\\Z. Shipped Projects\\PQRS\\PPS UK (Malcolm Hannah Limited)\\RMA 7892\\hello.txt");
  if (t === null)
  {
    Logger.log("Fail");
    return;
  }
  Logger.log(t);
}



/**
 * Create minimal Windows shortcut to URL.
 *
 * @param url  Target
 * @param filename  Name (including Dropbox Root-relative path)
 *                  of new shortcut file.  Typically ends with '.url'.
 */
function Dropbox_createShortcut(url, filename)
{
  // The only required field is 'InternetShortcut'
  const content = "[InternetShortcut]\nURL=" + url;
  Dropbox_storeText(content, filename);
}



function Dropbox_urlFromShortcut(filename)
{
  const content = Dropbox_download(filename);
  
  if (content === null)
  {
    Logger.log("Failed to download " + filename);
    return null;
  }
  
  //const content = "url = haha    \nProp3=19,11\n[InternetShortcut] \nURL=https://docs.google.com/presentation/d/16xw16zl1LqvYa_er9zTEVIOxr28MUpCM3xfQIuyBqpk/edit#slide=id.g2a812b5e25_3_26   \nIDList=";
  const matches = content.match(/url\s*=\s*([^\s]+)/i);
  if (matches.length < 2)
  {
    // Didn't find a URL
    return null;
  }
  
  // matches[0] is the entire match; matches[1] is the capture group
  const url = matches[1];
  
  if (!url.startsWith("http"))
  {
    // Reject other protocols, such as 'file' since they're
    // probably relative to the creator's computer.
    return null;
  }
  
  // We could do stuff like validate the link exists, remove direct page
  // links for Google Slides etc.  but there's currently no benefit.
  
  return url;
}



// Create a link to a file or folder.
// Link is not specific to current user.
function Dropbox_createLink(path)
{
  const accessToken = getProperty("PpsUserDropboxAccessToken");
  
  const url = "https://api.dropboxapi.com/2/sharing/create_shared_link_with_settings"

  var dropboxParameters = 
  {
    'path': Dropbox_compliantPath(path)
  };
    
  var options = 
  {
    "method" : "post",
    "headers" : {"Authorization": "Bearer " + accessToken},
    "contentType" : "application/json",
    "muteHttpExceptions" : true,
    "payload" : JSON.stringify(dropboxParameters)
  };
  
  Logger.log(url);
  Logger.log(options);

  try
  {
    var result = UrlFetchApp.fetch(url, options);
    var responseText = result.getContentText();
    Logger.log(responseText);
    
    var responseJson = JSON.parse(responseText);
    // Response is a dictionary.  If successful, it contains a 'url' key.
    
    if (responseJson["url"] != null)
    {
      // Success.
      Logger.log("New shared link: " + responseJson.url);
      return responseJson.url;
    }
    
    if (responseJson.error_summary != null)
    {
      // Fail.
      if (responseJson.error_summary.startsWith("shared_link_already_exists"))
      {
        // We failed because a link already exists.  We can
        // just use that.
        Logger.log("Using existing shared link");
        return responseJson.error.shared_link_already_exists.metadata.url;
      }
      Logger.log("Error creating link: " + responseJson.error_summary);
      return null;
    }
  } 
  catch (e)
  {
    Logger.log(e);
    return null;
  }  
}



/**
 * Convert path to form suitable for Windows File Explorer.
 *
 * @param path  Windows or Linux-style path to Dropbox content,
 *              optionally prefixed with Dropbox (PPS).
 *
 * @return Windows-compliant path, prefixed with "Dropbox (PPS)".
 */
function Dropbox_pathAsWindows(path)
{
  var compliantPath = "Dropbox (PPS)" + Dropbox_compliantPath(path);

  return compliantPath.replace(/\//g,"\\");
}



function testPathAsWindows()
{
  Logger.log(Dropbox_pathAsWindows("C:\\Users\\Malcolm\\Dropbox (PPS)\\+ Customer Projects Global\\Z. Shipped Projects\\DEF\\DEKA 2016-08-31 - (East Penn Mfg) 1072 Glove\\RMA 670"));
  Logger.log(Dropbox_pathAsWindows("/+ Customer Projects Global/Z. Shipped Projects/DEF/DEKA 2016-08-31 - (East Penn Mfg) 1072 Glove/RMA 670"));
}
