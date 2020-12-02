function Sheets_newRow(customer, description, parentFolder, techSupportTask) 
{
  var newRmaNumber = null;
  rmaSheetId = "1hXnJmGlIYshJJjkCZetyVdGDPWNOB35lF7dOkgQgtZg";
  
  try
  {
    var ss = SpreadsheetApp.openById(rmaSheetId);
    var spreadsheetName = ss.getName();
    Logger.log("Found " + spreadsheetName);
    
    var sheet = ss.getSheetByName("RMA");
    if (sheet == null)
    {
      Logger.log("No sheet named 'RMA' in " + spreadsheetName);
      return null;
    }
    
    var columnIndices = getColumnIndices(ss);
    
    var lastRowIndex = sheet.getLastRow();
    var lastRmaNumberRange = sheet.getRange(lastRowIndex, 1);
    var lastRmaNumber = lastRmaNumberRange.getValue();
    Logger.log(lastRmaNumber);

    newRmaNumber = lastRmaNumber + 1;
    
    const newRmaFolder = parentFolder + "/RMA " + newRmaNumber;
    
    // Get UTC date in yyyy-mm-dd format
    var d = new Date();
    var dateString = d.toISOString().split('T')[0];
    
    var newRow = [];
    
    newRow[0] = newRmaNumber;
    newRow[columnIndices['rma date']] = dateString;
    newRow[columnIndices['customer']] = customer;
    newRow[columnIndices['project description']] = description;
    newRow[columnIndices['customer rma folder']] = newRmaFolder;
    newRow[columnIndices['customer comm']] = techSupportTask;
    sheet.appendRow(newRow);
    
    return newRmaNumber;
  }
  catch (error)
  {
    Logger.log(error.message);
    return null;
  }
}



function testNewRow()
{
  Sheets_newRow("Acme", "Tactile Dynamite keeps exploding before RoadRunner reaches it.");
  Sheets_newRow("Beeblebrox", `
  Reduced sensitivity in 
  Heart Of Gold control panel.
  It keeps making tea.
  `);
}



/**
 * Get indices of selected RMA columns.
 *
 * @param spreadSheet  RMA sheet (not book).
 *
 * @return  Dictionary of named columns and their indices.
 */
function getColumnIndices(spreadSheet)
{
  var columnDict = 
      { 
        "rma date" : -1,
        "customer" : -1,
        "project description" : -1,
        "customer rma folder" : -1,
        "customer comm" : -1
      };
  const keys = Object.keys(columnDict);
  
  // Malcolm manually edited the sheet to create a range (named 
  // 'RmaHeadings') covering the top row (the column headings).
  var headingsRange = spreadSheet.getRangeByName('RmaHeadings');
  if (headingsRange == null)
  {
    Logger.log("No RmaHeadings range");
    return null;
  }
  
  Logger.log("Found " + headingsRange.getNumColumns() + " column headings.");
  
  var headings = headingsRange.getValues();
  
  var columnsFound = 0;
  
  for (let i = 0; i < headings[0].length; i++)
  {
    let oneHeading = headings[0][i];
    
    if (oneHeading.length == 0)
    {
      // Skip empty heading
      continue;
    }
    
    // Convert newlines to single space.
    oneHeading = oneHeading.replace(/\s+/g, " ");
    // Convert to standard form.
    oneHeading = oneHeading.trim().toLowerCase();
    
    for (const key of keys)
    {
      if (oneHeading === key)
      {
        columnDict[key] = i;
        columnsFound++;
        break;
      }
    }
    
    if (columnsFound == keys.length)
    {
      // Our work is done.
      break;
    }
  }

  for (const key of keys)
  {
    if (columnDict[key] < 0)
    {
      Logger.log("Error.  Missing column: " + key);
      return null;
    }
  }
  
  Logger.log(columnDict);
  
  return columnDict;
}



function testGetColumnIndices()
{
  rmaSheetId = "1hXnJmGlIYshJJjkCZetyVdGDPWNOB35lF7dOkgQgtZg";
  
  var ss = SpreadsheetApp.openById(rmaSheetId);
  var sheet = ss.getSheetByName("RMA");
  var columnIndices = getColumnIndices(ss);

  Logger.log(columnIndices);
}
