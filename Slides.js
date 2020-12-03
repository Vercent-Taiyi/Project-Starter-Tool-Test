/**
 * Generate a clean basic URL for a presentation.
 *
 * @param url  Messy URL with edit/copy/preview option
 *             and redirection to specific slide.  For example:
 *             https://docs.google.com/presentation/d/16xw16zl1LqvYa_er9zTEVIOxr28MUpCM3xfQIuyBqpk/edit#slide=id.g2a812b5e25_3_26
 *
 * @return  A generic Google Docs URL.  For example:
 *          https://docs.google.com/open?id=16xw16zl1LqvYa_er9zTEVIOxr28MUpCM3xfQIuyBqpk
 */
function Slides_getCleanUrl(url)
{
  const presentation = SlidesApp.openByUrl(url);
  return presentation.getUrl();
}


/**
 * Clone one file. 
 *
 * @originalId  Id (as string) of file to clone. 
 * @newName  Name of new file. 
 *
 * @return Handle of new file, or null.
 */
function cloneFile(originalId, newName)//test
{
  try
  {
    var originalFile = DriveApp.getFileById(originalId);
    return originalFile.makeCopy(newName);
  }
  catch (error)
  {
    Logger.log("cloneFile failed: " + error.message);
  }
  return null;
}



/**
 * Create an RMA presentation.
 *
 * @param rmaNumber  Pre-allocated number (as string).
 * @param rmaFolderPath  Path to pre-created dedicated Dropbox folder for RMA.
 * @param rmaFolderLink  Hyperlink for RMA folder.
 * @param techSupportLink  URL of parent Tech Support Asana task.
 * 
 * @return Id of new presentation, or null.
 */
function createRma(rmaNumber, rmaFolderPath, rmaFolderLink, techSupportLink)
{
  var now = new Date();
  const rmaIssueDate = now.toISOString().split('T')[0];
  
  // FIXME discover id from title "RMA Template"
  const templateId = "1CP_erYaa-H80cNvVTF2fPyTLUdXy_3tGHUXx0KUlmSM";
  var newId = null;
  
  try
  {
    const newName = "RMA " + rmaNumber + " " + rmaIssueDate;
    var newFile = cloneFile(templateId, newName);
    newId = newFile.getId();
    Logger.log("Created " + newName + ", id: " + newId);
    
    var newPres = SlidesApp.openById(newId);
    newPres.replaceAllText("{{rma_number}}", rmaNumber, true);
    newPres.replaceAllText("{{rma_folder}}", rmaFolderPath, true);
    newPres.replaceAllText("{{tech_support_link}}", techSupportLink, true);
    newPres.replaceAllText("{{rma_issue_date}}", rmaIssueDate, true);
    
    //var newSlides = rmaPres.getSlides();
    //var slide1Shapes = newSlides[1].getShapes();
    //for (const 
    newPres.saveAndClose();
  }
  catch (error)
  {
    Logger.log(error.message);
    if (newId != null)
    {
      DriveApp.getFileById(newId).setTrashed(true);
      newId = null;
    }
  }
  finally
  {
    return newId;
  }
}



function testCreateRma()
{
  var t = createRma(
    "670",
    "/+ Customer Projects Global/Z. Shipped Projects/DEF/DEKA 2016-08-31 - (East Penn Mfg) 1072 Glove/RMA 670",
    "http://mhannah.co.uk",
    "https://app.asana.com/0/1175836475924949/1184683782231764");
}



/**
 * Append RMA slides to given presentation.
 *
 * @param original  Deck to append slides to.
 * @param rmaNumber  Pre-allocated number (as string).
 * @param rmaFolderPath  Path to pre-created dedicated Dropbox folder for RMA.
 * @param rmaFolderLink  Hyperlink for RMA folder.
 * @param techSupportLink  URL of parent Tech Support Asana task.
 * 
 * @return true if RMA slides were successfully appended.
 */
function Slides_appendRmaInfo(original, rmaNumber, rmaFolderPath, rmaFolderLink, techSupportLink)
{
  var appended = false;
  
  // Create a temporary RMA presentation
  const rmaPresId = createRma(rmaNumber, rmaFolderPath, rmaFolderLink, techSupportLink);
  if (rmaPresId == null)
  {
    return false;
  }

  try
  {
    // Copy RMA slides into end of original deck.
    var originalPres = SlidesApp.openById(original);
    var rmaPres = SlidesApp.openById(rmaPresId);
    var rmaSlides = rmaPres.getSlides();
    
    for (const slide of rmaSlides)
    {
      originalPres.appendSlide(slide);
    }
    
    appended = true;
    
    originalPres.saveAndClose();
    rmaPres.saveAndClose();
  }
  catch (error)
  {
    Logger.log("Failed to append RMA slides:\n" + error.message);
  }

  try
  {
    // We no longer need the RMA deck.
    DriveApp.getFileById(rmaPresId).setTrashed(true);
  }
  catch (error)
  {
    Logger.log("Failed to trash temporary RMA slides:\n" + error.message);
  }
    
  return appended;
}



function testAppendRma()
{
  const sprayRmaDeck = "1JzXJ2H75VuNFxG05NE73sYx1WBfGkyNagssI723g0X8";
  const gloveRmaDeck = "1-jFboSTzBN5fEme2h-bcP8VJYRza4H72e-Oz8RH4f3w";
  const fingerTpsRmaDeck = "1zGiZFtEhCaUiqRln_yVYgA0YuBBmmUnIkpMf4YIf2qQ";
  
  const success = Slides_appendRmaInfo(
    gloveRmaDeck,
    "670",
    "/+ Customer Projects Global/Z. Shipped Projects/DEF/DEKA 2016-08-31 - (East Penn Mfg) 1072 Glove/RMA 670",
    "http://mhannah.co.uk",
    "https://app.asana.com/0/1175836475924949/1184683782231764");
  Logger.log(success);
}

  
  
// EXPERIMENTAL
// Could maybe use this to implement clickable links.
// Also see https://developers.google.com/apps-script/guides/slides/editing-styling#text_styling
function discoverSlideInfo()
{
  const id = "1LzNOASc4orflxbU9Fb_rqWbDvntgJFL_shVn-AQHRe4";
  //const id = "1StqWrp4CQPcnoCRdX9sma4TjrpwJNf-SHA-6TC6CuI4";

  var pres = SlidesApp.openById(id);
  var slides = pres.getSlides();
  for (const slide of slides)
  {
    var slideShapes = slide.getShapes();
    for (const shape of slideShapes)
    {
      Logger.log(shape.toString());
    }
    return;
  }
}
// END OF EXPERIMENTAL
