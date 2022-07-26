
/**
* @OnlyCurrentDoc
*
* The above comment directs Apps Script to limit the scope of file
* access for this add-on. It specifies that this add-on will only
* attempt to read or modify the files in which the add-on is used,
* and not all of the user's files. The authorization request message
* presented to users will reflect this limited scope.
*/

var DEBUG=false;
var DEBUGHEADING=false;

var VERSION="v2.0.8";
var NO_CAPTION_ALT_TEXT="NO_CAPTION";
var LIST_OF_TEXT="List Of ";
var TABLES_TEXT="Tables";
var IMAGES_TEXT="Images";
var NO_CAPTION_COLOR_BIT=1;
var colorLength="#RRGGBB".length; //WE'RE ABUSING BORDER COLOR TO ENABLE OR DISABLE CAPTIONING- i.e. IF BORDER COLOR ENDS ON "1" THEN CAPTIONING IS DISABLED!
//var ui = HtmlService.createHtmlOutputFromFile('captionmaker');

///END OF GLOBAL VARS ////
var activeCaptionCIList=[];
var inactiveCaptionBMList=[];
var inactiveCaptionCIList=[];
var updateBookListOld=[];
var updateBookListNew=[];
var currentDocHeadingNumbers=[];
var currentDocHeadingIndexes=[];
var currentDocHeadingTexts=[];

var lastRemovedBMID="";

function googleQuestion(title,message)
{
    var ui= DocumentApp.getUi();
    var result = ui.alert(
     title,
     message,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    // User clicked "Yes".
    ui.alert('Confirmation received.');
  } else {
    // User clicked "No" or X in the title bar.
    ui.alert('Permission denied.');
  }

}
function googleAlertWithTitle(title,message)
{
   var ui= DocumentApp.getUi();
    var result = ui.alert(
     title,
     message,
      ui.ButtonSet.OK);
}
function googleAlert(message)
{
  DocumentApp.getUi().alert(message);
}
function onOpen(e) {

         try
        {
         
          //console.log('onOpen '+ VERSION );
          DocumentApp.getUi().createAddonMenu()
          .addItem('Start', 'showSidebar')
          .addToUi();

        }
        catch(ex)
        {
          Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
          console.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
          DocumentApp.getUi().alert("Caption Maker was not able to add a 'Start' option to your 'Add-ons' / 'Caption Maker' menu item!\nTo resolve this issue, please click on 'Add-ons' / 'Manage Add-ons'" +
                                    "Once the list of all installed add-ons appears, click on the MANAGE button assigned to Caption Maker entry and select 'Use in this document' in order to allow it to function properly.");

        }


}


/**
* Opens a sidebar in the document containing the add-on's user interface.
* This method is only used by the regular add-on, and is never called by
* the mobile add-on version.
*/
function showSidebar() {
  var ui = HtmlService.createHtmlOutputFromFile('captionmaker');
  ui.setWidth(600);
  ui.setTitle('Caption Maker '+ VERSION );
  DocumentApp.getUi().showSidebar(ui);
  console.log('Caption Maker '+ VERSION );

}
function onInstall(e) { 
  onOpen(e);
}

function toggleDebug(debugcb)
{ 
   if(debugcb)
   {
    DEBUG=true;
    
   }
   else
   {
     DEBUG=false;
    

   }
  console.log('STATUS: Caption Maker '+ VERSION + 'DEBUG:' + DEBUG + ' debugcb: ' + debugcb); 
}




function format_label(newParagraph,labelStyle)
{
  newParagraph.setHeading(labelStyle[DocumentApp.Attribute.HEADING]);
  newParagraph.setAlignment(labelStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT]);
  newParagraph.editAsText().setItalic(labelStyle[DocumentApp.Attribute.ITALIC]);
  newParagraph.editAsText().setBold(labelStyle[DocumentApp.Attribute.BOLD]);
  newParagraph.editAsText().setFontSize(labelStyle[DocumentApp.Attribute.FONT_SIZE])
  newParagraph.editAsText().setForegroundColor(labelStyle[DocumentApp.Attribute.FOREGROUND_COLOR])

}

function check_starts_with(bigstring,smallstring)
{
  try{
    if(bigstring==null)
      return false;
    if (DEBUG)  console.log("SMALLSTRING LENGTH: "+ smallstring.length);
    if(bigstring.length<smallstring.length)
    {
      if (DEBUG)  console.log("BIGSTRING SMALLER THAN SMALLSTRING! #"+bigstring.substring(0,smallstring.length)+"# != #"+smallstring+"#");
      return false;
    }
    else
    {
      if(bigstring.substring(0,smallstring.length) == smallstring)
      {
        if (DEBUG)  console.log("STRING BEGINS WITH STRING! #"+bigstring.substring(0,smallstring.length)+"# == #"+smallstring+"#");
           
        return true;
      }
      else
      {
          if (DEBUG)  console.log("STRING DOESN'T BEGIN WITH STRING! #"+bigstring.substring(0,smallstring.length)+"# != #"+smallstring+"#");
        return false;
      }
    }
  }
  catch(ex)
  {
    console.log("CAUGHT EXCEPTION: #"+bigstring+"# == #"+smallstring+"# "+ ex);
    return false;
  }
}

function mark_single_element(addFlagToDisableCaption,elToProc)
{
  var ALT_TEXT=null;
  // var cache = getCache();
  var TABLE_BORDER_COLOR="0";
  if (addFlagToDisableCaption)
  {
    ALT_TEXT=NO_CAPTION_ALT_TEXT;
    TABLE_BORDER_COLOR=NO_CAPTION_COLOR_BIT;
  }
  
  var theTable=0;
  var elType=elToProc.getType();
  
  if (DEBUG)  console.log("marking el: " + elType + " as DISABLE: " + addFlagToDisableCaption );
  
  switch(elType)
      {
        case    DocumentApp.ElementType.PARAGRAPH:
          for(var chdidx=0;chdidx<elToProc.asParagraph().getNumChildren();chdidx++)
          {
            aChild=elToProc.asParagraph().getChild(chdidx);
            mark_single_element(addFlagToDisableCaption,aChild);
          }
          break;
        case    DocumentApp.ElementType.INLINE_IMAGE:
          var oldAltText=elToProc.asInlineImage().getAltDescription();
          
          if (addFlagToDisableCaption)
          {
            if (check_starts_with(oldAltText,NO_CAPTION_ALT_TEXT))
            {
              if (DEBUG)  console.log("NOTHING TO DO, ELEMENT ALREADY MARKED AS NO CAPTION: "+oldAltText);
              break; // nothing to do, element already marked as no caption!
            }
            else
            {
               if (DEBUG)  console.log("HAVE TO ADD NO CAPTION TO: "+oldAltText);
              elToProc.asInlineImage().setAltDescription(ALT_TEXT+ " " + oldAltText); // add the NO_CAPTION + space
            }
          }
          else
          {
            if (check_starts_with(oldAltText,NO_CAPTION_ALT_TEXT))
            {
              if (DEBUG)  console.log("HAVE TO REMOVE NO CAPTION FROM: "+oldAltText);
              oldAltText=oldAltText.substring(NO_CAPTION_ALT_TEXT.length+1); //remove NO_CAPTION and the space behind it...
              elToProc.asInlineImage().setAltDescription(oldAltText);
            }
            else
            {
              if (DEBUG)  console.log("NOTHING TO DO, ELEMENT ALREADY MARKED AS CAPTION: "+oldAltText);
              break; // nothing to do, element already marked as caption!
            }
          }
          if (DEBUG)  console.log(elToProc.asInlineImage().getAltDescription());
          break;
        case    DocumentApp.ElementType.INLINE_DRAWING:
         
          var oldAltText=elToProc.asInlineDrawing().getAltDescription();
          if (addFlagToDisableCaption)
          {
            if (check_starts_with(oldAltText,NO_CAPTION_ALT_TEXT))
            {
              if (DEBUG)  console.log("NOTHING TO DO, ELEMENT ALREADY MARKED AS NO CAPTION: "+oldAltText);
              break; // nothing to do, element already marked as no caption!
            }
            else
            {
              elToProc.asInlineDrawing().setAltDescription(ALT_TEXT+ " " + oldAltText);// add the NO_CAPTION + space
            }
          }
          else
          {
            if (check_starts_with(oldAltText,NO_CAPTION_ALT_TEXT))
            {
              oldAltText=oldAltText.substring(NO_CAPTION_ALT_TEXT.length+1);//remove NO_CAPTION and the space behind it...
              elToProc.asInlineDrawing().setAltDescription(oldAltText);
            }
            else
            {
              if (DEBUG)  console.log("NOTHING TO DO, ELEMENT ALREADY MARKED AS CAPTION: "+oldAltText);
              break; // nothing to do, element already marked as caption!
            }
          }
          
          
          if (DEBUG)  console.log(elToProc.asInlineDrawing().getAltDescription());
          break;

          //// THE NEXT THREE HAVE TO BE IN THIS ORDER ...
        case DocumentApp.ElementType.TABLE_CELL: //if it's a cell
        case DocumentApp.ElementType.TABLE_ROW:  //or a row
          theTable=elToProc.getParentTable();    // then get the parent table -> note, there is no BREAK, so it'll just continue to TABLE section...
        case DocumentApp.ElementType.TABLE:      // and then if it's a table
          if(theTable == 0)
          {
            theTable=elToProc;                   // use the table...
          }
          //Logger.log(JSON.stringify(theTable.asTable().getAttributes())+ " " + theTable.asTable().getBorderColor());
          borderColor=theTable.asTable().getBorderColor().toString().slice(0, -1) + TABLE_BORDER_COLOR;

          theTable.asTable().setBorderColor(borderColor);

          //Logger.log(JSON.stringify(theTable.asTable().getAttributes())+ " " + theTable.asTable().getBorderColor());
          break;

        default:
         if (DEBUG)  console.log("ELEMENT NOT MARKABLE:"+elToProc);
          break;

      }
}

function mark_element(addFlagToDisableCaption)
{
  // var cache = getCache();
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var selection=doc.getSelection();
  if (selection)
  {
    try
    {

      var selectedElements=selection.getRangeElements();
     if (DEBUG)  console.log("SELECTED ELEMENT COUNT: "+selectedElements.length)

      for (var i=0; i<selectedElements.length; i++)
      {
        var elToProc=null
        try
        {
          elToProc=selectedElements[i].getElement();
          mark_single_element(addFlagToDisableCaption,elToProc);

        }
        catch(ex)
        {
          Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack)
         if (DEBUG)  console.log("FAILED TO MARK ELEMENT:"+elToProc + " " + ex );
        }
      }

    }
    catch(ex2)
    {
      Logger.log("EXCEPTION CAUGHT: "+ex2 + " st: " +ex2.stack)
     if (DEBUG)  console.log("CAUGHT:" + ex2)
    }
  }
  else
  {
   if (DEBUG)  console.log("NO SELECTION MADE");
  }
  return Logger.getLog();
}



function removePreviousLabel(body,current_child_index,formatToUseForMatch)
{
  // var cache = getCache();
  var nextElement=body.getChild(current_child_index); // get element up or down from img/table

  var captionText=getCaptionTextOrNullIfNot(body,nextElement,formatToUseForMatch);
   if (DEBUG) console.log("caption text:" + captionText);
  if (DEBUG)  console.log("CHECK SHOULD REMOVE ON CI: " + current_child_index + "  USING FMT: " + formatToUseForMatch)
  if (captionText!=null)
  {
    var listInd = inactiveCaptionCIList.indexOf(current_child_index);
   if (DEBUG)  console.log(inactiveCaptionCIList);
   if (DEBUG)  console.log(listInd);
    var bmrk=0;
    if (listInd>-1)
    {
      bmrk=inactiveCaptionBMList[listInd].getId()
      lastRemovedBMID=bmrk; ///store it, so that we can replace the text in lists after insert new bookmark
    }
   if (DEBUG)  console.log("CI " + current_child_index + " BMidx: " + listInd + " BMR " +bmrk );
   if (DEBUG)  console.log("SHOULD REMOVE ON CI: " + current_child_index + "  TEXT FOUND: " + captionText)
    body.removeChild(nextElement);
  }
  return captionText;
}

function testCaptionCatcher()
{
  var someText="Figure ​6.5.9​-4. 1st UML call flow of Partial Import of Configuration Section";
   //var userProperties = PropertiesService.getDocumentProperties();
   var formatToUseForMatch="{reg}^(Figure|Image)\\s*([^a-z\\s]*)\\s*([^\\n]*)$";
   var regexString=formatToUseForMatch.replace('{reg}','');
   console.log("actregex:"+regexString);
  console.log("USING REGEX TO MATCH OLD LABEL:" + regexString);
  var properRegex = new RegExp(regexString,"i");
  var regFind=someText.match(properRegex)
  console.log("FOUND: " + regFind);
    console.log("FLen: " + regFind.length + " -> " + regFind[regFind.length-1]);

}

function getCaptionTextOrNullIfNot(body,nextElement,formatToUseForMatch)
{
   
  var reg=/(0|[1-9][0-9]*)/
  if( nextElement && nextElement.getType() == DocumentApp.ElementType.PARAGRAPH) // if it's not a paragraph, it's not a caption
  {

    var someText=nextElement.asText().getText().trim(); // get pure string from paragraph

    if (formatToUseForMatch.search('{reg}')==0)
    {
      //
      // example: {reg}Figure[.\-\s0-9]*([^\n]*)
      // or: {reg}Table[.\-\s0-9]*([^\n]*)
      // should match almost anything...
      if (DEBUG) console.log("someText: " + someText);
      var regexString=formatToUseForMatch.replace('{reg}','');
      if (DEBUG) console.log("USING REGEX TO MATCH OLD LABEL:" + regexString);
      var properRegex = new RegExp(regexString,"i");
      var regFind=someText.match(properRegex)
      if (DEBUG) console.log("FOUND: " + regFind);
      if (regFind)
      {
       if (DEBUG)  console.log("FLen: " + regFind.length + " _ " + regFind[regFind.length-1]);
        return regFind[regFind.length-1];
      }
        
    }

    if (DEBUG)  console.log("SOMETEXT:" + someText);
    var beginCust=formatToUseForMatch.search('{c}'); // locate where in caption definition the counter begins. e.g. if oldlabel="Image {c}. {t}" then we want the "Image " part to match against paragrap text

    if (beginCust>-1) // i.e. {c} was found, otherwise oldlabel is not formatted properly #NEED TO FIX IT - old label doesn't have to have {c} or {t} even.
    {
      var pureLabel=formatToUseForMatch.slice(0,beginCust); //this is our matching part of caption, i.e. "Image "
     if (DEBUG)  console.log("PURELABEL:" + pureLabel);
     if(someText.search(pureLabel)==0) // and the text needs to start with this label...
     {
       someText=someText.replace(pureLabel,""); //and now we replace it in the paragraph, so that if the paragraph had been such that it started with "Image 123 blalba" it would now start with the number, i.e. "123 blabla"
       if (DEBUG)  console.log("SOMETEXT2:" + someText);
       var position=someText.search(reg); // now we use our regex to see if the paragraph text now starts with number. (which represents {c} )
       if (position==0) // if it's not 0 then that means that this paragraph is not 100% formated like oldlabel suggests ... Maybe we could change it to >-1 instead of ==0 to autocorrect miniature errors in formatting...
       {
         var foundText=someText.match(reg)[0]; //we now extract that number, i.e. we get the old value for {c}
         if (DEBUG)  console.log("FOUND: @" + position + "->" + foundText +" len:" + foundText.length);
         var skipCharCnt=formatToUseForMatch.slice(beginCust+3,formatToUseForMatch.length).search('{t}'); // now we check how many characters are between end of {c} and beginning of {t} in the oldlabel
         var realText=someText.slice(skipCharCnt + foundText.length, someText.length).trim().replace(/^:/gm,'').trim().replace(/^\./gm,'').trim().replace(/^-/gm,'').trim(); //and we get {t} from paragraph by skipping those x characters that are between {c} and {t} in oldlabel, this allows for anything between {c} and {t} as long as the number of characters is correct
        if (DEBUG)  console.log("REALTEXT:" + realText);
         /// TODO: fix last image label -> EXCEPTION CAUGHT: Exception: Can't remove the last paragraph in a document section.
         // body.removeChild(nextElement);


         return realText; // and return the text we read for {t}
       }
     }
    }
    else
    {
       if (DEBUG) Logger.log("OLD LABEL DOESN'T CONTAIN NUMBERS:" + formatToUseForMatch);
      var beginCust=formatToUseForMatch.search('{t}');
      if (beginCust>-1)
      {
        var pureLabel=formatToUseForMatch.slice(0,beginCust); //this is our matching part of caption, i.e. "Image "
        if (DEBUG)  console.log("PURELABEL_NOCOUNT:" + pureLabel);
        if(someText.search(pureLabel)==0) // and the text needs to start with this label...
        {
          realText=someText.replace(pureLabel,""); //and now we replace it in the paragraph, so that if the paragraph had been such that it started with "Image 123 blalba" it would now start with the number, i.e. "123 blabla"

          return realText; // and return the text we read for {t}
        }
      }
      else
      {
        if (DEBUG) Logger.log("OLD LABEL DOESN'T CONTAIN NUMBERS NOR TEXT:" + formatToUseForMatch);

      }
    }

  }
  return null;

}


function get_element_child_index(element,body)
{
  var elType=element.getType();
  var parent=element.getParent()
  var parentType= parent.getType();

  try
  {
     switch (elType)
      {
        case DocumentApp.ElementType.TABLE:
          if (parentType == DocumentApp.ElementType.BODY_SECTION)  //captioning tables is possible only if they are in the body directly
          {
            return body.getChildIndex(element);
          }
          else
          {
            if (DEBUG) console.log("captioning requested on "+elType + " which isn't directly in BODY but " + parentType + " which is not supported.")
          }
          
          break;
        case DocumentApp.ElementType.INLINE_DRAWING:
        case DocumentApp.ElementType.INLINE_IMAGE:
          var parentsParentType=parent.getParent().getType();
          if (parentType == DocumentApp.ElementType.PARAGRAPH && parentsParentType == DocumentApp.ElementType.BODY_SECTION) //captioning images&drawings is possible only if they are in a paragraph which is in the body directly...
          {
            return body.getChildIndex(parent);
          }
          else
          {
            if (DEBUG) console.log("captioning requested on "+elType + " which is in a paragraph whose parent isn't BODY but " + parentsParentType + " which is not supported.")
          }
          break;
      }
  }
    catch (ex)
    {
     console.log("captioning requested caused exception, i.e. parentType: " + parentType + " elementType: " + element.getType());
     console.error('Caption Maker '+ VERSION +" EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }
  return -1;
}

function get_element_skip_flag(element)
{

  switch (element.getType())
  {
    case DocumentApp.ElementType.TABLE:

      /// AUTOMATICALLY DON'T ENUMERATE LIST OF tables, e.g. list of images...
      if(element.asTable().getNumRows() > 0 && element.asTable().getRow(0).getNumCells()>0)
      {
        var firstCell=element.getCell(0, 0);
        var fcText = firstCell.editAsText().getText().split("\n")[0];
        if(fcText.search(LIST_OF_TEXT)>-1)
          return true;
        /// IF IT's NOT A TABLE LIKE THAT, THEN USE COLOR TO DETERMINE... //TODO - > change to something better!
        if (element.asTable().getBorderColor().toString().slice(colorLength-1,colorLength) == NO_CAPTION_COLOR_BIT)
          return true;
        else
          return false;
        break;
      }
      else
      {
        console.log("PROBLEM WITH A TABLE THAT DOESN'T HAVE ROWS OR CELS...");
        DocumentApp.getUi().alert("Caption Maker ran into a situation it doesn't know how to handle: There is a table which either doesn't have any rows or any cels, please remove it or it will cause this error every time you ran Captionize.");
        return false;
        break;
      }
    case DocumentApp.ElementType.INLINE_DRAWING:
      if (check_starts_with(element.asInlineDrawing().getAltDescription(), NO_CAPTION_ALT_TEXT))
        return true;
      else
        return false;
      break;

    case DocumentApp.ElementType.INLINE_IMAGE:
      if (check_starts_with(element.asInlineImage().getAltDescription(), NO_CAPTION_ALT_TEXT))
        return true;
      else
        return false;
      break;

    default:
      return true; /// if we ask about any other type, it should be skipped...

  }

}


function captionize_element(doc,body,addCaption,objlist,label,labelStyle,ab_flag,start,oldlabel,use_heading_num,connector)
{
    //Logger.log("Objects:"+objlist);
  var prefixStr="";
  // var cache = getCache();
  try{
    prefixStr=label.split(" ")[0]; //crappy and buggy way of determining whether it's for Tables or Figures but it'll work most of the time and it's the simplest solution so whatever...
    if (DEBUG)  console.log("lost labels prefix:" + prefixStr);
  }
  catch(ex)
  {
    if (DEBUG)  console.log("caught ex when trying to get Figure or Table prefix for summary of lost labels:" + ex);
    prefixStr=label;
  }

  if ( oldlabel =="" ) oldlabel=label;
  var lostLabels=[];
  var counter=start;
  var oldHeadingNumber=null;
  var subHeadingCounter=1;
 if (DEBUG)  console.log("PROCESSING: " + objlist.length + " ELEMENTS " + " as " + label );
  for (var i = 0; i < objlist.length; i++)
  {
   if (DEBUG)  console.log("-----------------------------");
    var elToProc=objlist[i];
    var childIndex= -1;
    var textToKeep="";
    try
    {
      var shouldSkip=get_element_skip_flag(elToProc);
      childIndex=get_element_child_index(elToProc,body);
     if (DEBUG)  console.log("AC LIST:" + activeCaptionCIList)
     if (DEBUG)  console.log("EL:" + elToProc + " SKIP: " + shouldSkip + " CI: " + childIndex + " INLIST: " +activeCaptionCIList.indexOf(childIndex) );
      if (childIndex>0 && activeCaptionCIList.indexOf(childIndex-1)==-1) // if it's not the first element in document, look for old captioning above the element (except if it's one of the new captionings e.g. when two pics are one next to another)
      {
        var CIabove=childIndex - 1;
       if (DEBUG)  console.log("REMOVE PREVIOUS LABEL: " + childIndex + "  CIabove: " + CIabove)
        textToKeep=removePreviousLabel(body,CIabove,oldlabel)
       if (DEBUG)  console.log("TUP:" + textToKeep);
      }
      if (textToKeep!=null) //reindex if search returned something, even an empty string.
      {
        childIndex=get_element_child_index(elToProc,body);
       if (DEBUG)  console.log("new CI: " + childIndex);
      }

      if(childIndex<body.getNumChildren()-1) // if it's not the last child in the document...
      {



        if (textToKeep == null || textToKeep.length==0) //if the search above didn't return anything or returned an empty string, try to look below the element for caption
        {
          var textBelow=removePreviousLabel(body,childIndex+1,oldlabel);
          textToKeep=textBelow;
         if (DEBUG)  console.log("TDN:" + textToKeep);
        }
        else if(textBelow!=null && textBelow.length>0)
        {
         if (DEBUG)  console.log("TEXT disregarded: " + textBelow + " because of upper caption: " + textToKeep)
        }
      }
      else
       if (DEBUG)  console.log("LAST ELEMENT SO WON'T EVEN TRY TO REMOVE BELOW IT!")


      /// by now, we should have removed old captioning below or above (NOTE: we do not remove both unless the caption above was empty "".)

      if(shouldSkip || !addCaption)
      {
        if (textToKeep!=null && textToKeep!="")
          lostLabels.push(textToKeep);


      }
      else
      {
         if (addCaption)
        {
          if (textToKeep==null)
            textToKeep=""
          var lastIndex=body.getNumChildren()-1;
          var insertPosition=childIndex + (!ab_flag);

          var labelToInsert="";
          if(use_heading_num)
          {
            var headingNumber=get_header_number_for_chdIdx(childIndex);
            if(headingNumber!=oldHeadingNumber)
            {
              oldHeadingNumber=headingNumber;
              subHeadingCounter=1;
            }
            labelToInsert = label.replace("{c}",headingNumber+connector+subHeadingCounter).replace("{t}",textToKeep);
            subHeadingCounter++;
          }    
          else
            labelToInsert = label.replace("{c}",counter).replace("{t}",textToKeep);

         if (DEBUG)  console.log("WANT TO INSERT @"+insertPosition +" (MAX: " +lastIndex + ") LABEL:" +labelToInsert)
          if(insertPosition>lastIndex)
          {
            body.appendParagraph("\n");
          }
          var newParagraph=body.insertParagraph(insertPosition,labelToInsert);


          var position = doc.newPosition(newParagraph.getChild(0), 0);
          var bookmark = doc.addBookmark(position);
          var bmid=bookmark.getId();
          if (DEBUG)  console.log("OLD: " +lastRemovedBMID +"NEW BMID: " + bmid);

          if(lastRemovedBMID!="")
          {
            if(DEBUG) Logger.log("should replace " + lastRemovedBMID + " with " +bmid);

            //body.replaceText(new RegExp(lastRemovedBMID),bmid);
            updateBookListOld.push(lastRemovedBMID);
            updateBookListNew.push(bmid);
            lastRemovedBMID="";
            //Logger.log(body.replaceText("^"+url+lastRemovedBMID+"$", bmid));
          }

          activeCaptionCIList.push(insertPosition);

          format_label(newParagraph,labelStyle);

          counter++;

        }

      }
    }
    catch (ex)
    {
     Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
     console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }
  }
  if (lostLabels.length>0){
    var lostLabelText="";
    for (var i = 0; i < lostLabels.length; i++)
    {
      lostLabelText+="- " + lostLabels[i]+"\n";
    }
    DocumentApp.getUi().alert("You've marked some " + prefixStr + " elements as non caption or disabled captioning for such elements and this will cause the following captions to be lost:\n"+lostLabelText);

  }
  return counter - start;
}



function getPreferences() {
  // var cache = getCache();

 if (DEBUG)  console.log("GETTING PREFEREN");

  try
  {
    var userProperties = PropertiesService.getDocumentProperties();
    var docPref=userProperties.getProperties();
    if (docPref["version"]==VERSION)
      return docPref;

  }
  catch(ex)
  {
    Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
  }
  return;

}
function insert_images(above_flag,listOfBookmarks,listOfBTexts)
{
    // var cache = getCache();
  if (DEBUG)  console.log("insert_images ");
 if (DEBUG)  console.log("INSERT IMAGE LIST");
  try
  {
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var objList=[];
  var positionedList=[];
  populate_images_and_drawings(body,objList,positionedList);
  insert_list_of_items(objList,above_flag,LIST_OF_TEXT+IMAGES_TEXT,listOfBookmarks,listOfBTexts);
  }
  catch(ex)
    {
     Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }
    return Logger.getLog();

}
function insert_tables(above_flag,listOfBookmarks,listOfBTexts)
{
    if (DEBUG)  console.log("insert_tables ");
    // var cache = getCache();
  try
  {
    if (DEBUG)  console.log("INSERT TABLE LIST");
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var objList=body.getTables();
    insert_list_of_items(objList,above_flag, LIST_OF_TEXT+TABLES_TEXT,listOfBookmarks,listOfBTexts);
  }
  catch(ex)
  {
    Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
  }
  return Logger.getLog();
}

function get_all_lists_tables()
{
  // var cache = getCache();
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var tableList=body.getTables();
  var resultList=[];
   for (var i=0; i<tableList.length ;i++)
  {
    try
    {
      var ctable=tableList[i];
      if(ctable.asTable().getNumRows() > 0 && ctable.asTable().getRow(0).getNumCells()>0)
      {
        var firstCell=ctable.getCell(0, 0);
        var fcText = firstCell.editAsText().getText().split("\n")[0];
        if (DEBUG)  console.log(fcText);
        
        if(fcText.search(LIST_OF_TEXT)>-1)
        {
          var ofWhat=fcText.replace(LIST_OF_TEXT,"");
          resultList.push([ofWhat,ctable]);
        }
      }
      else
      {
         console.log("PROBLEM WITH A TABLE THAT DOESN'T HAVE ROWS OR CELS...");
         DocumentApp.getUi().alert("Caption Maker ran into a situation it doesn't know how to handle: There is a table which either doesn't have any rows or any cels, please remove it or it will cause this error every time you ran Captionize.");
      }
      
    }
     catch(ex)
    {
     Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }
  }
  return resultList;
}

function update_table_and_image_lists(img_ab_flag,tbl_ab_flag)
{
  if (DEBUG)  console.log("update_table_and_image_lists ");
    // var cache = getCache();
  try
  {

    if (DEBUG)  console.log("DEBUG: "+DEBUG);
    var doc = DocumentApp.getActiveDocument();
    var body = doc.getBody();
    var tableList=get_all_lists_tables();

    var resultset=getListOfBookmarksAndTexts(doc);
    var listOfBookmarks=resultset[0];
    var listOfBTexts=resultset[1];
    //var listOfCIndex=resultset[2];
    if (DEBUG)  console.log("UPDATE LIST ->: "+tableList.length);
    for (var i=0; i<tableList.length ;i++)
    {
      try
      {
        var ofWhat=tableList[i][0];
        var ctable=tableList[i][1];
        var chdIdx = body.getChildIndex(ctable);
        var pos= doc.newPosition(ctable.getChild(0), 0);
        doc.setCursor(pos)
        body.removeChild(ctable);
        switch(ofWhat)
        {
          case TABLES_TEXT:
            insert_tables(tbl_ab_flag,listOfBookmarks,listOfBTexts);
            break;
          case IMAGES_TEXT:
            insert_images(img_ab_flag,listOfBookmarks,listOfBTexts);
            break;
          default:
            break;
        }

      }
      catch(ex)
      {
        Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
        Logger.log("ERR: " +i+" -> " +ex);
        console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      }
    }
  }
    catch(ex)
    {
     Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }


  return Logger.getLog();
}

function getListOfBookmarksAndTexts(doc)
{
    // var cache = getCache();
  var body=doc.getBody();
  var listOfBTexts=[];
  var listOfCIndex=[];
  var listOfBookmarks=doc.getBookmarks();
  for (var i=0; i< listOfBookmarks.length;i++)
  {
    var bmrk=listOfBookmarks[i];

    var bpos=bmrk.getPosition();
    var bel=bpos.getElement();
    //console.log("BOOKMARK: " +i + " id: "+bmrk.getId() +" : " + bel.getType() + " : " + bel.asText().getText());
    try
    {
      if (bel.getType() == DocumentApp.ElementType.PARAGRAPH)
      {
        var btext=bel.asParagraph().getText();
        //if(btext.startswith("TODO"))
        if(true)
        {
          var cIndex=body.getChildIndex(bel);
          listOfBTexts.push(btext);
          listOfCIndex.push(cIndex);
          if (DEBUG)  console.log("valid BOOKMARK: " +i + " id: "+bmrk.getId() +" TEXT: " + btext+ " CI:" +cIndex);
        }
      }
      else
      {

          listOfBTexts.push("OTH" + i);
          listOfCIndex.push(0);
        console.log("other BOOKMARK: " +i + " id: "+bmrk.getId() +" Points to something else other than a paragraph! : " + bel.getType() + " " + bel.asText().getText());  
      }
    }
    catch(ex)
    {

      listOfBTexts.push("EXC" + i);
      listOfCIndex.push(0);
      Logger.log("problem BOOKMARK: " +i + " id: "+bmrk.getId() +" Points to something else other than a paragraph! : " + bel.getType());
      Logger.log("EX:" + ex + " ST: "+ex.stack);
      console.error("problem BOOKMARK: " +i + " id: "+bmrk.getId() +" Points to something else other than a paragraph! : " + bel.getType());
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
    }

  }
  if(DEBUG)
  {
    console.log("STORED BOOKMARKS:")
    for (var i=0; i< listOfBookmarks.length;i++)
    {
      var bmrk=listOfBookmarks[i];
      console.log(i + " : " + bmrk.getId() + "  " +listOfBTexts[i]);
    }
  }
  return [listOfBookmarks,listOfBTexts,listOfCIndex];
}

function insert_list_of_items(objList,above_flag,header,listOfBookmarks,listOfBTexts)
{
  if (DEBUG)  console.log("insert_list_of_items ");
    // var cache = getCache();
 if (DEBUG)  console.log("INSERT AB:" + above_flag)
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var cursor=doc.getCursor();
  var el=cursor.getElement();
  if (DEBUG) console.log("INSERT AT ELEMENT:" + el.getType() + " PARENT: " + el.getParent().getType())
  if (el.getParent().getType() != DocumentApp.ElementType.BODY_SECTION )
  {
    DocumentApp.getUi().alert("The location of the cursor has to be outside of any existing element in order to insert a list of images or a list of tables. E.g. you need to position the cursor between paragraphs rather than within a paragraph or an existing table, etc.");
    return;
  }
  var chdIds=body.getChildIndex(el);
  
  
  var TOC=body.insertTable(chdIds);
  var row=TOC.appendTableRow();
  var cell = row.appendTableCell(header);
  //mark_single_element(true,TOC);


  var formattedText="";




  var ind=0;

  var positionListTexts=[];

  if (listOfBookmarks==null || listOfBTexts==null)
  {
    var resultset=getListOfBookmarksAndTexts(doc);
    listOfBookmarks=resultset[0];
    listOfBTexts=resultset[1];
    // var listOfCIndex=resultset[2];
  }

  for (var i=0; i< objList.length;i++)
  {
    try
    {
      var offset=1;
      if(above_flag)
        offset=-1;
      var elementInQ=objList[i];

      var skipThis=get_element_skip_flag(elementInQ);
      var child_index=get_element_child_index(elementInQ,body);
      //    var skipThis=resultset[1];
      if (!skipThis)
      {
        
        var paraLabel=body.getChild(child_index+offset).asParagraph().getText();
        //var paraLabel=captionElement.asParagraph().getText();
       if (DEBUG)  console.log("ParaLabel: " +paraLabel)
        if (paraLabel.length)
        {
          var bookIndex=listOfBTexts.indexOf(paraLabel);
         if (DEBUG) 
         {
           console.log("bookIndex: "+bookIndex +" / " + listOfBTexts.length + " ID: " +listOfBookmarks[bookIndex].getId())
           console.log("srcLabel: "+ paraLabel + "   ====   found label:"+listOfBTexts[bookIndex]);
         }
         
          var paraLabel=listOfBTexts[bookIndex];
          //var aListItem=body.appendListItem(chdIds++, paraLabel);
          var aListItem=cell.appendListItem(paraLabel);
          //var url = DocumentApp.getActiveDocument().getUrl().replace("https://docs.google.com/open?id=","https://docs.google.com/document/d/");
          //aListItem.setLinkUrl(url +"/edit#bookmark="+listOfBookmarks[bookIndex].getId());
          aListItem.setLinkUrl("#bookmark="+listOfBookmarks[bookIndex].getId());
          
          aListItem.setGlyphType(DocumentApp.GlyphType.SQUARE_BULLET);



        }

      }
    }
    catch(ex)
    {
      Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
     if (DEBUG)
     {
       Logger.log("OBJ:" + objList[i]+" ERR:" +ex);
       console.error("OBJ:" + objList[i]+" ERR:" +ex);
     }
    }
  }


}
function get_document_properties()
{
  var allProps =JSON.stringify( PropertiesService.getDocumentProperties().getProperties()); 
  console.log(allProps)
  return allProps;
}

function captionize_document(labelStyleToUse,use_center,use_italic,use_bold,img_caption_on,tbl_caption_on,
                             img_label,tbl_label,img_ab_flag,tbl_ab_flag,
                             start_img,start_tbl,img_oldlabel,tbl_oldlabel,font_size,font_color,heading_num_img,heading_num_tbl,update_links,
                             connector_img,connector_tbl)
{
// var cache = getCache();
  if (DEBUG)  console.log("captionize_document ");
  try
  {
  Logger.clear();

    try
    {
      var userProperties = PropertiesService.getDocumentProperties();
      userProperties.setProperty('version', VERSION);
      userProperties.setProperty('labelStyleToUse', labelStyleToUse);
      userProperties.setProperty('use_center', use_center);
      userProperties.setProperty('use_italic', use_italic);
      userProperties.setProperty('use_bold', use_bold);
      userProperties.setProperty('img_caption_on', img_caption_on);
      userProperties.setProperty('tbl_caption_on', tbl_caption_on);
      userProperties.setProperty('img_label', img_label);
      userProperties.setProperty('tbl_label', tbl_label);
      userProperties.setProperty('img_ab_flag', img_ab_flag);
      userProperties.setProperty('tbl_ab_flag', tbl_ab_flag);
      userProperties.setProperty('start_img', start_img);
      userProperties.setProperty('start_tbl', start_tbl);
      userProperties.setProperty('img_oldlabel', img_oldlabel);
      userProperties.setProperty('tbl_oldlabel', tbl_oldlabel);
      userProperties.setProperty('font_size', font_size);
      userProperties.setProperty('font_color', font_color);
      userProperties.setProperty('img_bl_flag', !img_ab_flag);
      userProperties.setProperty('tbl_bl_flag', !tbl_ab_flag);
      userProperties.setProperty('heading_num_img',heading_num_img);
      userProperties.setProperty('heading_num_tbl',heading_num_tbl);
      userProperties.setProperty('update_links',update_links);
      userProperties.setProperty('connector_img',connector_img);
      userProperties.setProperty('connector_tbl',connector_tbl);

      //userProperties = PropertiesService.getUserProperties(); //DON'T USE THIS, USE DOCUMENT PROP
      if (DEBUG) Logger.log("SAVE PROP:"+JSON.stringify(userProperties.getProperties()));
    }
    catch (ex)
    {
      Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      Logger.log("COULDN'T STORE PROPERTIES: " +ex);

      Logger.log(ex.stack)
      try
      {
        console.log("getScriptProperties")
        var prop=null;
        prop = PropertiesService.getScriptProperties();
        prop.deleteAllProperties();
        console.log("getUserProperties")
        prop = PropertiesService.getUserProperties();
        prop.deleteAllProperties();
        console.log("getDocumentProperties")
        prop = PropertiesService.getDocumentProperties();
        prop.deleteAllProperties();


        Logger.log("DONE");
      }
      catch (ex2)
      {
        Logger.log("EXCEPTION CAUGHT: "+ex2 + " st: " +ex2.stack);
        console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      }

    }


  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  //DocumentApp.getUi().alert('This process might take a while, please be patient. \nAnother message will appear after completion!');
  if(heading_num_img || heading_num_tbl)
    populateDocHeadingsList();

  var resultset=getListOfBookmarksAndTexts(doc);
  inactiveCaptionBMList=resultset[0];
  //var listOfBTexts=resultset[1];
  inactiveCaptionCIList=resultset[2];
"https://www.googleapis.com/auth/userinfo.email"
  var labelStyle = {};

  if(labelStyleToUse==0){
     labelStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.NORMAL ;
  }
  else if (labelStyleToUse==1){
      labelStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.SUBTITLE ;
  }
  else if (labelStyleToUse==2){
      labelStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.HEADING5 ;
  }
  else if (labelStyleToUse==3){
      labelStyle[DocumentApp.Attribute.HEADING] = DocumentApp.ParagraphHeading.HEADING6;
  }

  labelStyle[DocumentApp.Attribute.ITALIC] = use_italic;
  labelStyle[DocumentApp.Attribute.BOLD] = use_bold;
  if (use_center)
    labelStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.CENTER;
  else
    labelStyle[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = DocumentApp.HorizontalAlignment.LEFT;
  labelStyle[DocumentApp.Attribute.FONT_SIZE]=font_size;
  labelStyle[DocumentApp.Attribute.FOREGROUND_COLOR]=font_color;




    //var objList= body.getImages();
    var objList=[];
    var positionedList=[];
    populate_images_and_drawings(body,objList,positionedList);
    var totalImg=objList.length;
    var capImg=captionize_element(doc,body,img_caption_on,objList,img_label,labelStyle,img_ab_flag,start_img,img_oldlabel,heading_num_img, connector_img);
  //captionize_images(body,img_label,labelStyle,skip_img,start_img,img_oldlabel);


    var objList= body.getTables();
    var totalTab = objList.length;
    var capTab = captionize_element(doc,body,tbl_caption_on,objList,tbl_label,labelStyle,tbl_ab_flag,start_tbl,tbl_oldlabel,heading_num_tbl,connector_tbl);
  //captionize_tables(body,tbl_label,labelStyle,skip_tbl,start_tbl,tbl_oldlabel);


  /////////////AUTOMATICALLY UPDATE ALL LISTS OF ... //////////////////////
  if(DEBUG==true) Logger.log("Old: " +JSON.stringify(updateBookListOld));
  if(DEBUG==true) Logger.log("New:" + JSON.stringify(updateBookListNew));
  var tablesToUpdate=get_all_lists_tables();
  for (var j=0;j<tablesToUpdate.length;j++)
  {
    var ctable=tablesToUpdate[j][1];
    if(ctable.asTable().getNumRows() > 0 && ctable.asTable().getRow(0).getNumCells()>0)
    {
      var cell=ctable.getCell(0,0);
      var cellCNum=cell.getNumChildren();
      for (var k=0;k<cellCNum;k++)
      {
        var listof=cell.getChild(k);
        var cLinkUrl=listof.getLinkUrl();
        if (cLinkUrl)
        {
          for (var i=0; i < updateBookListOld.length;i++)
          {
            var oldId=updateBookListOld[i];
            var newId=updateBookListNew[i];
            if(DEBUG==true) Logger.log("CHECK IF:"+  oldId + " is in URL: " + cLinkUrl);
            if (cLinkUrl.search(oldId) > -1)
            {
              cLinkUrl= cLinkUrl.replace(oldId,newId);
              listof.setLinkUrl(cLinkUrl);
              if(DEBUG==true) Logger.log("REPLACE:"+  listof.getText() + " URL: " + cLinkUrl);
            }
            
          }
        }
        
      }
    }
    else
    {
      console.log("PROBLEM WITH A TABLE THAT DOESN'T HAVE ROWS OR CELS...");
      DocumentApp.getUi().alert("Caption Maker ran into a situation it doesn't know how to handle: There is a table which either doesn't have any rows or any cels, please remove it or it will cause this error every time you ran Captionize.");
    }   

  }

  if(update_links)
    updateLinksToBookmarks();

  var additionalWarning="";
  updateBookListOld=[];
  updateBookListNew=[];
  if(positionedList.length)
  {
    
    additionalWarning="\n\n\n WARNING!!!\nCaptionize process found " + positionedList.length + " positioned images which aren't supported (i.e. 'Wrap Text' or 'Break Text').\n\nTo fix this, check this support article:\n\n https://support.google.com/docs/answer/97447?hl=en&ref_topic=9045752#zippy=%2Cposition-edit-an-image-in-a-document ";
    console.log(additionalWarning);
  }
  var messageToUser="Captioning done!\n->"+capImg+ " of "+ totalImg+" Inline Images and Drawings captioned!\n->"+capTab + " of " +totalTab + " Tables captioned!" + additionalWarning;



  DocumentApp.getUi().alert(messageToUser);
  }
  catch(gex)
  {
    Logger.log("EXCEPTION CAUGHT: "+gex + " st: " +gex.stack);
    console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+gex + " st: " +gex.stack);
  }
  return Logger.getLog();

}


function populate_images_and_drawings(element,aList, positionedList)
{
  //if (DEBUG)  console.log("populate_images_and_drawings "); //gets called too often!
// var cache = getCache();
  if (DEBUG)  console.log("PROCESSING: "+element.getType());
  switch (element.getType())
  {
    
    case DocumentApp.ElementType.EQUATION:
    case DocumentApp.ElementType.EQUATION_FUNCTION:
    case DocumentApp.ElementType.EQUATION_FUNCTION_ARGUMENT_SEPARATOR:
    case DocumentApp.ElementType.EQUATION_SYMBOL:
    case DocumentApp.ElementType.COMMENT_SECTION:
    case DocumentApp.ElementType.FOOTNOTE_SECTION:
    case DocumentApp.ElementType.HEADER_SECTION:
    case DocumentApp.ElementType.HORIZONTAL_RULE:
    case DocumentApp.ElementType.FOOTNOTE:
    case DocumentApp.ElementType.FOOTNOTE_SECTION:
    case DocumentApp.ElementType.LIST_ITEM:
    case DocumentApp.ElementType.UNSUPPORTED:
    case DocumentApp.ElementType.TEXT:
    case DocumentApp.ElementType.PAGE_BREAK:
    
      break;
    case DocumentApp.ElementType.INLINE_DRAWING:
    case DocumentApp.ElementType.INLINE_IMAGE:
     //if (DEBUG)  console.log("ADDING "+element)
      aList.push(element)
      return;
      break;
    default:
      try
      {
        if(element.getType()==DocumentApp.ElementType.PARAGRAPH)
        {
          var positionedImages=element.getPositionedImages();
          if(DEBUG) console.log("FOUND POSITIONED IMAGES: "+positionedImages.length);
          for (var i = 0; i < positionedImages.length; i++)
          {
            var positionedImage = positionedImages[i];
            positionedList.push(positionedImage);
          }
        } 
        var numOfChildren = element.getNumChildren();
        if(DEBUG) console.log("FOUND CHILDREN: "+numOfChildren);
        if (numOfChildren)
        {

          for (var i = 0; i < numOfChildren; i++)
          {
            var child = element.getChild(i);
            populate_images_and_drawings(child,aList,positionedList);
          }

        }
        

      }
      catch(ex)
      {
       Logger.log("EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
        console.error('Caption Maker '+ VERSION +"EXCEPTION CAUGHT: "+ex + " st: " +ex.stack);
      }
      break;


  }


}


function updateLinksToBookmarks(element)
{
  var allChildren=[];
  if (element)
  {
    for(var a=0; a<element.getNumChildren();a++)
    {
      allChildren.push(element.getChild(a));
    }
  }
  else
  {
    var doc= DocumentApp.getActiveDocument();
    allChildren=doc.getBody().getParagraphs();
  }
  for (var j=0;j<allChildren.length;j++)
  {
    var child=allChildren[j];
    var ctype=child.getType();
    if(DEBUG)
      console.log("CT_ " +j + " : " + child.getType());

    if([DocumentApp.ElementType.PARAGRAPH, DocumentApp.ElementType.LIST_ITEM, DocumentApp.ElementType.TABLE ].includes(ctype))
   {
      var pcCnt=child.getNumChildren(); 
      if (pcCnt >0)
      {
        updateLinksToBookmarks(child);
      }
      else
      {
        
        var linkUrl=child.getLinkUrl();
        if(linkUrl)
        {
          if(DEBUG)
            console.log("found in " + child.getText()+" link: "+linkUrl);
          fixLink(linkUrl);
        }
      }      
    }
    else if(ctype==DocumentApp.ElementType.TEXT)
    {
      var wholeText=child.getText()
      if(wholeText)
      {
        if(DEBUG)
          console.log("processing " + wholeText)
        var currentPos=0;
        var begin=-1;
        var end=-1;
        var wholeLink=null;
        var wordsInText=wholeText.split(" ");
        for(k=0;k<wordsInText.length;k++)
        {
          var link=child.getLinkUrl(currentPos)
          
          if(DEBUG)
            console.log("currentPos:" + currentPos + " word: " + wordsInText[k] + " link "+ link);
          
          if(link != wholeLink)
          {
            if(link)
            {
              begin=currentPos;
              wholeLink=link;
            }
            else
            { 
               
              end=currentPos-1;
              if(DEBUG)
                console.log("found in " + child.getText()+" beg:" + begin + " end:" + end + " link: "+wholeLink);
              fixLink(wholeLink,child,begin,end);
              begin=-1;
              end=-1;
              wholeLink=null;
            }
          }
          currentPos+=wordsInText[k].length;
        
        }
        if(begin>=0 && wholeLink)
        {
          if(end<begin)
            end=currentPos;
          if(DEBUG)
            console.log("found in " + child.getText()+" beg:" + begin + " end:" + end + " link: "+wholeLink);

          fixLink(wholeLink,child,begin,end);
        }
          
      }
     
    }
  
  
  }
}

function fixLink(linkUrl,elementToSet,startOffset,endOffset)
{
  if(linkUrl.startsWith("#bookmark"))
  {
    var sublink=linkUrl.substr(10);
    if(updateBookListOld.includes(sublink))
    {
      var idxOld=updateBookListOld.indexOf(sublink)
      if(DEBUG)
          console.log("replacing "+ linkUrl + " with "+updateBookListNew[idxOld]);

        if(startOffset && endOffset)
          elementToSet.setLinkUrl(startOffset,endOffset,"#bookmark="+updateBookListNew[idxOld])
        else
          elementToSet.setLinkUrl("#bookmark="+updateBookListNew[idxOld])
    }
    
    /*
    for (var i=0; i < updateBookListOld.length;i++)
    {
      if(DEBUG)
        console.log(linkUrl.substr(10) + " check against "+ updateBookListOld[i]);
      if(linkUrl.substr(10) === updateBookListOld[i])
      {
        if(DEBUG)
          console.log("replacing "+ linkUrl + " with "+updateBookListNew[i]);

        if(startOffset && endOffset)
          elementToSet.setLinkUrl(startOffset,endOffset,"#bookmark="+updateBookListNew[i])
        else
          elementToSet.setLinkUrl("#bookmark="+updateBookListNew[i])
      }
      
    }*/
  } 
}

function populateDocHeadingsList()
{
  currentDocHeadingNumbers=[];
  currentDocHeadingIndexes=[];
  currentDocHeadingTexts=[];
  if(DEBUG)
  {
    console.log("populateDocHeadingsList");
    console.log("BEGIN HLIST CNT_:"+currentDocHeadingNumbers.length);
  }
  var doc = DocumentApp.getActiveDocument();
  var body = doc.getBody();
  var allChildren=body.getParagraphs();
  for (var i=0;i<allChildren.length;i++)
  {
    var paragraph=allChildren[i];
    var heading=paragraph.getHeading();
    if(heading != DocumentApp.ParagraphHeading.NORMAL && heading != DocumentApp.ParagraphHeading.TITLE && 
       heading != DocumentApp.ParagraphHeading.SUBTITLE)
    {
        var hFull=paragraph.getText(); 
        if(hFull.match(/\d.*/g))
        {
          var eon = hFull.search(/\s/);
          if (eon > -1)
          {
            var hNum=hFull.substring(0,eon);
            var hText=hFull.substring(eon+1);
            currentDocHeadingNumbers.push(hNum);
            currentDocHeadingIndexes.push(body.getChildIndex(paragraph));
            if(DEBUG)
              console.log("MATCHES AND EXTRACTED:" + hNum + "|" + hText) ;
          }
          else{
            if(DEBUG)
              console.log("MATCHES BUT NOT EXTRACTED:", hFull) ;
          }
           
        }
        else
        {
          if(DEBUG)
            console.log("HEADING BUT DOESN'T MATCH", hFull) ;
        }
    }
      
  }
  if(DEBUG)
  {
    console.log("populateDocHeadingsList");
    console.log("END HLIST CNT_:"+currentDocHeadingNumbers.length);
    for(var j=0;j<currentDocHeadingIndexes.length;j++)
    {
      console.log("#" + j + " ch_idx: " + currentDocHeadingIndexes[j]+ " hnum: "+ currentDocHeadingNumbers[j]);
    }
  }
    
}
function get_header_number_for_chdIdx(childIndex)
{
  if(currentDocHeadingNumbers.length==0)
    return "";

  for(var i=1; i<currentDocHeadingNumbers.length; i++) /// we can't have heading numbers on images inserted before the first heading number...
  {
    var hidx=currentDocHeadingIndexes[i];
    if (hidx>childIndex)
    {
      if(i>0)
        return currentDocHeadingNumbers[i-1];
      else
        return "";
    }  
  }
  return currentDocHeadingNumbers[currentDocHeadingNumbers.length-1]; // if there isn't a heading whose child index in body is greater than the child idx of 
                                                                      // the captioned element, then simply return the last heading
}

