/// <reference path="../common/shared.js" />
/* global global, self, window, Office, console, redactProcess, unredactProcess */
var unsupported = false; // to help debug and use in Word online

/**
 * When any FunctionFile file button is clicked, this method
 * gets called first...
 */
Office.initialize = function (reason) {
  console.log("here: " + reason);
  Office.onReady(function(info) {
    if(info.host === Office.HostType.Word && info.platform === Office.PlatformType.OfficeOnline){
      unsupported = true;
      // NOTE: In Word Online, content control locking is not allowed
      // NOTE: In Word online, the Word.customXmlParts collection is not available
      // NOTE: As of (beta 5-17-2020), Word.customXmlParts does not work correctly
    }
  });
};
/**
 * MANIFEST FUNCTION
 * Calls to the shared.js/redactProcess to redact the current selection
 * @param {Office.AddinCommands.Event} event 
 */
function redactSelection(event) {
  redactProcess(unsupported, function() { 
    event.completed();
  });
}

/**
 * MANIFEST FUNCTION
 * Call to the shared.js/unredactProcess to unredact the currently 
 * selected contentcontrol
 * @param {Office.AddinCommands.Event} event 
 */
function unredactSelection(event) {
  unredactProcess(unsupported, function() {
    event.completed();
  }, function(error){
    console.log(error);
  })
}

/**************************************************************************************/
/* HELPERS                                                                            */
/**************************************************************************************/

/**
 * This is the sucess callback
 * @callback redactionPartSuccessCallback
 * @param {shared.redactionPart} result an instance of a redaction CustomXMLPart class 
 * @returns {void}
 */
// eslint-disable-next-line no-unused-vars
var redactionPartSuccessCallback = function(result) { };

/**
 * This is the error callback
 * @callback stringErrorCallback
 * @param {string} error - the error message passed back to the caller
 * @returns {void}
 */
// eslint-disable-next-line no-unused-vars
var stringErrorCallback = function(error) { };

/*************************************************/
/*    REQUIRED BY WEB PACK - DO NOT DELETE       */
/*************************************************/
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();
g.redactSelection = redactSelection;
g.unredactSelection = unredactSelection;
/*************************************************/
/*    REQUIRED BY WEB PACK - DO NOT DELETE       */
/*************************************************/


/*************************************************/
/*   REMOVED CODE - DO NOT DELETE                */
/*************************************************/

// /**
//  * MANIFEST FUNCTION
//  * Takes the current selection in Word and performs these steps:
//  * 1) gets the current HTML (for display), the current OOXML (for rehydration) and the text
//  * 2) creates the redaction text string from the text length
//  * 3) Adds a customXMLPart with a guid ID
//  * 4) Create a content control with the same Guid ID
//  * 5) Deleted the selected text and adds the redaction to the content control
//  * @param {Office.AddinCommands.Event} event 
//  */
// function redactSelection(event) {
//   Word.run(/** @param {Word.RequestContext} context */ async function (context) {
//     var range = context.document.getSelection();
//     var html = range.getHtml();
//     var ooxml = range.getOoxml();
//     range.load("text");
//     range.contentControls.load();
//     // -- SYNC --
//     context.sync().then(function() {
//       // validate that test is selected
//       if(range.text.length == 0) {
//         console.log("no text selected");
//         return; // not a valid selection
//       }
//       // validate that there are no content controls in the selection
//       // this is because we cannot allow overlapping redactions
//       if(range.contentControls.items.length > 0) {
//         console.log("cannot redact areas with content controls");
//         return;
//       }
//       // build the replacement string - redacted
//       var replacementText = "";
//       for(var i=0; i < range.text.length; i++) {
//         replacementText += "█";
//       }
//       var cleanHtml = Base64.encode(html.value);
//       var cleanOoxml = Base64.encode(ooxml.value);
//       var ccId = uuidv4();
//       var xmlPart = "<redaction xmlns='" + REDACT_NS + "'>" +
//                       "<ccid>" + ccId + "</ccid>" +
//                       "<html>" + cleanHtml + "</html>" +
//                       "<ooxml>" + cleanOoxml + "</ooxml>" +
//                     "</redaction>";
//       /** @type {Word.CustomXmlPart} */
//       var part = null;
//       if(unsupported == false) {
//         part = context.document.customXmlParts.add(xmlPart); // not supported in the browser (PREVIEW)
//       }
//       // next, clear the range and insert the CC
//       range.clear();
//       var cc = range.insertContentControl();
//       cc.tag = ccId; 
//       cc.insertText(replacementText, 'Replace');
//       if(unsupported == false) {
//         cc.cannotDelete = true; // not supported in the browser
//         cc.cannotEdit = true; //not supported in the browser
//       }
//       cc.appearance = Word.ContentControlAppearance.hidden;
//       // load items
//       cc.load();
//       if(unsupported == false) {
//         part.load("id");
//       }
//       // -- SYNC --
//       context.sync().then(function() {
//         // done
//         console.log("completed!");
//         console.log("PART:" + xmlPart);
//         if(unsupported == false) {
//           console.log("PARTID:" + part.id);
//           cc.tag = part.id; // update the tag of the CC to the part ID
//         } else {
//           Office.context.document.settings.set(cc.tag, cleanOoxml);
//           Office.context.document.settings.saveAsync();
//         }
//         context.sync();
//         event.completed();
//       });
//     });
//   });
// }

// /**
//  * MANIFEST FUNCTION
//  * Takes the selected content control and rehydrates it. It does the following:
//  * 1) Verifies the selected text is inside a contentcontrol
//  * 2) Finds the content control related customXmlPart by tag
//  * 3) Removes the content control and the customXmlPart
//  * 3) Takes the Ooxml content from the customXmlPart and insert it
//  * @param {Office.AddinCommands.Event} event 
//  */
// function unredactSelection(event) {
//   Word.run(/** @param {Word.RequestContext} context */ async function (context) {
//     // get the selection and any content controls in the selection
//     var range = context.document.getSelection();
//     var contentControls = range.contentControls;
//     range.load();
//     contentControls.load();
//     // -- SYNC --
//     context.sync().then(function () {
//       // if we do not have any content controls, we stop here
//       if(contentControls.items.length > 0) {
//         // we only care about the first content control
//         var cc = contentControls.items[0];
//         cc.load();
//         // -- SYNC --
//         context.sync().then(function() {
//           // get the tag from the content control and then grab the
//           // customxmlpart with the same ID
//           var id = cc.tag;
//           console.log("ID:" + id);
//           /** @type {OfficeExtension.ClientResult} */
//           var xml = null;
//           /** @type {Word.CustomXmlPart} */
//           //var part = null;
//           Office.context.document.customXmlParts.getByIdAsync(id, function(asyncResult) {
//             var xml = asyncResult.value;
//           });
//           // if(unsupported == false) {
//           //   part = context.document.customXmlParts.getItemOrNullObject(id);
//           //   xml = part.getXml();
//           //   part.load();
//           // }
//           // -- SYNC --
//           context.sync().then(function() {
//             var f = xml.value === null || xml.value === undefined;
//             range.insertText("here: " + f, 'End');
//             context.sync();
//             event.completed();
//             // if(unsupported == true || part != null)
//             // {
//             //   var ooxml = "";
//             //   // load the xml, grab the ooxml and decode it
//             //   if(unsupported == false) {
//             //     var xmlDoc = new DOMParser().parseFromString(xml,'text/xml');
//             //     ooxml = Base64.decode(xmlDoc.getElementsByTagName("ooxml")[0].textContent);
//             //   } else {
//             //     ooxml = Base64.decode(Office.context.document.settings.get(id));
//             //   }
//             //   console.log("OOXML:" + ooxml);
//             //   // now delete the content control, the customXMLPart
//             //   // and finally insert the OOXML
//             //   if(unsupported == false) {
//             //     cc.cannotDelete = false; // not supported in the browser
//             //     cc.cannotEdit = false; //not supported in the browser
//             //   }
//             //   range = cc.getRange();
//             //   cc.delete();
//             //   if(unsupported == false) {
//             //     part.delete();
//             //   }
//             //   range.insertOoxml(ooxml, 'Replace');
//             //   // -- SYNC --
//             //   context.sync().then(function() {
//             //     console.log("all done!");
//             //     event.completed();
//             //   });
//             // } else {
//             //   console.log("part not found");
//             //   event.completed();
//             // }
//           });
//         });
//       }
//       else
//       {
//         console.log("No content control in selection.");
//         event.completed();
//       }
//     });
//   });
// }

// function getXmlPart(id, successCallback, failCallback) {
//   Office.context.document.customXmlParts.getByIdAsync(id, function(partResult) {
//     if(partResult.status == Office.AsyncResultStatus.Failed) {
//       failCallback("Cannot get part - error: " + partResult.error);
//       return;
//     }
//     var part = partResult.value;
//     part.getXmlAsync(function(xmlResult) {
//       if(xmlResult.status == Office.AsyncResultStatus.Failed) {
//         failCallback("Cannot get part xml - error: " + xmlResult.error);
//       }
//       var xml = xmlResult.value;
//       var partData = new redactionPart();
//       partData.loadValues(xml);
//       successCallback(partData);
//     });
//   });
// }

// /**
//  * MANIFEST FUNCTION
//  * Takes the current selection in Word and performs these steps:
//  * 1) gets the current HTML (for display), the current OOXML (for rehydration) and the text
//  * 2) creates the redaction text string from the text length
//  * 3) Adds a customXMLPart with a guid ID
//  * 4) Create a content control with the same Guid ID
//  * 5) Deleted the selected text and adds the redaction to the content control
//  * @param {Office.AddinCommands.Event} event 
//  */
// function redactSelection(event) {
//   Word.run(/** @param {Word.RequestContext} context */ async function (context) {
//     var range = context.document.getSelection();
//     var html = range.getHtml();
//     var ooxml = range.getOoxml();
//     range.load("text");
//     range.contentControls.load();
//     // -- SYNC --
//     context.sync().then(function() {
//       // validate that test is selected
//       if(range.text.length == 0) {
//         console.log("no text selected");
//         return; // not a valid selection
//       }
//       // validate that there are no content controls in the selection
//       // this is because we cannot allow overlapping redactions
//       if(range.contentControls.items.length > 0) {
//         console.log("cannot redact areas with content controls");
//         return;
//       }
//       // build the replacement string - redacted
//       var replacementText = "";
//       for(var i=0; i < range.text.length; i++) {
//         if(range.text[i] == "\r") {
//           replacementText += "\r";
//         } else if(range.text[i] == "\n") {
//           replacementText += "\n";
//         } else {
//           replacementText += "█";
//         }
//       }
//       /**@type {redactionPart} */
//       var xmlPart = new redactionPart();
//       xmlPart.setValues(html.value, ooxml.value);
//       var xml = xmlPart.export();
//       Office.context.document.customXmlParts.addAsync(xml, function(asyncResult) {
//         if(asyncResult.status == Office.AsyncResultStatus.Failed) {
//           range.insertText("error creating part: " + asyncResult.error, 'End');
//         }
//         // next, clear the range and insert the CC
//         range.clear();
//         var cc = range.insertContentControl();
//         cc.tag = xmlPart.getId();
//         cc.insertText(replacementText, 'Replace');
//         if(unsupported == false) {
//           cc.cannotDelete = true; // not supported in the browser
//           cc.cannotEdit = true; //not supported in the browser
//         }
//         cc.appearance = Word.ContentControlAppearance.hidden;
//         // load items
//         cc.load();
//         context.sync().then(function() {
//           event.completed();
//         });
//       });
//     });
//   });
// }

// /**
//  * MANIFEST FUNCTION
//  * Takes the selected content control and rehydrates it. It does the following:
//  * 1) Verifies the selected text is inside a contentcontrol
//  * 2) Finds the content control related customXmlPart by tag
//  * 3) Removes the content control and the customXmlPart
//  * 3) Takes the Ooxml content from the customXmlPart and insert it
//  * @param {Office.AddinCommands.Event} event 
//  */
// function unredactSelection(event) {
//   Word.run(/** @param {Word.RequestContext} context */ async function (context) {
//     // get the selection and any content controls in the selection
//     var range = context.document.getSelection();
//     range.parentContentControl.load();
//     var contentControls = range.contentControls;
//     range.load();
//     contentControls.load();
//     // -- SYNC --
//     context.sync().then(function () {
//       var cc = range.parentContentControl;
//       // if we do not have any content controls, we stop here
//       if((cc === undefined || cc === null) && contentControls.items.length > 0) {
//         // we only care about the first content control
//         cc = contentControls.items[0];
//       }
//       range = cc.getRange();
//       range.load();
//       cc.load();
//       // -- SYNC --
//       context.sync().then(function() {
//         // get the tag from the content control and then grab the
//         // customxmlpart with the same ID
//         getXmlPart(cc.tag, /** @param {redactionPart} xmlPart */ function(xmlPart) {
//           if(unsupported == false) {
//               cc.cannotDelete = false; // not supported in the browser
//               cc.cannotEdit = false; //not supported in the browser
//           }
//           range.insertOoxml(xmlPart.getOoxml(), 'After');
//           cc.delete();
//           context.sync();
//           event.completed();
//         }, function(errorString) {
//           //range.insertText(errorString, 'End');
//           console.log(errorString);
//           context.sync();
//           event.completed();
//         });
//       });
//     });
//   });
// }