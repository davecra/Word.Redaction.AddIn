/// <reference path="../common/shared.js" />
/* global document, Office, Word, getXmlPart, redactProcess, unredactProcess, 
   finalizeDocument, window, REDACT_NS, console, REDACTED_TITLE */

var unsupported = false; // to help debug and use in Word online
var lineCount = 1; // for help with the log

/**
 * When the Taskpane is loaded, this code is executed
 */
Office.initialize = function (reason) {
  log(reason);
  Office.onReady(function(info) {
    log("Started...");
    if(info.host === Office.HostType.Word && info.platform === Office.PlatformType.OfficeOnline){
      unsupported = true;
      // NOTE: In Word Online, content control locking is not allowed
      // NOTE: In Word online, the Word.customXmlParts collection is not available
      // NOTE: As of (beta 5-17-2020), Word.customXmlParts does not work correctly
    }
    if (info.host === Office.HostType.Word) {
        if(unsupported == false) {
          document.getElementById("warningRow").hidden = "hidden";
        }
        document.getElementById("header").hidden = "hidden";
        document.getElementById("warningRow").hidden = "hidden";
        document.getElementById("hideHeader").hidden = "hidden";
        document.getElementById("logArea").hidden = "hidden";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("revealInfo").hidden = "hidden";
        document.getElementById("reveal").onclick = reveal;
        document.getElementById("clear").onclick = function() {
          document.getElementById("redaction").innerHTML = "";
          document.getElementById("revealInfo").hidden = "hidden";
        }
        // document.getElementById("hideHeader").onclick = function() {
        //   // header, warningRow, hideHeader, note1, note2, legend

        // };
        document.getElementById("redact").onclick = function() {
          redactProcess(unsupported, function() {
            log("The selection was successfully reddacted.");
            updateTaskPane();
          }, function(error) {
            displayNotification(error);
            log(error);
          })
        };
        document.getElementById("unredact").onclick = function () {
          unredactProcess(unsupported, function() {
            log("The selected redaction was unredacted successfully.");
            updateTaskPane();
          }, function(error) {
            displayNotification(error);
            log(error);
          });
        }
        document.getElementById("finalize").onclick = function() {
          finalizeDocument(unsupported, function() {
            log("The document has been finalized.");
            displayNotification("The document has been finalized.");
            updateTaskPane();
          }, function(error) {
            displayNotification(error);
            log(error);
          });
        }
        document.getElementById("notification").hidden = "hidden";
        // Now get the title of the document and the total number of redactions on the document
        window.setInterval(function() {
          updateTaskPane();
        }, 1000);
    }
  });
}

/**
 * Displays a yellow message at the top of the taskpane
 * Disappears after 5 seconds
 * @param {string} message The message to display
 */
function displayNotification(message) {
  document.getElementById("notificationText").textContent = message;
  var ele = document.getElementById("notification");
  ele.hidden = "";
  ele.style.top = "-20px";
  slideIt("notification", 0);
  window.setTimeout(function() {
    document.getElementById("notification").hidden = "hidden";
  }, 5000);
}

/**
 * Slides the given element into position slowly
 * @param {string} id id of the element
 * @param {number} max the final position 
 */
function slideIt(id, max){
  var ele = document.getElementById(id);
  var pos = parseInt(ele.style.top) 
  if(pos >= max)
  {
    return;
  }
  ele.style.top = (pos + 1) + "px";
  window.setTimeout(function() { slideIt(id, max); }, 15);
}

/**
 * Logs to the log textarea on the taskpane
 * @param {string} info 
 */
function log(info) {
  var data = document.getElementById("log").textContent;
  data = "(" + lineCount + ") " + info + "\n" + data;
  document.getElementById("log").textContent = data;
  console.log(info);
  lineCount++;
}

/**
 * Updates the information at the top of the taskpane with the
 * total number of redactions and the document name
 */
function updateTaskPane() {
  /**@type {Office.Document} */
  var doc = Office.context.document;
  var pos = doc.url.lastIndexOf("/");
  if(pos <= 0) {
    pos = doc.url.lastIndexOf("\\");
  }
  var name = doc.url.substr(pos + 1);
  document.getElementById("documentName").textContent = name;
  Office.context.document.customXmlParts.getByNamespaceAsync(REDACT_NS, function(asyncResult) {
    document.getElementById("redactionCount").textContent = asyncResult.value.length;
    enableDisableButton("finalize", asyncResult.value.length != 0);
    enableDisableButton("reveal", asyncResult.value.length != 0);
    enableDisableButton("unredact", asyncResult.value.length != 0);
  });
}

function enableDisableButton(id, enable) {
  var button = document.getElementById(id);
  if(enable) {
    button.disabled = false;
    button.style.boder = "";
    button.style.backgroundColor = "";
    button.style.color = "";
    button.className = "ms-Button ms-Button--hero ms-font-m";
  } else {
    button.disabled = true;
    //border: 1px solid #999999;
    //background-color: #cccccc;
    //color: #666666;
    button.className = "";
    button.style.boder = "1px solid #999999";
    button.style.backgroundColor = "#cccccc";
    button.style.color = "#666666";
  }
}

/**
 * Reveals the selected redaction in the taskpane with formatting
 * from the stored HTML in the customXmlPart
 */
function reveal() {
  log("Reveal started...");
  Word.run(/** @param {Word.RequestContext} context */ async function (context) {
    // get the selection and any content controls in the selection
    var range = context.document.getSelection();
    log("Checking for content controls in selection.");
    var contentControls = range.contentControls;
    //range.parentContentControl.load();
    range.load();
    contentControls.load();
    // -- SYNC --
    context.sync().then(function () {
      if(contentControls.isNullObject) {
        displayNotification("[1] No redactions found in current selection.");
        log("no content controls.");
        return;
      }
      /**@type {Word.ContentControl} */
      var cc = range.parentContentControl;  
      if((cc === undefined || cc === null) && contentControls.items.length > 0) {
        // we only care about the first content control
        cc = contentControls.items[0];
      }
      if(contentControls.items.length == 0 && (cc === undefined || cc === null)) {
        log("No content control in selection.");
        displayNotification("[2] No redactions found in current selection.");
        return;
      } else {
        log("Content control found.");
      }
      log("Loading content control information.");
      range = cc.getRange();
      range.load();
      cc.load();
      // -- SYNC --
      context.sync().then(function() {
        if(cc.isNullObject == false) {
          if(cc.title === REDACTED_TITLE) {
            // get the tag from the content control and then grab the
            // customxmlpart with the same ID
            log("Reading associated data part.");
            // eslint-disable-next-line no-unused-vars
            getXmlPart(cc.tag, /** @param {redactionPart} xmlPart */ function(xmlPart, _) {
              document.getElementById("redaction").innerHTML = xmlPart.getHtml();
              document.getElementById("revealInfo").hidden = "";
              log("Completed!");
            }, function(errorString) {
              displayNotification("[3] No redactions found in current selection: " + errorString + ".");
              log("Error: " + errorString);
            });
          } else {
            displayNotification("Selected item is either fully redacted or not a valid redaction.");
          }
        } else {
          log("No content control in selection.");
          displayNotification("[4] No redactions found in current selection.");
          return;
        }
      }, function(reason) {
        if(reason.message.indexOf("NotFound") >= 0) {
          log("No content control in selection.");
          displayNotification("[5] No redactions found in current selection.");
          return;
        }
      });
    }, function(reason) {
      if(reason.message.indexOf("NotFound") >= 0) {
        displayNotification("[6] No redactions found in current selection.");
        log("No content control was found in selection.");
      }
    }).catch(function(error) {
      log(error);
      displayNotification("[7] No redactions found in current selection: " + error);
    });
  });
}