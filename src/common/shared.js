/* global self, window, global, crypto, Uint8Array, DOMParser, Office, Word, console */
const REDACT_NS = "http://schemas.microsoft.com/redaction/1.0";
const REDACTED_TITLE = "Redacted";
/**
 * This class handles the creation, encoding, decoding and export of 
 * the customerXmlPart XML
 */
var redactionPart = function() {

    var partId = "";
    var partHtml = "";
    var partOoxml = "";
    /**
     * Sets the values for the part
     * @param {string} html The html from the text range
     * @param {string} ooxml The ooxml from the text range
     */
    this.setValues = function(html, ooxml) {
      partId = uuidv4();
      console.log("partId in class: " + partId);
      partHtml = html;
      partOoxml = ooxml;
      return partId;
    };
    /**
     * Loads the values from the provided xml string
     * @param {string} xml The xml from the customXmlPart
     */
    this.loadValues = function(xml) {
      var xmlDoc = new DOMParser().parseFromString(xml,'text/xml');
      partOoxml = Base64.decode(xmlDoc.getElementsByTagName("ooxml")[0].textContent);
      partHtml = Base64.decode(xmlDoc.getElementsByTagName("html")[0].textContent);
      partId = xmlDoc.getElementsByTagName("ccid")[0].textContent;
    };
    /**
     * PROPERTY GET
     * Get the OOXML for the part
     */
    this.getOoxml= function (){
      return partOoxml;
    };
    /**
     * PROPERTY GET
     * Get the HTML for the part
     */
    this.getHtml = function() {
      return partHtml;
    };
    /**
     * PROPERTY GET
     * Gets the ID for the part
     */
    this.getId = function() {
      return partId;
    };
    /**
     * Exports the information as XML string and with
     * the HTML and OOXML base64 encoded
     */
    this.export = function() {
      var base64Html = Base64.encode(partHtml);
      var base64Ooxml = Base64.encode(partOoxml);
      return "<redaction xmlns='" + REDACT_NS + "'>" +
                "<ccid>" + partId + "</ccid>" +
                "<html>" + base64Html + "</html>" +
                "<ooxml>" + base64Ooxml + "</ooxml>" +
              "</redaction>";
    }
    /**
     * Generates a GUID
     */
    function uuidv4() {
      return ([1e7]+-1e3+-4e3+-8e3+-1e11).replace(/[018]/g, c =>
        (c ^ crypto.getRandomValues(new Uint8Array(1))[0] & 15 >> c / 4).toString(16)
      );
    }
  
    /**
     * Encodes/Decodes a string as Base64
     */
    // eslint-disable-next-line no-undef, no-useless-escape
    var Base64={_keyStr:"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=",encode:function(e){var t="";var n,r,i,s,o,u,a;var f=0;e=Base64._utf8_encode(e);while(f<e.length){n=e.charCodeAt(f++);r=e.charCodeAt(f++);i=e.charCodeAt(f++);s=n>>2;o=(n&3)<<4|r>>4;u=(r&15)<<2|i>>6;a=i&63;if(isNaN(r)){u=a=64}else if(isNaN(i)){a=64}t=t+this._keyStr.charAt(s)+this._keyStr.charAt(o)+this._keyStr.charAt(u)+this._keyStr.charAt(a)}return t},decode:function(e){var t="";var n,r,i;var s,o,u,a;var f=0;e=e.replace(/[^A-Za-z0-9\+\/\=]/g,"");while(f<e.length){s=this._keyStr.indexOf(e.charAt(f++));o=this._keyStr.indexOf(e.charAt(f++));u=this._keyStr.indexOf(e.charAt(f++));a=this._keyStr.indexOf(e.charAt(f++));n=s<<2|o>>4;r=(o&15)<<4|u>>2;i=(u&3)<<6|a;t=t+String.fromCharCode(n);if(u!=64){t=t+String.fromCharCode(r)}if(a!=64){t=t+String.fromCharCode(i)}}t=Base64._utf8_decode(t);return t},_utf8_encode:function(e){e=e.replace(/\r\n/g,"\n");var t="";for(var n=0;n<e.length;n++){var r=e.charCodeAt(n);if(r<128){t+=String.fromCharCode(r)}else if(r>127&&r<2048){t+=String.fromCharCode(r>>6|192);t+=String.fromCharCode(r&63|128)}else{t+=String.fromCharCode(r>>12|224);t+=String.fromCharCode(r>>6&63|128);t+=String.fromCharCode(r&63|128)}}return t},_utf8_decode:function(e){var t="";var n=0;var r=c1=c2=0;while(n<e.length){r=e.charCodeAt(n);if(r<128){t+=String.fromCharCode(r);n++}else if(r>191&&r<224){c2=e.charCodeAt(n+1);t+=String.fromCharCode((r&31)<<6|c2&63);n+=2}else{c2=e.charCodeAt(n+1);c3=e.charCodeAt(n+2);t+=String.fromCharCode((r&15)<<12|(c2&63)<<6|c3&63);n+=3}}return t}}
  };

  /**
 * Gets the CustomXMLPart for the redacted content control with the provided ID (cc.tag)
 * @param {string} id The ID id of the content control (tag) that will be searched for
 * @param {redactionPartSuccessCallback} successCallback Result with a successful find
 * @param {stringErrorCallback} failCallback String returned with the error message
 */
function getXmlPart(id, successCallback, failCallback) {
    Office.context.document.customXmlParts.getByNamespaceAsync(REDACT_NS, function(partsResult) {
      if(partsResult.status == Office.AsyncResultStatus.Failed) {
        console.log("error getting parts");
        failCallback("error: " + partsResult.error);
        return;
      }
      var parts = partsResult.value;
      if(parts.length == 0)
      {
        console.log("no parts found");
        failCallback("no parts");
        return;
      }
      for(var idx = 0; idx < parts.length; idx++) {
        var part = parts[idx];
        var ctx = { asyncContext: part };
        part.getXmlAsync(ctx, function(xmlResult) {
          if(xmlResult.status === Office.AsyncResultStatus.Failed) {
            console.log("unable to get part xml: " + xmlResult.error);
            failCallback("unable to get part xml: " + xmlResult.error);
            return;
          } else {
            var xml = xmlResult.value;
            var partData = new redactionPart();
            partData.loadValues(xml);
            console.log(partData.getId() + " == " + id);
            if(partData.getId() === id) {
              successCallback(partData, xmlResult.asyncContext);
              return; // stop looking
            }
          }
        });
      }
    });
  }

/**
 * MAIN REDACT FUNCTION
 * Takes the current selection in Word and performs these steps:
 * 1) gets the current HTML (for display), the current OOXML (for rehydration) and the text
 * 2) creates the redaction text string from the text length
 * 3) Adds a customXMLPart with a guid ID
 * 4) Create a content control with the same Guid ID
 * 5) Deleted the selected text and adds the redaction to the content control
 * NOTE: This is used by both the Ribbon buttons and the TaskPane
 * @param {Boolean} unsupported - is true is Word online
 * @param {redactionSuccessCallback} successCallback - callback on success
 * @param {stringErrorCallback} failResult - callback on fail / not used
 */
function redactProcess(unsupported, successCallback, failCallback) {
  Word.run(/** @param {Word.RequestContext} context */ async function (context) {
    try {
      var range = context.document.getSelection();
      var html = range.getHtml();
      var ooxml = range.getOoxml();
      range.load("text");
      range.contentControls.load();
      // -- SYNC --
      context.sync().then(function() {
        // validate that test is selected
        if(range.text.length == 0) {
          console.log("no text selected");
          failCallback("No text selected.");
          return; // not a valid selection
        }
        // validate that there are no content controls in the selection
        // this is because we cannot allow overlapping redactions
        if(range.contentControls.items.length > 0) {
          console.log("Cannot redact areas with content controls, or other redactions.");
          failCallback("Cannot redact areas with content controls, or other redacitons.");
          return;
        }
        // build the replacement string - redacted
        var replacementText = "";
        for(var i=0; i < range.text.length; i++) {
          if(range.text[i] == "\r") {
            replacementText += "\r";
          } else if(range.text[i] == "\n") {
            replacementText += "\n";
          } else {
            replacementText += "░"; // light
          }
        }
        /**@type {redactionPart} */
        var xmlPart = new redactionPart();
        var id = xmlPart.setValues(html.value, ooxml.value);
        console.log("xmlPart.id = " + id);
        var xml = xmlPart.export();
        var asyncData = { asyncContext: id };
        // NOTE: Word.context.document.customXmlParts does not work yet.
        // NOTE: No demo documentation on usages either.
        Office.context.document.customXmlParts.addAsync(xml, asyncData, function(asyncResult) {
          if(asyncResult.status == Office.AsyncResultStatus.Failed) {
            return;
          }
          // next, clear the range and insert the CC
          range.clear();
          var cc = range.insertContentControl();
          console.log("asyncResult.asyncContext.partId = " + asyncResult.asyncContext);
          cc.tag = asyncResult.asyncContext;
          cc.title = REDACTED_TITLE;
          cc.insertText(replacementText, 'Replace');
          if(unsupported == false) {
            cc.cannotDelete = true; // not supported in the browser
            cc.cannotEdit = true; //not supported in the browser
          }
          cc.appearance = Word.ContentControlAppearance.hidden;
          // load items
          cc.load();
          context.sync().then(function() {
            // now delete the customXmlPart
            successCallback();
          });
        });
      });
    } catch { 
      failCallback("Unable to redact selection.");
    }
  });
}

/**
 * Finalizes the document by prompting the user, asking if they are sure
 * and if they confirm, then deleting all the customXmlParts in the 
 * document and also removing the tag and title on all the content 
 * controls.
 * @param {boolean} unsupported
 * @param {emptySuccessCallback} successCallback 
 * @param {stringErrorCallback} failCallback 
 */
function finalizeDocument(unsupported, successCallback, failCallback) {
  /**@type {Office.Dialog} */
  var dialog;
  Office.context.ui.displayDialogAsync(window.location.origin + "/dialog.html", 
    { height: 18, width: 45 , displayInIframe: true}, function(asyncResult) {
      if (asyncResult.status != "failed") {
        dialog = asyncResult.value;
        /*Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)*/
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function(arg) {
          console.log("The arg: " + arg.message);
          if(arg.message === "true" || arg.message == true)
          {
            // if we are here, the user clicked YES
            Office.context.document.customXmlParts.getByNamespaceAsync(REDACT_NS, function(partsResult) {
              var parts = partsResult.value;
              if(parts.length == 0)
              {
                failCallback("no parts found");
                return;
              }
              for(var idx = 0; idx < parts.length; idx++) {
                parts[idx].deleteAsync({ asyncContext: { index: idx } }, function(deleteResult) {
                  console.log(" > " + deleteResult.asyncContext.index);
                  if(deleteResult.asyncContext.index >= parts.length - 1) {
                    console.log("final part processed");
                    // once all parts are deleted - we need to
                    // update the content controls...
                    updateCCs(unsupported, successCallback, failCallback);
                  }   
                });
              }
            });
            // also go thorugh the document and update the content controls
            // to have a darker shade - to show they are finalized
          } else{
            failCallback("[1] The finalize process was cancelled.")
          }
          dialog.close();
        });
        /*Events are sent by the platform in response to user actions or errors. For example, the dialog is closed via the 'x' button*/
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function() {
          failCallback("[2] The finalize process was cancelled.");
         });
      }
    });
}

/**
 * Gets all the content controls in the document
 * @param {Word.RequestContext} context
 * @param {contentControlsSuccessCallback} successCallback 
 * @param {stringErrorCallback} failCallback 
 */
function getAllContentControls(context, successCallback, failCallback) {
  try {
    var contentControls = context.document.contentControls;
    contentControls.load();
    return context.sync().then(function() {
      /**@type {Word.ContentControl[]} */
      var ccs = [];
      console.log("found " + contentControls.items.length + " controls");
      for(var idx = 0; idx < contentControls.items.length; idx++) {
        console.log("processing: " + idx);
        /**@type {Word.ContentControl} */
        var cc = contentControls.items[idx];
        cc.load();
        ccs.push(cc);
      }
      console.log("last sync");
      return context.sync().then(function() {
        console.log("all done");
        successCallback(ccs);
      });
    });
  } catch {
    console.log("ERROR");
    failCallback("unable to catalog content controls");
  }
}

/**
 * Updates all the REDACTION content controls parts
 * @param {boolean} unsupported 
 * @param {emptySuccessCallback} successCallback 
 * @param {stringErrorCallback} failCallback 
 */
function updateCCs(unsupported, successCallback, failCallback){
  console.log("starting update...");
  Word.run(function (context) {
    getAllContentControls(context, function(ccs) {
      console.log("got all content controls: " + ccs.length);
      try{
        for(var idx = 0; idx < ccs.length; idx++) {
          var cc = ccs[idx];
          console.log("item" + idx + ": " + cc.title);
          if(cc.title == REDACTED_TITLE) {
            var text = cc.text;
            if(unsupported == false) {
              cc.cannotDelete = false; // not supported in the browser
              cc.cannotEdit = false; //not supported in the browser
            }
            cc.clear();
            var replacementText = "";
            for(var i=0; i < text.length; i++) {
              if(text[i] == "\r") {
                replacementText += "\r";
              } else if(text[i] == "\n") {
                replacementText += "\n";
              } else {
                replacementText += "█"; // dark
              }
            } // end for resplacementText
            cc.insertText(replacementText,"Replace");
            if(unsupported == false) {
              cc.cannotDelete = true; // not supported in the browser
              cc.cannotEdit = true; //not supported in the browser
            }
          } // end for ccs}
        }
        // all done
        return context.sync().then(function () {
          successCallback();
        }); // commit all changes
      } catch { 
        failCallback("unable to update content controls");
      }
    }, function(error) { failCallback(error); });
  });
}

/**
 * MAIN UNREDACT FUNCTION
 * Takes the selected content control and rehydrates it. It does the following:
 * 1) Verifies the selected text is inside a contentcontrol
 * 2) Finds the content control related customXmlPart by tag
 * 3) Removes the content control and the customXmlPart
 * 4) Takes the Ooxml content from the customXmlPart and insert it
 * NOTE: This is used by both the Ribbon buttons and the TaskPane
 * @param {Office.AddinCommands.Event} event 
 * @param {boolean} unsupported - Word online?
 * @param {emptySuccessCallback} successCallback - callback on success
 * @param {stringErrorCallback} failCallback - callback on fail / not used
 */
function unredactProcess(unsupported, successCallback, failCallback) {
  Word.run(/** @param {Word.RequestContext} context */ async function (context) {
    try {
      // get the selection and any content controls in the selection
      var range = context.document.getSelection();
      range.parentContentControl.load();
      var contentControls = range.contentControls;
      range.load();
      contentControls.load();
      // -- SYNC --
      context.sync().then(function () {
        var cc = range.parentContentControl;
        // if we do not have any content controls, we stop here
        if((cc === undefined || cc === null) && contentControls.items.length > 0) {
          // we only care about the first content control
          cc = contentControls.items[0];
        }
        if(cc === undefined || cc === null) {
          failCallback("No content control in selection.");
          return;
        }
        range = cc.getRange();
        range.load();
        cc.load();
        // -- SYNC --
        context.sync().then(function() {
          // get the tag from the content control and then grab the
          // customxmlpart with the same ID
          getXmlPart(cc.tag, 
            /**
             *  @param {redactionPart} xmlPart
             *  @param {Office.CustomXmlPart} part
             */ 
            function(xmlPart, part) {
              part.deleteAsync(function(asyncResult) {
                if(asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                  if(unsupported == false) {
                    cc.cannotDelete = false; // not supported in the browser
                    cc.cannotEdit = false; //not supported in the browser
                  }
                  range.insertOoxml(xmlPart.getOoxml(), 'After');
                  cc.delete();
                  context.sync().then(function() {
                    successCallback();
                  });
                } else { 
                  if(failCallback != null) {
                    failCallback("The customXmlPart was not deleted successfully.");
                  }
                }
              });
            }, function(errorString) {
              failCallback(errorString);
            });
        });
      });
    } catch { 
      failCallback("Unable to identify redaction.");
    }
  });
}

/**************************************************************************************/
/* HELPERS                                                                            */
/**************************************************************************************/

/**
 * This is the sucess callback
 * @callback emptySuccessCallback
 * @returns {void}
 */
// eslint-disable-next-line no-unused-vars
var emptySuccessCallback = function() { };

/**
 * This is the sucess callback
 * @callback redactionPartSuccessCallback
 * @param {redactionPart} result an instance of a redaction CustomXMLPart class 
 * @param {Office.CustomXMLPart} part the CustomXmlPart part found
 * @returns {void}
 */
// eslint-disable-next-line no-unused-vars
var redactionPartSuccessCallback = function(result, part) { };

/**
 * This is the success callback for ContentControls
 * @callback contentControlsSuccessCallback
 * @param {Word.ContentContol[]} ccs a list of content controls
 * @returns {void}
 */
// eslint-disable-next-line no-unused-vars
var contentControlsSuccessCallback = function(ccs) { };

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
  g.getXmlPart = getXmlPart;
  g.redactionPart = redactionPart;
  g.redactProcess = redactProcess;
  g.unredactProcess = unredactProcess;
  g.finalizeDocument = finalizeDocument;
  g.REDACT_NS = REDACT_NS;
  g.REDACTED_TITLE = REDACTED_TITLE;
  /*************************************************/
  /*    REQUIRED BY WEB PACK - DO NOT DELETE       */
  /*************************************************/