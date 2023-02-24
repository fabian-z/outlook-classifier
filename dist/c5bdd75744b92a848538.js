function _createForOfIteratorHelper(o, allowArrayLike) { var it = typeof Symbol !== "undefined" && o[Symbol.iterator] || o["@@iterator"]; if (!it) { if (Array.isArray(o) || (it = _unsupportedIterableToArray(o)) || allowArrayLike && o && typeof o.length === "number") { if (it) o = it; var i = 0; var F = function F() {}; return { s: F, n: function n() { if (i >= o.length) return { done: true }; return { done: false, value: o[i++] }; }, e: function e(_e) { throw _e; }, f: F }; } throw new TypeError("Invalid attempt to iterate non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method."); } var normalCompletion = true, didErr = false, err; return { s: function s() { it = it.call(o); }, n: function n() { var step = it.next(); normalCompletion = step.done; return step; }, e: function e(_e2) { didErr = true; err = _e2; }, f: function f() { try { if (!normalCompletion && it.return != null) it.return(); } finally { if (didErr) throw err; } } }; }
function _unsupportedIterableToArray(o, minLen) { if (!o) return; if (typeof o === "string") return _arrayLikeToArray(o, minLen); var n = Object.prototype.toString.call(o).slice(8, -1); if (n === "Object" && o.constructor) n = o.constructor.name; if (n === "Map" || n === "Set") return Array.from(o); if (n === "Arguments" || /^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)) return _arrayLikeToArray(o, minLen); }
function _arrayLikeToArray(arr, len) { if (len == null || len > arr.length) len = arr.length; for (var i = 0, arr2 = new Array(len); i < len; i++) arr2[i] = arr[i]; return arr2; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global g, global, Office, self, window, mailbox, mailboxItem, classifications, classifierRegexp, classifiedSubjectRegexp */

var mailboxItem;
var mailbox;
var g = getGlobal();
Office.onReady(function (info) {
  // If needed, Office.js is ready to be called
  mailboxItem = Office.context.mailbox.item;
  mailbox = Office.context.mailbox;
  for (var name in classifications) {
    var classification = classifications[name];
    g[classification.globalFunction] = actionMarkFactory(classification);
  }
  console.log("Office.js is now ready in ".concat(info.host, " on ").concat(info.platform));
});
function getGlobal() {
  return typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : typeof global !== "undefined" ? global : undefined;
}

// The add-in command functions need to be available in global scope
g.validateBody = validateBody;
var classifications = {
  "green": {
    "name": "TLP Green",
    "globalFunction": "actionMarkGreen",
    "subject": "[Classified Green 🟢]",
    "icon80": "IconGreen.80x80"
  },
  "amber": {
    "name": "TLP Amber",
    "globalFunction": "actionMarkAmber",
    "subject": "[Classified Amber 🟠]",
    "icon80": "IconOrange.80x80"
  },
  "red": {
    "name": "TLP Red",
    "globalFunction": "actionMarkRed",
    "subject": "[Classified Red 🔴]",
    "icon80": "IconRed.80x80"
  }
};
var classifierRegexp = /[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]*\[cla[s\u017F][s\u017F]ified (red|green|amber) (?:[\0-\/:-@\[-\^`\{-\u017E\u0180-\u2129\u212B-\uD7FF\uE000-\uFFFF]|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF])\][\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]*/gi;
var classifiedSubjectRegexp = /^(?:[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]?re:[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]?|[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]?aw:[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]?)*[\t-\r \xA0\u1680\u2000-\u200A\u2028\u2029\u202F\u205F\u3000\uFEFF]*\[cla[s\u017F][s\u017F]ified (red|green|amber) (?:[\0-\/:-@\[-\^`\{-\u017E\u0180-\u2129\u212B-\uD7FF\uE000-\uFFFF]|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF])\](?:[\0-\t\x0B\f\x0E-\u2027\u202A-\uD7FF\uE000-\uFFFF]|[\uD800-\uDBFF][\uDC00-\uDFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF])*/i;
function removeClassification(str) {
  return str.replace(classifierRegexp, " ").trim();
}
function addClassificationPrefix(classification, str) {
  if (!classification) {
    return str;
  }
  return classification.subject + " " + str;
}
function getClassification(subject) {
  subject = subject.toLowerCase();
  if (classifiedSubjectRegexp.test(subject)) {
    var matches = subject.match(classifiedSubjectRegexp);
    return classifications[matches[1]];
  } else {
    return null;
  }
}
function normalizeClassification(subject) {
  var classification = getClassification(subject);
  if (!classification) {
    return subject;
  }
  subject = removeClassification(subject);
  subject = addClassificationPrefix(classification, subject);
  return subject;
}
function actionMarkFactory(classification) {
  return function (event) {
    var successMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
      message: "Marked message " + classification.name,
      icon: classification.icon80,
      persistent: false
    };
    var errorMessage = {
      type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
      message: "Failed to mark message (requested " + classification.name + ")"
    };
    setSubjectPrefix(classification, function (ret) {
      if (ret) {
        // Show a notification message
        Office.context.mailbox.item.notificationMessages.replaceAsync("action", successMessage);
      } else {
        // Show an error message
        Office.context.mailbox.item.notificationMessages.replaceAsync("action", errorMessage);
      }

      // Be sure to indicate when the add-in command function is complete
      event.completed();
    });
  };
}

// Set the subject of the item that the user is composing.
function setSubjectPrefix(requestedClassification, callback) {
  // Check conversation history
  findConversationSubjects(mailboxItem.conversationId, function (values) {
    var classifiedConversation = false;
    var classificationConversation = "";
    var _iterator = _createForOfIteratorHelper(values),
      _step;
    try {
      for (_iterator.s(); !(_step = _iterator.n()).done;) {
        value = _step.value;
        var curClassification = getClassification(value);
        if (curClassification) {
          classifiedConversation = true;
          classificationConversation = curClassification;
          break;
        }
      }

      // Check current subject
    } catch (err) {
      _iterator.e(err);
    } finally {
      _iterator.f();
    }
    mailboxItem.subject.getAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        callback(false);
      } else {
        // Successfully got the subject, display it.

        var _subject = asyncResult.value;
        var curClassification = getClassification(_subject);
        if (curClassification) {
          // Item subject classified

          if (classifiedConversation) {
            // Item is marked and part of classified conversation
            if (curClassification.subject === classificationConversation.subject && classificationConversation.subject === requestedClassification.subject) {
              // Classification already matches, normalize
              _subject = normalizeClassification(_subject);
            } else {
              // Do not allow reclassifying	                         
              callback(false);
              return;
            }
            //prefix = curClassification.subject;
          } else {
            // Item is marked and not part of classified conversation, allow changing
            _subject = removeClassification(_subject);
            _subject = addClassificationPrefix(requestedClassification, _subject);
          }
        } else {
          if (classifiedConversation) {
            // Iten is unmarked, and part of classified conversation, force mark
            _subject = addClassificationPrefix(classificationConversation, _subject);
          } else {
            // Proceed with marking image
            _subject = addClassificationPrefix(requestedClassification, _subject);
          }
        }
        mailboxItem.subject.setAsync(_subject, null, function (asyncResult) {
          if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            callback(false);
          } else {
            // Successfully set the subject.
            // Do whatever appropriate for your scenario
            // using the arguments var1 and var2 as applicable.
            callback(true);
          }
        });
      }
    });
  }, function (error) {
    callback(false);
  });
}
function validateBody(event) {
  /*mailboxItem.body.getAsync("html", {
      asyncContext: event
  }, checkBodyOnlyOnSendCallBack);*/
  forceClassificationSubject(event);
}

// Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
// <param name="event">MessageSend event passed from the calling function.</param>
function forceClassificationSubject(event) {
  mailboxItem.subject.getAsync({
    asyncContext: event
  }, function (asyncResult) {
    var subject = asyncResult.value;
    var curClassification = getClassification(subject);
    if (!curClassification) {
      mailboxItem.notificationMessages.addAsync('NoSend', {
        type: 'errorMessage',
        message: 'Please choose a classification for this email.'
      });
      asyncResult.asyncContext.completed({
        allowEvent: false
      });
      return;
    }

    // Got valid classification, force normalization and category
    Office.context.mailbox.item.saveAsync(function callback(result) {
      var itemId = result.value;
      setCategory(itemId, curClassification.name, asyncResult.asyncContext, function (context) {
        subject = normalizeClassification(subject);
        subjectOnSendChange(subject, context);
      });
    });

    // Process the result.
  });
}

function subjectOnSendChange(subject, event) {
  mailboxItem.subject.setAsync(subject, {
    asyncContext: event
  }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      mailboxItem.notificationMessages.addAsync('NoSend', {
        type: 'errorMessage',
        message: 'Unable to set the subject.'
      });

      // Block send.
      asyncResult.asyncContext.completed({
        allowEvent: false
      });
    } else {
      // Allow send.
      asyncResult.asyncContext.completed({
        allowEvent: true
      });
    }
  });
}

// Demo functions adapted from Microsoft
// MIT License, https://github.com/OfficeDev/Office-Add-in-samples

// Invoke by Contoso Subject and CC Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateSubjectAndCC(event) {
  shouldChangeSubjectOnSend(event);
}

// Check if the subject should be changed. If it is already changed allow send. Otherwise change it.
// <param name="event">MessageSend event passed from the calling function.</param>
function shouldChangeSubjectOnSend(event) {
  mailboxItem.subject.getAsync({
    asyncContext: event
  }, function (asyncResult) {
    addCCOnSend(asyncResult.asyncContext);
    //console.log(asyncResult.value);
    // Match string.
    var checkSubject = new RegExp(/\[Checked\]/).test(asyncResult.value);
    // Add [Checked]: to subject line.
    subject = '[Checked]: ' + asyncResult.value;

    // Check if a string is blank, null or undefined.
    // If yes, block send and display information bar to notify sender to add a subject.
    if (asyncResult.value === null || /^\s*$/.test(asyncResult.value)) {
      mailboxItem.notificationMessages.addAsync('NoSend', {
        type: 'errorMessage',
        message: 'Please enter a subject for this email.'
      });
      asyncResult.asyncContext.completed({
        allowEvent: false
      });
    } else {
      // If can't find a [Checked]: string match in subject, call subjectOnSendChange function.
      if (!checkSubject) {
        subjectOnSendChange(subject, asyncResult.asyncContext);
        //console.log(checkSubject);
      } else {
        // Allow send.
        asyncResult.asyncContext.completed({
          allowEvent: true
        });
      }
    }
  });
}

// Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">MessageSend event passed from calling function</param>
function addCCOnSend(event) {
  mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], {
    asyncContext: event
  });
}

// Check if the body contains a specific set of blocked words. If it contains the blocked words, block email from being sent. Otherwise allows sending.
// <param name="asyncResult">MessageSend event passed from the calling function.</param>
function checkBodyOnlyOnSendCallBack(asyncResult) {
  var listOfBlockedWords = new Array("blockedword", "blockedword1", "blockedword2");
  var wordExpression = listOfBlockedWords.join('|');

  // \b to perform a "whole words only" search using a regular expression in the form of \bword\b.
  // i to perform case-insensitive search.
  var regexCheck = new RegExp('\\b(' + wordExpression + ')\\b', 'i');
  var checkBody = regexCheck.test(asyncResult.value);
  if (checkBody) {
    mailboxItem.notificationMessages.addAsync('NoSend', {
      type: 'errorMessage',
      message: 'Blocked words have been found in the body of this email. Please remove them.'
    });
    // Block send.
    asyncResult.asyncContext.completed({
      allowEvent: false
    });
  } else {
    // Allow send.
    asyncResult.asyncContext.completed({
      allowEvent: true
    });
  }
}

// Following functions adapted from easyEws
// GNU Public License v3, https://github.com/davecra/easyEWS

function asyncEws(soap, successCallback, errorCallback) {
  mailbox.makeEwsRequestAsync(soap, function (ewsResult) {
    if (ewsResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log("makeEwsRequestAsync success. " + ewsResult.status);
      var parser = new DOMParser();
      var xmlDoc = parser.parseFromString(ewsResult.value, "text/xml");
      successCallback(xmlDoc);
    } else {
      console.log("makeEwsRequestAsync failed. " + ewsResult.error);
      errorCallback(ewsResult.error);
    }
  });
}
;
function getNodes(node, elementNameWithNS) {
  var elementWithoutNS = elementNameWithNS.substring(elementNameWithNS.indexOf(":") + 1);
  var retVal = node.getElementsByTagName(elementNameWithNS);
  if (retVal == null || retVal.length == 0) {
    retVal = node.getElementsByTagName(elementWithoutNS);
  }
  return retVal;
}
;
function getSoapHeader(request) {
  var result = '<?xml version="1.0" encoding="utf-8"?>' + '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' + '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' + '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' + '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' + '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' + '   <soap:Header>' + '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' + '   </soap:Header>' + '   <soap:Body>' + request + '</soap:Body>' + '</soap:Envelope>';
  return result;
}
;
function findConversationSubjects(conversationId, successCallback, errorCallback) {
  if (!conversationId) {
    // Trivial case, no conversations here
    successCallback([]);
  }
  var soap = '       <m:GetConversationItems>' + '           <m:ItemShape>' + '               <t:BaseShape>IdOnly</t:BaseShape>' + '               <t:AdditionalProperties>' + '                   <t:FieldURI FieldURI="item:Subject" />' + '                   <t:FieldURI FieldURI="item:DateTimeReceived" />' + '               </t:AdditionalProperties>' + '           </m:ItemShape>' + '           <m:FoldersToIgnore>' + '               <t:DistinguishedFolderId Id="deleteditems" />' + '               <t:DistinguishedFolderId Id="drafts" />' + '           </m:FoldersToIgnore>' + '           <m:SortOrder>TreeOrderDescending</m:SortOrder>' + '           <m:Conversations>' + '               <t:Conversation>' + '                   <t:ConversationId Id="' + conversationId + '" />' + '               </t:Conversation>' + '           </m:Conversations>' + '       </m:GetConversationItems>';
  soap = getSoapHeader(soap);
  // Make EWS call
  asyncEws(soap, function (xmlDoc) {
    var nodes = getNodes(xmlDoc, "t:Subject");
    var msgs = [];
    var _iterator2 = _createForOfIteratorHelper(nodes),
      _step2;
    try {
      for (_iterator2.s(); !(_step2 = _iterator2.n()).done;) {
        var msg = _step2.value;
        msgs.push(msg.textContent);
      }
    } catch (err) {
      _iterator2.e(err);
    } finally {
      _iterator2.f();
    }
    successCallback(msgs);
  }, function (errorDetails) {
    if (errorCallback != null) errorCallback(errorDetails);
  });
}
;
function setCategory(itemId, category, context, callback) {
  // ignore missing item ID to improve UX
  if (!itemId) {
    console.log("Ignoring invalid itemId in setCategory: " + itemId);
    callback(context);
    return;
  }
  var soapUpdate = "<UpdateItem MessageDisposition=\"SaveOnly\" ConflictResolution=\"AlwaysOverwrite\" xmlns=\"http://schemas.microsoft.com/exchange/services/2006/messages\">\n\t\t\t<ItemChanges>\n\t\t\t\t<t:ItemChange>\n\t\t\t\t\t<t:ItemId Id=\"" + itemId + "\"/>\n\t\t\t\t\t<t:Updates>\n\t\t\t\t\t\t<t:SetItemField>\n\t\t\t\t\t\t\t<t:ExtendedFieldURI PropertySetId=\"00020329-0000-0000-C000-000000000046\" PropertyName=\"Keywords\" PropertyType=\"StringArray\" />\n\t\t\t\t\t\t\t<t:Message>\n\t\t\t\t\t\t\t\t<t:ExtendedProperty>\n\t\t\t\t\t\t\t\t\t<t:ExtendedFieldURI PropertySetId=\"00020329-0000-0000-C000-000000000046\" PropertyName=\"Keywords\" PropertyType=\"StringArray\" />\n\t\t\t\t\t\t\t\t\t<t:Values>\n\t\t\t\t\t\t\t\t\t\t<t:Value>" + category + "</t:Value>\n\t\t\t\t\t\t\t\t\t</t:Values>\n\t\t\t\t\t\t\t\t</t:ExtendedProperty>\n\t\t\t\t\t\t\t</t:Message>\n\t\t\t\t\t\t</t:SetItemField>\n\t\t\t\t\t</t:Updates>\n\t\t\t\t</t:ItemChange>\n\t\t\t</ItemChanges>\n\t\t</UpdateItem>";
  var soap = getSoapHeader(soapUpdate);
  asyncEws(soap, function (xmlDoc) {
    console.log("Successfully set category: " + xmlDoc);
    callback(context);
  }, function (errorDetails) {
    console.log("Error setting category: " + errorDetails);
    callback(context);
  });
}