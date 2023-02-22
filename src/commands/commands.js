/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window, mailbox, mailboxItem */


let mailboxItem;
let mailbox;
Office.onReady(function(info) {
    // If needed, Office.js is ready to be called


    mailboxItem = Office.context.mailbox.item;
    mailbox = Office.context.mailbox;

    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);


});

function action(event) {
    const message = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Performed action.",
        icon: "Icon.80x80",
        persistent: true,
    };

    // Show a notification message
    Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

const g = getGlobal();

function getGlobal() {
    return typeof self !== "undefined" ?
        self :
        typeof window !== "undefined" ?
        window :
        typeof global !== "undefined" ?
        global :
        undefined;
}

// The add-in command functions need to be available in global scope
g.action = action;
g.actionMarkGreen = actionMarkGreen;
g.actionMarkAmber = actionMarkAmber;
g.actionMarkRed = actionMarkRed;
g.validateBody = validateBody;

/* TODO dynamic evaluation of action functions */

function actionMarkGreen(event) {

    const successMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Marked message green",
        icon: "IconGreen.80x80",
        persistent: false,
    };

    const errorMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Failed to mark message green",
        icon: "IconGreen.80x80",
        persistent: true,
    };

    setSubjectPrefix("[Classified Green 🟢]", function(ret) {

        if (ret) {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", successMessage);

        } else {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", errorMessage);
        }


        // Be sure to indicate when the add-in command function is complete
        event.completed();

    });
}


function actionMarkAmber(event) {
    const successMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Marked message amber",
        icon: "IconAmber.80x80",
        persistent: false,
    };

    const errorMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Failed to mark message amber",
        icon: "IconAmber.80x80",
        persistent: true,
    };

    setSubjectPrefix("[Classified Amber 🟠]", function(ret) {

        if (ret) {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", successMessage);
        } else {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", errorMessage);
        }


        // Be sure to indicate when the add-in command function is complete
        event.completed();

    });
}



function actionMarkRed(event) {
    const successMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Marked message red",
        icon: "IconRed.80x80",
        persistent: false,
    };

    const errorMessage = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Failed to mark message red",
        icon: "IconRed.80x80",
        persistent: true,
    };

    setSubjectPrefix("[Classified Red 🔴]", function(ret) {

        if (ret) {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", successMessage);

        } else {
            // Show a notification message
            Office.context.mailbox.item.notificationMessages.replaceAsync("action", errorMessage);

        }


        // Be sure to indicate when the add-in command function is complete
        event.completed();

    });
}

let classifications = {
    "green": {
        "subject": "[Classified Green 🟢]"
    },
    "amber": {
        "subject": "[Classified Amber 🟠]"
    },
    "red": {
        "subject": "[Classified Red 🔴]"
    }
}

const regexp = /^(?:\s?re:\s?|\s?awr:\s?)*\s?\[classified (red|green|amber) \W\].*/u;

function checkSubjectClassified(subject) {
    subject = subject.toLowerCase();
    if (regexp.test(subject)) {

        console.log("Getting matches for " + subject);
        let matches = subject.match(regexp);

        return classifications[matches[1]];
    } else {
        return false;
    }
}

// Set the subject of the item that the user is composing.
function setSubjectPrefix(prefix, callback) {

    // Check conversation history
    findConversationSubjects(mailboxItem.conversationId, function(values) {

        let classifiedConversation = false;
        let classification = "";

        for (value of values) {
            let curClassification = checkSubjectClassified(value);
            if (curClassification) {
                classifiedConversation = true;
                classification = curClassification;
                break;
            }
        }

        // Check current subject
        mailboxItem.subject.getAsync(
            function(asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    console.log(asyncResult.error.message);
                    callback(false);
                } else {
                    // Successfully got the subject, display it.

                    let curClassification = checkSubjectClassified(asyncResult.value)

                    if (curClassification) {
                        // Item subject classified

                        if (classifiedConversation) {
                            // TODO refactor into preflight methods, force marking for now
                            // Item is marked and part of classified conversation
                            
                            if (curClassification.subject === classification.subject && classification.subject === prefix) {
								// Classification already matches, nothing to do :)
								callback(true);
								return;
							} else {	                         
                                callback(false);
                                return;
                            }
                            //prefix = curClassification.subject;
                        } else {
                            // Item is marked and part of classified conversation
                            callback(false);
                            return;
                        }


                    } else {

                        if (classifiedConversation) {
                            // TODO refactor into preflight methods, force marking for now
                            // Iten is unmarked, and part of classified conversation, force mark
                            prefix = classification.subject;
                        } else {
                            // Proceed with marking image
                        }

                    }

                    subject = prefix + ' ' + asyncResult.value;

                    mailboxItem.subject.setAsync(
                        subject, null,
                        function(asyncResult) {
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



    }, function(error) {

        callback(false);

    });



}

// Entry point for Contoso Message Body Checker add-in before send is allowed.
// <param name="event">MessageSend event is automatically passed by BlockOnSend code to the function specified in the manifest.</param>
function validateBody(event) {
    mailboxItem.body.getAsync("html", {
        asyncContext: event
    }, checkBodyOnlyOnSendCallBack);
}

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
        },
        function(asyncResult) {
            addCCOnSend(asyncResult.asyncContext);
            //console.log(asyncResult.value);
            // Match string.
            var checkSubject = (new RegExp(/\[Checked\]/)).test(asyncResult.value)
            // Add [Checked]: to subject line.
            subject = '[Checked]: ' + asyncResult.value;

            // Check if a string is blank, null or undefined.
            // If yes, block send and display information bar to notify sender to add a subject.
            if (asyncResult.value === null || (/^\s*$/).test(asyncResult.value)) {
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

        }
    )
}

// Add a CC to the email.  In this example, CC contoso@contoso.onmicrosoft.com
// <param name="event">MessageSend event passed from calling function</param>
function addCCOnSend(event) {
    mailboxItem.cc.setAsync(['Contoso@contoso.onmicrosoft.com'], {
        asyncContext: event
    });
}

// Check if the subject should be changed. If it is already changed allow send, otherwise change it.
// <param name="subject">Subject to set.</param>
// <param name="event">MessageSend event passed from the calling function.</param>
function subjectOnSendChange(subject, event) {
    mailboxItem.subject.setAsync(
        subject, {
            asyncContext: event
        },
        function(asyncResult) {
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



// Borrowed from easyEws

function asyncEws(soap, successCallback, errorCallback) {

    console.log("Starting call");
    mailbox.makeEwsRequestAsync(soap, function(ewsResult) {
        console.log("Returning from call");
        if (ewsResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("makeEwsRequestAsync success. " + ewsResult.status);
            let parser = new DOMParser();
            let xmlDoc = parser.parseFromString(ewsResult.value, "text/xml");
            successCallback(xmlDoc);

        } else {
            console.log("makeEwsRequestAsync failed. " + ewsResult.error);
            errorCallback(ewsResult.error);
        }
    });


};


function getNodes(node, elementNameWithNS) {
    /** @type {string} */
    var elementWithoutNS = elementNameWithNS.substring(elementNameWithNS.indexOf(":") + 1);
    /** @type {array} */
    var retVal = node.getElementsByTagName(elementNameWithNS);
    if (retVal == null || retVal.length == 0) {
        retVal = node.getElementsByTagName(elementWithoutNS);
    }
    return retVal;
};

function getSoapHeader(request) {
    /** @type {string} */
    var result =
        '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '   <soap:Header>' +
        '       <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '   </soap:Header>' +
        '   <soap:Body>' + request + '</soap:Body>' +
        '</soap:Envelope>';
    return result;
};

function findConversationSubjects(conversationId, successCallback, errorCallback) {
    /** @type {string} */
    var soap =
        '       <m:GetConversationItems>' +
        '           <m:ItemShape>' +
        '               <t:BaseShape>IdOnly</t:BaseShape>' +
        '               <t:AdditionalProperties>' +
        '                   <t:FieldURI FieldURI="item:Subject" />' +
        '                   <t:FieldURI FieldURI="item:DateTimeReceived" />' +
        '               </t:AdditionalProperties>' +
        '           </m:ItemShape>' +
        '           <m:FoldersToIgnore>' +
        '               <t:DistinguishedFolderId Id="deleteditems" />' +
        '               <t:DistinguishedFolderId Id="drafts" />' +
        '           </m:FoldersToIgnore>' +
        '           <m:SortOrder>TreeOrderDescending</m:SortOrder>' +
        '           <m:Conversations>' +
        '               <t:Conversation>' +
        '                   <t:ConversationId Id="' + conversationId + '" />' +
        '               </t:Conversation>' +
        '           </m:Conversations>' +
        '       </m:GetConversationItems>';
    soap = getSoapHeader(soap);
    console.log("got soap header");
    // Make EWS call
    asyncEws(soap, function(xmlDoc) {
        let nodes = getNodes(xmlDoc, "t:Subject");
        var msgs = [];
        for (msg of nodes) {
            msgs.push(msg.textContent);
        }
        successCallback(msgs);

    }, function(errorDetails) {
        if (errorCallback != null)
            errorCallback(errorDetails);
    });
};
