// /*
// * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// * See LICENSE in the project root for license information.
// */
const customerDomain = "@microsoft.com";
let dialog
let sendEvent
let externalRecipients = [];
let toRecipients = [];
let ccRecipients = [];
let bccRecipients = [];
function onMessageSendHandler(event){
    // let externalRecipients = [];
    let toRecipientsChecked = false;
    let ccRecipientsChecked = false;
    let bccRecipientsChecked = false;
    sendEvent = event

//     // Function to check recipients in a given field (To or CC or BCC)
    function checkRecipients(field, fieldName) {
        field.getAsync((asyncResult) => {
            if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                var recipients = asyncResult.value;
                recipients.forEach((recipient) => {
                    if (!recipient.emailAddress.includes(customerDomain)) {
                        if(fieldName==="To")
                        {
                            toRecipients.push(`${recipient.emailAddress}`);
                        }
                        else if(fieldName === "CC"){
                            ccRecipients.push(`${recipient.emailAddress}`);
                        }
                        else if(fieldName === "BCC"){
                            bccRecipients.push(`${recipient.emailAddress}`);
                        }
                    }
                });
            }
    
        externalRecipients = [toRecipients,ccRecipients,bccRecipients];

//             // Set the corresponding field check flag
            if (fieldName === "To") {
                toRecipientsChecked = true;
            } else if (fieldName === "CC") {
                ccRecipientsChecked = true;
            } else if (fieldName === "BCC"){
              bccRecipientsChecked = true;
            }

//             // Check if both "To" and "CC" checks are completed
            if (toRecipientsChecked && ccRecipientsChecked && bccRecipientsChecked) {
                if (externalRecipients.length > 0) {
                //   client.trackEvent({name: "my custom event", properties: {customProperty: "custom property value"}});
                //   document.getElementById('mainLabel').innerHTML = "Hello!!You are successful"
                  const url = 'https://localhost:3000/dialog.html'
                  Office.context.ui.displayDialogAsync(url, { height: 50, width: 50, displayInIframe: true },
                    function (asyncResult) {
                      //  If dialog failed to open (probably popup blocker) then do 'dialogClosed' function.
                      console.log(asyncResult.status)
                      dialog = asyncResult.value
                    //   dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog)
                      dialog.addEventHandler(Office.EventType.DialogMessageReceived, closeDialog)
                    })                   

                } else {
                    event.completed({ allowEvent: true });
                }
            }
        });
    }

//     // Check "To" recipients
    checkRecipients(Office.context.mailbox.item.to, "To");
//     // Check "CC" recipients
    checkRecipients(Office.context.mailbox.item.cc, "CC");
//     // check "BCC" recipients
    checkRecipients(Office.context.mailbox.item.bcc, "BCC");
}


// // IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    // Office.actions.associate("itemSendHandler", itemSendHandler);

}
// /**
//  * Function that sends the initial recipient data to the dialog box.
//  * @param {object} arg - The message object that is passed from the host to the dialog that contains the emails .
//  */
// function sendEmailsToDialog (arg) {
//     if (JSON.parse(arg.message).messageType === 'initialise') {
//       dialog.messageChild(JSON.stringify(externalRecipients.concat(Office.context.mailbox.item.itemType)), { targetOrigin: '*' })
//     //   dialog.messageChild(JSON.stringify(externalRecipients.concat(Office.context.mailbox.item.itemType)), { targetOrigin: '*' })
//     //   dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
//     }
//   }
/**
 * Function that sends the message with the updated recipients from the checkbox form. The message doesn't send if  there are no recipients.
 * @param {object} arg - A message object from the dialog that contains selected recipient data from the checkbox form.
 */
function closeDialog (arg) {
    // $(window).bind('resize', function (e) { dialog.close() })
    const message = JSON.parse(arg.message)
    if (message.messageType === "cancel") {
      dialog.close();
      sendEvent.completed({ allowEvent: false });
    }
    else if(message.messageType === "sendanyways")
    {
       sendEvent.completed({ allowEvent: true });
       dialog.close();
    }
    else{
        dialog.messageChild(JSON.stringify(externalRecipients), { targetOrigin: '*' })
    }
  }

