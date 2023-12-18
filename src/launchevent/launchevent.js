// /*
// * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
// * See LICENSE in the project root for license information.
// */
// // Import the TelemetryClient class from the Application Insights SDK for JavaScript.
// // const { TelemetryClient } = require("applicationinsights");
// // Create a new TelemetryClient instance.
// // const telemetryClient = new TelemetryClient();
// // import { ApplicationInsights } from '@microsoft/applicationinsights-web';
// // const appInsights = require("applicationinsights");
// // const { useAzureMonitor } = require("@azure/monitor-opentelemetry");
// // const { metrics } = require("@opentelemetry/api");
// // Enable Azure Monitor integration
// // useAzureMonitor();
// let appInsights = require('applicationinsights');
// appInsights.setup("InstrumentationKey=a7c07799-bacf-41dd-9b74-2014f06f64ae;IngestionEndpoint=https://eastus-8.in.applicationinsights.azure.com/;LiveEndpoint=https://eastus.livediagnostics.monitor.azure.com/").setSendLiveMetrics(true).start();
// let client = appInsights.defaultClient;
const customerDomain = "@microsoft.com";
let dialog
let sendEvent
// // let appInsights = require('applicationinsights');
// // Initialize Application Insights
// // var appInsights = new ai.ApplicationInsights({
// //     instrumentationKey: 'a7c07799-bacf-41dd-9b74-2014f06f64ae'
// //   });
// // appInsights.setup("InstrumentationKey=a7c07799-bacf-41dd-9b74-2014f06f64ae;IngestionEndpoint=https://eastus-8.in.applicationinsights.azure.com/;LiveEndpoint=https://eastus.livediagnostics.monitor.azure.com/").enableWebInstrumentation(true).start();
// // let client = appInsights.defaultClient;
// // appInsights.loadAppInsights();

function onMessageSendHandler(event){
    let externalRecipients = [];
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
                        externalRecipients.push(`${recipient.emailAddress} in ${fieldName}`);
                    }
                });
            }

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
                  const url = 'https://localhost:3000/dialog.html'
                  Office.context.ui.displayDialogAsync(url, { height: 30, width: 50, displayInIframe: true },
                    function (asyncResult) {
                      //  If dialog failed to open (probably popup blocker) then do 'dialogClosed' function.
                      console.log(asyncResult.status)
                      dialog = asyncResult.value
                    //   dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog)
                      dialog.addEventHandler(Office.EventType.DialogMessageReceived, closeDialog)
                    //   if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    //     Office.context.ui.closeContainer()
                    //     // If dialog box already open, close dialog and do not send email.
                    //     if (asyncResult.error.code === 12007) {
                    //       dialog.addEventHandler(Office.EventType.DialogMessageReceived, closeDialog)
                    //     }
                    //     // event.completed({ allowEvent: false })
                    //   } else {
                    //     dialog = asyncResult.value
                    //     //  Once dialog box has sent message to confirm it is ready. Send dialog box the recipient emails.
                    //     // dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog)
                    //     //  If dialog  sends event (probably user closes), then do not send the email.
                    //     dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed)
                    //   }
                    })
//                     // // Create an event telemetry object.
//                     // let eventTelemetry = {
//                     //     name: "testEvent"
//                     // };
                
//                     // // Send the event telemetry object to Azure Monitor Application Insights.
//                     // telemetryClient.trackEvent(eventTelemetry);
//                     // Log event when the pop-up is displayed
//                     // client.trackEvent({
//                     //     name: 'PopupDisplayed',
//                     //     properties: {
//                     //         recipients: externalRecipients.join(';'),
//                     //     },
//                     // });
//                     // const meter =  metrics.getMeter("testMeter");

//                     // // Create a histogram metric
//                     // let histogram = meter.createHistogram("histogram");

//                     // // Record values to the histogram metric with different tags
//                     // histogram.record(1, { "testKey": "testValue" });
//                     // histogram.record(30, { "testKey": "testValue2" });
//                     // histogram.record(100, { "testKey2": "testValue" });
//                     event.completed({
//                         allowEvent: false,
//                         errorMessage:
//                             "The mail includes some external recipients, are you sure you want to send it?\n\n" +
//                             externalRecipients.join("\n") +
//                             "\n\nClick Send to send the mail anyway.",
//                     });

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
//       dialog.messageChild(JSON.stringify(allRecipientData.concat(Office.context.mailbox.item.itemType)), { targetOrigin: '*' })
//       dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
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
  }

// setTimeout(function () {
//     // Your action to be performed after the timer (e.g., confirm send or take any other action)
//     if (confirm("Are you sure you want to send this message?")) {
//       // User confirmed, proceed with sending the message
//       eventArgs.completed({ allowEvent: true });
//     } else {
//       // User canceled, prevent the message from being sent
//       eventArgs.completed({ allowEvent: false });
//     }
//   }, 5000);
// let appInsights = require("applicationinsights");
// appInsights.setup("InstrumentationKey=a7c07799-bacf-41dd-9b74-2014f06f64ae;IngestionEndpoint=https://eastus-8.in.applicationinsights.azure.com/;LiveEndpoint=https://eastus.livediagnostics.monitor.azure.com/")
    
// // appInsights.setAutoCollectConsole(true, true) // the second true argument enables tracing the console methods 
// appInsights.start();

// const methodName = "My method";
// const count = 10;
// console.log("Function %s is called %d times ", methodName, count); // severity level: 1
// console.info("Here is a sample info"); // severity level: 1
// console.warn("Here is a sample warn"); //severity level: 2
// console.error("Here is a sample error"); //severity level: 2

// import { ApplicationInsights } from '@microsoft/applicationinsights-web';
// import { ClickAnalyticsPlugin } from '@microsoft/applicationinsights-clickanalytics-js';

// const clickPluginInstance = new ClickAnalyticsPlugin();
// // Click Analytics configuration
// const clickPluginConfig = {
//   autoCapture: true
// };
// // Application Insights Configuration
// const configObj = {
//   connectionString: "InstrumentationKey=a7c07799-bacf-41dd-9b74-2014f06f64ae;IngestionEndpoint=https://eastus-8.in.applicationinsights.azure.com/;LiveEndpoint=https://eastus.livediagnostics.monitor.azure.com/", 
//   // Alternatively, you can pass in the instrumentation key,
//   // but support for instrumentation key ingestion will end on March 31, 2025.  
//   // instrumentationKey: "YOUR INSTRUMENTATION KEY",
//   extensions: [clickPluginInstance],
//   extensionConfig: {
//     [clickPluginInstance.identifier]: clickPluginConfig
//   },
// };

// const appInsights = new ApplicationInsights({ config: configObj });
// appInsights.loadAppInsights();

// import { ApplicationInsights } from '@microsoft/applicationinsights-web'

// const appInsights = new ApplicationInsights({ config: {
//   connectionString: 'InstrumentationKey=a7c07799-bacf-41dd-9b74-2014f06f64ae;IngestionEndpoint=https://eastus-8.in.applicationinsights.azure.com/;LiveEndpoint=https://eastus.livediagnostics.monitor.azure.com/"'
//   /* ...Other Configuration Options... */
// } });
// appInsights.loadAppInsights();
// appInsights.trackPageView();
// const express = require('express');
// const app = express();

// // ... Other routes and configurations ...

// // Health check endpoint
// app.get('/health', (req, res) => {
//   res.status(200).send('OK');
// });

// // ... Start server and other code ...

// const port = process.env.PORT || 3000;
// app.listen(port, () => {
//   console.log(`Server is running on port ${port}`);
// });

// function to handle the ItemSend event
// function itemSendHandler(event) {
//     // Display a custom dialog using DialogAPI
//     Office.context.ui.displayDialogAsync('https://www.contoso.com/dialog.html', { height: 200, width: 400 }, function (result) {
//       var dialog = result.value;
  
//       dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
//         // Handle the message received from the dialog
//         var message = arg.message;
  
//         // Implement your logic based on the message from the dialog
//         if (message === 'dontSend') {
//           // User clicked "Don't Send" button
//           console.log("Don't Send clicked");
//         } else if (message === 'anywaysSend') {
//           // User clicked "Anyways Send" button
//           console.log("Anyways Send clicked");
//         }
  
//         // Close the dialog after handling the message
//         dialog.close();
//       });
  
//       dialog.addEventHandler(Office.EventType.DialogEventReceived, function (arg) {
//         // Handle the events received from the dialog (e.g., dialog closed)
//         console.log('DialogEventReceived', arg.error);
  
//         // Optionally, handle any cleanup or additional logic
//       });
//     });
  
//     // Prevent the default behavior (e.g., sending the item)
//     event.completed({ allowEvent: false });
//   }
  
//   // Register the function for the ItemSend event
//   Office.actions.associate("itemSendHandler", itemSendHandler);
  

// let dialog
// let recipients
// let allRecipientData
// let item
// let sendEvent

// Office.onReady(() => {
//   // Initialise Office JS
// })

// /**
//  * Function that is run when the send button is pressed by the user.
//  * @param {object} event - The email send event that is to be controlled.
//  */
// function openDialog (event) {
//   //  Get email compose information from Outlook (using promises since they are asynchronous functions).
//   item = Office.context.mailbox.item

//   // Verify if the composed item is an appointment or message.
//   let promise1
//   let promise2
//   let promise3
//   let promise4
//   if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
//     promise1 = getToEmailsAppointment()
//     promise2 = getCCEmails_appointment()
//     // Use promises to ensure required and optional attendees have been fetched.
//     promise4 = Promise.all([promise1, promise2]).then(function (result) {
//       allRecipientData = result
//       return allRecipientData
//     })
//   } else {
//     promise1 = getToEmails()
//     promise2 = getCCEmails()
//     promise3 = getBCCEmails()
//     // Use promises to ensure bcc, cc and to recipients have been fetched.
//     promise4 = Promise.all([promise1, promise2, promise3]).then(function (result) {
//       allRecipientData = result
//       return allRecipientData
//     })
//   }
  
//   //  Check if multiple external recipients are present to decide to display dialog box.
//   promise4.then(function (result) {
//     sendEvent = event
//     console.log(allRecipientData)
//     const multipleExternalBool = checkMultipleExternal(processEmails(allRecipientData))
//     if (!multipleExternalBool) {
//       event.completed({ allowEvent: true })
//     } else {
//       //  Display dialog box (callback function in dialog is to create event handler in host page to recieve info from dialog page).
//       const url = 'https://localhost:3000/dialog.html'
//       Office.context.ui.displayDialogAsync(url, { height: 50, width: 50, displayInIframe: true },
//         function (asyncResult) {
//           //  If dialog failed to open (probably popup blocker) then do 'dialogClosed' function.
//           console.log(asyncResult.status)
//           if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//             Office.context.ui.closeContainer()
//             // If dialog box already open, close dialog and do not send email.
//             if (asyncResult.error.code === 12007) {
//               dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
//             }
//             event.completed({ allowEvent: false })
//           } else {
//             dialog = asyncResult.value
//             //  Once dialog box has sent message to confirm it is ready. Send dialog box the recipient emails.
//             dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailsToDialog)
//             //  If dialog  sends event (probably user closes), then do not send the email.
//             dialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed)
//           }
//         })
//     }
//   })
// }

// /**
//  * Function that prevents the message from being sent when the dialog is closed without correct input from the user.
//  */
// function dialogClosed () {
//   sendEvent.completed({ allowEvent: false })
// }

// /**
//  * Function that sends the initial recipient data to the dialog box.
//  * @param {object} arg - The message object that is passed from the host to the dialog that contains the emails .
//  */
// function sendEmailsToDialog (arg) {
//   if (JSON.parse(arg.message).messageType === 'initialise') {
//     dialog.messageChild(JSON.stringify(allRecipientData.concat(Office.context.mailbox.item.itemType)), { targetOrigin: '*' })
//     dialog.addEventHandler(Office.EventType.DialogMessageReceived, sendEmailwithUpdatedRecipients)
//   }
// }

// /**
//  * Function that sends the message with the updated recipients from the checkbox form. The message doesn't send if  there are no recipients.
//  * @param {object} arg - A message object from the dialog that contains selected recipient data from the checkbox form.
//  */
// function sendEmailwithUpdatedRecipients (arg) {
// //   $(window).bind('resize', function (e) { dialog.close() })
// //   const message = JSON.parse(arg.message)
// //   // If checkbox form results recieved from dialog, send with selected recipients otherwise do not send email and close dialog.
// //   if (message.messageType === 'form_output') {
// //     if ((message.toRecipients.length + message.ccRecipients.length + message.bccRecipients) === 0) {
// //       dialog.close()
// //       sendEvent.completed({ allowEvent: false })
// //     } else {
// //       setRecipients(message.toRecipients, message.ccRecipients, message.bccRecipients)
// //       sendEvent.completed({ allowEvent: true })
// //       dialog.close()
// //     }
// //   } else if (message.messageType === 'cancel') {
// //     dialog.close()
// //     sendEvent.completed({ allowEvent: false })
// //   }
// }

// /**
//  * Function that updates the 'to' and 'cc' (or 'required' and 'optional') fields in Outlook.
//  * @param {object} toRecipients - Object that contains the 'to' or 'required' recipients data.
//  * @param {object} ccRecipients - Object that contains the 'cc' or 'optional' recipients data.
//  */
// function setRecipients (toRecipients, ccRecipients, bccRecipients) {
//   // Local objects to point to recipients of either the appointment or message that is being composed.
//   // bccRecipients applies to only messages, not appointments.
//   let RecipientsTo, RecipientsCC, RecipientsBCC
//   item = Office.context.mailbox.item
//   // Verify if the composed item is an appointment or message.
//   if (item.itemType === Office.MailboxEnums.ItemType.Appointment) {
//     RecipientsTo = item.requiredAttendees
//     RecipientsCC = item.optionalAttendees
//   } else {
//     RecipientsTo = item.to
//     RecipientsCC = item.cc
//     RecipientsBCC = item.bcc
//   }

//   // Use asynchronous method setAsync to set each type of recipients
//   // of the composed item. Each time, this example passes a set of
//   // names and email addresses to set, and an anonymous
//   // callback function that doesn't take any parameters.
//   RecipientsTo.setAsync(toRecipients,
//     function (asyncResult) {
//       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//         write(asyncResult.error.message)
//       } else {
//         // Async call to set to-recipients of the item completed.
//       }
//     }) // End to setAsync.

//   // Set any cc-recipients.
//   RecipientsCC.setAsync(ccRecipients,
//     function (asyncResult) {
//       if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//         write(asyncResult.error.message)
//       } else {
//         // Async call to set cc-recipients of the item completed.
//       }
//     }) // End cc setAsync.
  
//   if (item.itemType !== Office.MailboxEnums.ItemType.Appointment) {
//     // Set any cc-recipients.
//     RecipientsBCC.setAsync(bccRecipients,
//       function (asyncResult) {
//         if (asyncResult.status === Office.AsyncResultStatus.Failed) {
//           write(asyncResult.error.message)
//         } else {
//           // Async call to set cc-recipients of the item completed.
//         }
//       }) // End bcc setAsync.
//   }
// }

// /**
//  * A function that gets the 'to' recipient data from the email.
//  */
// function getToEmails () {
//   return new Promise(function (resolve, reject) {
//     try {
//       Office.context.mailbox.item.to.getAsync(function (asyncResult) {
//         resolve(asyncResult.value)
//       })
//     }
//     catch (error) {
//       reject('Error')
//     }
//   })
// }

// /**
//  * A function that gets the 'cc' recipient data from the email.
//  */
// function getCCEmails () {
//   return new Promise(function (resolve, reject) {
//     try {
//       Office.context.mailbox.item.cc.getAsync(function (asyncResult) {
//         resolve(asyncResult.value)
//       })
//     }
//     catch (error) {
//       reject('Error')
//     }
//   })
// }

// /**
//  * A function that gets the 'bcc' recipient data from the email.
//  */
// function getBCCEmails () {
//   return new Promise(function (resolve, reject) {
//     try {
//       Office.context.mailbox.item.bcc.getAsync(function (asyncResult) {
//         resolve(asyncResult.value)
//       })
//     }
//     catch (error) {
//       reject('Error')
//     }
//   })
// }

// /**
//  * A function that gets the 'required' recipient data from the meeting request.
//  */
// function getToEmailsAppointment() {
//   return new Promise(function (resolve, reject) {
//     try {
//       Office.context.mailbox.item.requiredAttendees.getAsync(function (asyncResult) {
//         resolve(asyncResult.value)
//       })
//     }
//     catch (error) {
//       reject('Error')
//     }
//   })
// }

// /**
//  * A function that gets the 'optional' recipient data from the meeting request.
//  */
// function getCCEmails_appointment () {
//   return new Promise(function (resolve, reject) {
//     try {
//       Office.context.mailbox.item.optionalAttendees.getAsync(function (asyncResult) {
//         resolve(asyncResult.value)
//       })
//     }
//     catch (error) {
//       reject('Error')
//     }
//   })
// }

// /**
//  * Function that a recipient data object and returns just the email addresses as an array.
//  * @param {array} result - An array containing the recipient data.
//  */
// function processEmails (result) {
//   // Combine cc and to recipients if needed.
//   let recipientData
//   if (result.length > 2) {
//     recipientData = result[0].concat(result[1]).concat(result[2])
//   } else if (result.length > 1) {
//     recipientData = result[0].concat(result[1])
//   } else {
//     recipientData = result[0]
//   }
//   // Add email address information to a list.
//   let emails = []
//   for (let i = 0; i < recipientData.length; i++) {
//     let Email = recipientData[i].emailAddress
//     emails.push(Email)
//   }
//   return emails
// }

// /**
//  * Function that returns a boolean value based on if the number of external emails is larger than
//  * @param {array} emails - An array containing the emails to be checked.
//  */
// function checkMultipleExternal (emails) {
//   // Create list of external emails.
//   let externalEmails = []
//   for (let i = 0; i < emails.length; i++) {
//     let domain = emails[i].slice(emails[i].indexOf('@'), emails[i].length).toUpperCase()
//     if (domain !== '@microsoft.com') {
//       externalEmails.push(domain)
//     }
//   }
//   // Return true if number of unique external domains is more than 1.
//   const numberExternalDomains = new Set(externalEmails).size
//   return (numberExternalDomains > 1)
// }