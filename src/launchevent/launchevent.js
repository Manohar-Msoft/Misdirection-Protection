/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/
const customerDomain = "@microsoft.com";

function onMessageSendHandler(event) {
  let externalRecipients = [];
  let toRecipientsChecked = false;
  let ccRecipientsChecked = false;

  // Function to check recipients in a given field (To or CC)
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

      // Set the corresponding field check flag
      if (fieldName === "To") {
        toRecipientsChecked = true;
      } else if (fieldName === "CC") {
        ccRecipientsChecked = true;
      }

      // Check if both "To" and "CC" checks are completed
      if (toRecipientsChecked && ccRecipientsChecked) {
        if (externalRecipients.length > 0) {
          event.completed({
            allowEvent: false,
            errorMessage:
              "The mail includes some external recipients, are you sure you want to send it?\n\n" +
              externalRecipients.join("\n") +
              "\n\nClick Send to send the mail anyway.",
          });
        } else {
          event.completed({ allowEvent: true });
        }
      }
    });
  }

  // Check "To" recipients
  checkRecipients(Office.context.mailbox.item.to, "To");
  // Check "CC" recipients
  checkRecipients(Office.context.mailbox.item.cc, "CC");
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest's LaunchEvent element to its JavaScript counterpart.
// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
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