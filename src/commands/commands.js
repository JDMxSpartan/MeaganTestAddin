/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  const message = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };

  // Show a notification message.
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);

function prependHeaderOnSend(event) {
  // It's recommended to call the getTypeAsync method and pass its returned value to the options.coercionType parameter of the prependOnSendAsync call.
  Office.context.mailbox.item.body.getTypeAsync(
    {
      asyncContext: event
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
      }

      // Sets the header to be prepended to the body of the message on send.
      const bodyFormat = asyncResult.value;
      // Because of the various ways in which HTML text can be formatted, the content may render differently when it's prepended to the mail item body.
      // In this scenario, a <br> tag is added to the end of the HTML string to preserve its format.
      const header = '<div style="border:3px solid #000;padding:15px;"><h1 style="text-align:center;">Contoso Limited</h1></div><br>';

      Office.context.mailbox.item.body.prependOnSendAsync(
        header,
        {
          asyncContext: asyncResult.asyncContext,
          coercionType: bodyFormat
        },
        (asyncResult) => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }

          console.log("The header will be prepended when the mail item is sent.");
          asyncResult.asyncContext.completed();
        }
      );
  });
}

function appendDisclaimerOnSend(event) { 
  // Calls the getTypeAsync method and passes its returned value to the options.coercionType parameter of the appendOnSendAsync call.
  Office.context.mailbox.item.body.getTypeAsync(
    {
      asyncContext: event
    }, 
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.log(asyncResult.error.message);
        return;
    }

    // Sets the disclaimer to be appended to the body of the message on send.
    const bodyFormat = asyncResult.value;
    const disclaimer =
      '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  

    Office.context.mailbox.item.body.appendOnSendAsync(
      disclaimer,
      {
        asyncContext: asyncResult.asyncContext,
        coercionType: bodyFormat
      },
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
          return;
        }

        console.log("The disclaimer will be appended when the mail item is sent.");
        asyncResult.asyncContext.completed();
      }
    );
  });
}

Office.actions.associate("prependHeaderOnSend", prependHeaderOnSend);
Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
