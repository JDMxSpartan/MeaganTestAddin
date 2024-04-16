/*
* Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
* See LICENSE in the project root for license information.
*/

function onMessageSendHandler(event) {
    Office.context.mailbox.item.body.getAsync(
      "text",
      { asyncContext: event },
      getBodyCallback
    );
  }
  
  function getBodyCallback(asyncResult){
    const event = asyncResult.asyncContext;
    let body = "";
    if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
      body = asyncResult.value;
    } else {
      const message = "Failed to get body text";
      console.error(message);
      event.completed({ allowEvent: false, errorMessage: message });
      return;
    }
  
    const banner = getBannerFromBody(body);
    // Check if the banner is null error
    bannerNullHandler(banner, event);

    // const matches = hasMatches(body);
    // if (matches) {
    //   Office.context.mailbox.item.getAttachmentsAsync(
    //     { asyncContext: event },
    //     getAttachmentsCallback);
    // } else {
    //   event.completed({ allowEvent: true });
    // }
  }
  
  function getBannerFromBody(body) {
    const banner_regex =
      /^(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im;
  
    const banner = body.match(banner_regex);
    console.log(banner);
    if (banner) {
      console.log("banner found");
      return banner[0];
    } else {
      console.log("banner null");
      return null;
    }
  }

  function bannerNullHandler(banner, event){

    if (banner == null) {
      event.completed(
        {
            allowEvent: false,
             cancelLabel: "Ok",
             commandId: "msgComposeOpenPaneButton",
             contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
             errorMessage: "Please enter a banner, banner error detected.",
        //     //underneath with enable the user to press send anyways, might need later
             sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
        }
        );
      }
      else{
        event.completed({ allowEvent: true });
      }
    }


  function hasMatches(body) {
    if (body == null || body == "") {
      return false;
    }
  
    const arrayOfTerms = ["send", "picture", "document", "attachment"];
    for (let index = 0; index < arrayOfTerms.length; index++) {
      const term = arrayOfTerms[index].trim();
      const regex = RegExp(term, 'i');
      if (regex.test(body)) {
        return true;
      }
    }
  
    return false;
  }
  
  function getAttachmentsCallback(asyncResult) {
    const event = asyncResult.asyncContext;
    if (asyncResult.value.length > 0) {
      for (let i = 0; i < asyncResult.value.length; i++) {
        if (asyncResult.value[i].isInline == false) {
          event.completed({ allowEvent: true });
          return;
        }
      }
  
      event.completed({
        allowEvent: false,
        errorMessage: "Looks like the body of your message includes an image or an inline file. Would you like to attach a copy of it to the message?",
        cancelLabel: "Attach a copy",
        commandId: "msgComposeOpenPaneButton",
        sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
      });
    } else {
      event.completed({
        allowEvent: false,
        errorMessage: "Looks like you're forgetting to include an attachment blayh blah testing.",
        cancelLabel: "Add an attachment",
        commandId: "msgComposeOpenPaneButton"
      });
    }
  }
  
  // IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
  if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  }