Office.onReady((function(){})),Office.actions.associate("action",(function(e){var t={type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:"Performed action.",icon:"Icon.80x80",persistent:!0};Office.context.mailbox.item.notificationMessages.replaceAsync("action",t),e.completed()})),Office.actions.associate("prependHeaderOnSend",(function(e){Office.context.mailbox.item.body.getTypeAsync({asyncContext:e},(function(e){if(e.status!==Office.AsyncResultStatus.Failed){var t=e.value;Office.context.mailbox.item.body.prependOnSendAsync('<div style="border:3px solid #000;padding:15px;"><h1 style="text-align:center;">Contoso Limited</h1></div><br>',{asyncContext:e.asyncContext,coercionType:t},(function(e){e.status!==Office.AsyncResultStatus.Failed?(console.log("The header will be prepended when the mail item is sent."),e.asyncContext.completed()):console.log(e.error.message)}))}else console.log(e.error.message)}))})),Office.actions.associate("appendDisclaimerOnSend",(function(e){Office.context.mailbox.item.body.getTypeAsync({asyncContext:e},(function(e){if(e.status!==Office.AsyncResultStatus.Failed){var t=e.value;Office.context.mailbox.item.body.appendOnSendAsync('<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>',{asyncContext:e.asyncContext,coercionType:t},(function(e){e.status!==Office.AsyncResultStatus.Failed?(console.log("The disclaimer will be appended when the mail item is sent."),e.asyncContext.completed()):console.log(e.error.message)}))}else console.log(e.error.message)}))}));
//# sourceMappingURL=commands.js.map