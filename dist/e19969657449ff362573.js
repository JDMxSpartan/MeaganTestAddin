function _slicedToArray(e,t){return _arrayWithHoles(e)||_iterableToArrayLimit(e,t)||_unsupportedIterableToArray(e,t)||_nonIterableRest()}function _nonIterableRest(){throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}function _unsupportedIterableToArray(e,t){if(e){if("string"==typeof e)return _arrayLikeToArray(e,t);var n=Object.prototype.toString.call(e).slice(8,-1);return"Object"===n&&e.constructor&&(n=e.constructor.name),"Map"===n||"Set"===n?Array.from(e):"Arguments"===n||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(n)?_arrayLikeToArray(e,t):void 0}}function _arrayLikeToArray(e,t){(null==t||t>e.length)&&(t=e.length);for(var n=0,r=new Array(t);n<t;n++)r[n]=e[n];return r}function _iterableToArrayLimit(e,t){var n=null==e?null:"undefined"!=typeof Symbol&&e[Symbol.iterator]||e["@@iterator"];if(null!=n){var r,o,a,i,l=[],c=!0,s=!1;try{if(a=(n=n.call(e)).next,0===t){if(Object(n)!==n)return;c=!1}else for(;!(c=(r=a.call(n)).done)&&(l.push(r.value),l.length!==t);c=!0);}catch(e){s=!0,o=e}finally{try{if(!c&&null!=n.return&&(i=n.return(),Object(i)!==i))return}finally{if(s)throw o}}return l}}function _arrayWithHoles(e){if(Array.isArray(e))return e}function onMessageSendHandler(e){Promise.all([getToRecipientsAsync(),getSenderAsync(),getBodyAsync(),getCCAsync(),getBCCAsync()]).then((function(t){var n=_slicedToArray(t,5),r=n[0],o=n[1],a=n[2],i=n[3],l=n[4];console.log("To recipients:"),r.forEach((function(e){return console.log(e.emailAddress)})),console.log("Sender:"+o.displayName+" "+o.emailAddress),console.log("CC: "+i.emailAddress),console.log("BCC: "+l.emailAddress),console.log("Body:"+a),bannerNullHandler(getBannerFromBody(a),e)}))}function getBannerFromBody(e){var t=e.match(/^(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im);return console.log(t),t?(console.log("banner found"),t[0]):(console.log("banner null"),null)}function bannerNullHandler(e,t){null==e?t.completed({allowEvent:!1,cancelLabel:"Ok",commandId:"msgComposeOpenPaneButton",contextData:JSON.stringify({a:"aValue",b:"bValue"}),errorMessage:"Please enter a banner, banner error detected.",sendModeOverride:Office.MailboxEnums.SendModeOverride.PromptUser}):t.completed({allowEvent:!0})}function hasMatches(e){if(null==e||""==e)return!1;for(var t=["send","picture","document","attachment"],n=0;n<t.length;n++){var r=t[n].trim();if(RegExp(r,"i").test(e))return!0}return!1}function getAttachmentsCallback(e){var t=e.asyncContext;if(e.value.length>0){for(var n=0;n<e.value.length;n++)if(0==e.value[n].isInline)return void t.completed({allowEvent:!0});t.completed({allowEvent:!1,errorMessage:"Looks like the body of your message includes an image or an inline file. Would you like to attach a copy of it to the message?",cancelLabel:"Attach a copy",commandId:"msgComposeOpenPaneButton",sendModeOverride:Office.MailboxEnums.SendModeOverride.PromptUser})}else t.completed({allowEvent:!1,errorMessage:"Looks like you're forgetting to include an attachment blayh blah testing.",cancelLabel:"Add an attachment",commandId:"msgComposeOpenPaneButton"})}function getToRecipientsAsync(){return new Promise((function(e,t){Office.context.mailbox.item.to.getAsync((function(n){n.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'To' recipients. Error: "+JSON.stringify(n.error)),t("Failed to get 'To' recipients.")):e(n.value)}))}))}function getSenderAsync(){return new Promise((function(e,t){Office.context.mailbox.item.from.getAsync((function(n){n.status!==Office.AsyncResultStatus.Succeeded?(console.log("unable to get sender"),t("Failed to get sender. "+JSON.stringify(n.error))):e(n.value)}))}))}function getBodyAsync(){return new Promise((function(e,t){Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text,(function(n){n.status!==Office.AsyncResultStatus.Succeeded?(console.log("unable to get body"),t("Failed to get body. "+JSON.stringify(n.error))):e(n.value)}))}))}function getCCAsync(){return new Promise((function(e,t){Office.context.mailbox.item.cc.getAsync((function(n){n.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'CC' recipients. Error: "+JSON.stringify(n.error)),t("Failed to get 'CC' recipients.")):e(n.value)}))}))}function getBCCAsync(){return new Promise((function(e,t){Office.context.mailbox.item.bcc.getAsync((function(n){n.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'BCC' recipients. Error: "+JSON.stringify(n.error)),t("Failed to get 'BCC' recipients.")):e(n.value)}))}))}Office.context.platform!==Office.PlatformType.PC&&null!=Office.context.platform||Office.actions.associate("onMessageSendHandler",onMessageSendHandler);