function _slicedToArray(e,n){return _arrayWithHoles(e)||_iterableToArrayLimit(e,n)||_unsupportedIterableToArray(e,n)||_nonIterableRest()}function _nonIterableRest(){throw new TypeError("Invalid attempt to destructure non-iterable instance.\nIn order to be iterable, non-array objects must have a [Symbol.iterator]() method.")}function _unsupportedIterableToArray(e,n){if(e){if("string"==typeof e)return _arrayLikeToArray(e,n);var t=Object.prototype.toString.call(e).slice(8,-1);return"Object"===t&&e.constructor&&(t=e.constructor.name),"Map"===t||"Set"===t?Array.from(e):"Arguments"===t||/^(?:Ui|I)nt(?:8|16|32)(?:Clamped)?Array$/.test(t)?_arrayLikeToArray(e,n):void 0}}function _arrayLikeToArray(e,n){(null==n||n>e.length)&&(n=e.length);for(var t=0,o=new Array(n);t<n;t++)o[t]=e[t];return o}function _iterableToArrayLimit(e,n){var t=null==e?null:"undefined"!=typeof Symbol&&e[Symbol.iterator]||e["@@iterator"];if(null!=t){var o,r,i,a,s=[],c=!0,l=!1;try{if(i=(t=t.call(e)).next,0===n){if(Object(t)!==t)return;c=!1}else for(;!(c=(o=i.call(t)).done)&&(s.push(o.value),s.length!==n);c=!0);}catch(e){l=!0,r=e}finally{try{if(!c&&null!=t.return&&(a=t.return(),Object(a)!==a))return}finally{if(l)throw r}}return s}}function _arrayWithHoles(e){if(Array.isArray(e))return e}function onMessageSendHandler(e){Promise.all([getToRecipientsAsync(),getSenderAsync(),getBodyAsync(),getCCAsync(),getBCCAsync()]).then((function(n){var t=_slicedToArray(n,5),o=t[0],r=t[1],i=t[2],a=t[3],s=t[4];console.log("To recipients:"),o.forEach((function(e){return console.log(e.emailAddress)})),console.log("Sender:"+r.displayName+" "+r.emailAddress),console.log("CC: "+a.emailAddress),console.log("BCC: "+s.emailAddress),console.log("Body:"+i);var c=getBannerFromBody(i);bannerNullHandler(c,e);var l=parseBannerMarkings(c);""!==l.message&&errorPopupHandler(l.message,e)}))}function getBannerFromBody(e){var n=e.match(/^(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im);return console.log(n),n?(console.log("banner found"),n[0]):(console.log("banner null"),null)}function bannerNullHandler(e,n){null!=e||n.completed({allowEvent:!1,cancelLabel:"Ok",commandId:"msgComposeOpenPaneButton",contextData:JSON.stringify({a:"aValue",b:"bValue"}),errorMessage:"Please enter a banner, banner error detected.",sendModeOverride:Office.MailboxEnums.SendModeOverride.PromptUser})}function hasMatches(e){if(null==e||""==e)return!1;for(var n=["send","picture","document","attachment"],t=0;t<n.length;t++){var o=n[t].trim();if(RegExp(o,"i").test(e))return!0}return!1}function getAttachmentsCallback(e){var n=e.asyncContext;if(e.value.length>0){for(var t=0;t<e.value.length;t++)if(0==e.value[t].isInline)return void n.completed({allowEvent:!0});n.completed({allowEvent:!1,errorMessage:"Looks like the body of your message includes an image or an inline file. Would you like to attach a copy of it to the message?",cancelLabel:"Attach a copy",commandId:"msgComposeOpenPaneButton",sendModeOverride:Office.MailboxEnums.SendModeOverride.PromptUser})}else n.completed({allowEvent:!1,errorMessage:"Looks like you're forgetting to include an attachment blayh blah testing.",cancelLabel:"Add an attachment",commandId:"msgComposeOpenPaneButton"})}function getToRecipientsAsync(){return new Promise((function(e,n){Office.context.mailbox.item.to.getAsync((function(t){t.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'To' recipients. Error: "+JSON.stringify(t.error)),n("Failed to get 'To' recipients.")):e(t.value)}))}))}function getSenderAsync(){return new Promise((function(e,n){Office.context.mailbox.item.from.getAsync((function(t){t.status!==Office.AsyncResultStatus.Succeeded?(console.log("unable to get sender"),n("Failed to get sender. "+JSON.stringify(t.error))):e(t.value)}))}))}function getBodyAsync(){return new Promise((function(e,n){Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text,(function(t){t.status!==Office.AsyncResultStatus.Succeeded?(console.log("unable to get body"),n("Failed to get body. "+JSON.stringify(t.error))):e(t.value)}))}))}function getCCAsync(){return new Promise((function(e,n){Office.context.mailbox.item.cc.getAsync((function(t){t.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'CC' recipients. Error: "+JSON.stringify(t.error)),n("Failed to get 'CC' recipients.")):e(t.value)}))}))}function getBCCAsync(){return new Promise((function(e,n){Office.context.mailbox.item.bcc.getAsync((function(t){t.status!==Office.AsyncResultStatus.Succeeded?(console.error("Failed to get 'BCC' recipients. Error: "+JSON.stringify(t.error)),n("Failed to get 'BCC' recipients.")):e(t.value)}))}))}function parseBannerMarkings(e){var n=/ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi,t=e.split("//");console.log(t);var o=Category(t[0],/TOP\s*SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U/gi,1),r=null,i=null;return t[1]?t[1].toUpperCase().match(n)?(console.log("second category matches category 7"),r=null,i=Category(t[1],n,7)):(console.log("second category doesnt match category 7, running normal program"),r=Category(t[1],/COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS/gi,4),i=Category(t[2],n,7)):console.log("second category returned null"),{banner:[o,r,i],message:checkDisseminations(o,i)}}function getSubMarkings(e){var n=e.split("/");return n.length<=1?(console.log("There is only one submarking"),null):(console.log(n),n)}function Category(e,n,t){return e?e.toUpperCase().match(n)?(console.log("returning category "+t),console.log(e.toUpperCase()),e.toUpperCase()):(console.log("String did not match category "+t+"'s regex"),null):(console.log("Category "+t+" string returned null"),null)}function ValidateClassification(e){return regex=/TS|S|C|U/gi,!!e.match(regex)}function validateSCI(e,n,t){var o=0,r="";return n.split("/").ForEach((function(i){i.match(/HCS/gi)&&((e.includes("U")||e.includes("UNCLASSIFIED"))&&(o=1,r+="CANNOT USE HCS with UNCLASSIFIED. "),t.includes("NOFORN")||t.includes("NOT RELEASABLE TO FOREIGN NATIONALS")||(o=1,r+="HCS MUST USE NOFORN. ")),i.match(/SI/gi)&&(e.includes("U")||e.includes("UNCLASSIFIED"))&&(o=1,r+="CANNOT USE SI with UNCLASSIFIED. "),i.match(/-G/gi)&&(e.includes("TS")&&e.includes("TOP SECRET")||(o=1,r+="CANNOT USE -G with UNCLASSIFIED, CONFIDENTIAL, or SECRET. "),n.includes("SI")&&n.includes("COMINT")||(o=1,r+="MUST USE -G with SI. "),n.includes("ORCON")&&n.includes("ORIGINATOR CONTROLLED")||(o=1,r+="MUST USE -G with ORCON. ")),i.match(/-ECI/gi)&&(e.includes("TS")&&e.includes("TOP SECRET")||(o=1,r+="CANNOT USE -ECI with UNCLASSIFIED, CONFIDENTIAL, or SECRET. "),n.includes("SI")&&n.includes("COMINT")||(o=1,r+="MUST USE -ECI with SI. ")),i.match(/TK/gi)&&(e.includes("TS")&&e.includes("TOP SECRET")&&e.includes("S")&&e.includes("SECRET")||(o=1,r+="CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. "))})),[o,r]}function checkDisseminations(e,n){console.log("CLASSIFICATION: "+e+"\n"),console.log("DISSEM: "+n+"\n");for(var t="",o=n.split("/"),r=[],i=0;i<o.length;i++)r.push(o[i]);for(var a=!1,s=!1,c=!1,l=!1,u=0;u<r.length;u++)"FOUO"===r[u]&&"UNCLASSIFIED"!==e&&(t="Cannot use FOUO with classified information."),"ORCON"===r[u]&&"UNCLASSIFIED"===e&&(t="Cannot use ORCON with unclassified information."),"IMCON"===r[u]&&"SECRET"!==e&&(t="IMCON can only be used with SECRET information."),"SAMI"===r[u]&&"UNCLASSIFIED"===e&&(t="Cannot use SAMI with unclassified information."),"NOFORN"===r[u]?(a=!0,"UNCLASSIFIED"===e&&(t="Cannot use NOFORN with unclassified information.")):r[u].includes("EYES ONLY")?(r[u].match(/[A-Z]{3}\sEYES ONLY/g)?s=!0:t="Wrong formatting of EYES ONLY.","UNCLASSIFIED"===e&&(t="EYES ONLY cannot be used with unclassified information.")):"RELIDO"===r[u]?c=!0:r[u].includes("REL TO")&&(r[u].match(/REL TO\s[A-Z]{3}/g)?l=!0:t="Wrong formatting of REL TO.","UNCLASSIFIED"===e&&(t="Cannot use REL TO with unclassified information.")),a&&"EYES ONLY"===r[u]?t="NOFORN cannot be used with EYES ONLY.":s&&"NOFORN"===r[u]?t="EYES ONLY cannot be used with NOFORN.":a&&"RELIDO"===r[u]?t="NOFORN cannot be used with RELIDO.":c&&"NOFORN"===r[u]?t="RELIDO cannot be used with NOFORN.":a&&r[u].includes("REL TO")?t="NOFORN cannot be used with REL TO.":l&&"NOFORN"===r[u]?t="REL TO cannot be used with NOFORN.":s&&r[u].includes("REL TO")?t="EYES ONLY cannot be used with REL TO.":l&&"EYES ONLY"===r&&(t="REL TO cannot be used with EYES ONLY."),r[u].includes("RD")||r[u].includes("FRD")?"RD"===r[u]||"FRD"===r[u]?"UNCLASSIFIED"===e&&(t="Cannot use RD or FRD with unclassified information."):r[u].match(/(RD|FRD)-CNWDI/g)?"CONFIDENTIAL"!==e&&"UNCLASSIFIED"!==e||(t="-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED."):r[u].match(/(RD|FRD)-SG\[(?:[1-9]|[1-9][0-9]|99)\]/g)?"UNCLASSIFIED"===e&&(t="-SG cannot be used with UNCLASSIFIED information."):t="Wrong format of banner of RD and FRD.":r[u].includes("-CNWDI")?r[u].match(/(RD|FRD)-CNWDI/g)?"CONFIDENTIAL"!==e&&"UNCLASSIFIED"!==e||(t="-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED."):t="RD or FRD is required for -CNWDI.":r[u].includes("-SG")&&(r[u].match(/(RD|FRD)-SG\[(?:[1-9]|[1-9][0-9]|99)\]/g)?"UNCLASSIFIED"===e&&(t="-SG cannot be used with UNCLASSIFIED information."):t="RD or FRD is required for -SG[#]."),"DOD UCNI"!==r[u]&&"DOE UCNI"!==r[u]||"UNCLASSIFIED"!==e&&(t="DOD/DOE UCNI can only be used with unclassified information."),"DSEN"===r[u]&&"UNCLASSIFIED"!==e&&(t="DSEN can only be used with unclassified.");return t}function errorPopupHandler(e,n){n.completed({allowEvent:!1,cancelLabel:"Ok",commandId:"msgComposeOpenPaneButton",contextData:JSON.stringify({a:"aValue",b:"bValue"}),errorMessage:e})}Office.context.platform!==Office.PlatformType.PC&&null!=Office.context.platform||Office.actions.associate("onMessageSendHandler",onMessageSendHandler);