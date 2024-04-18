/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// const { getToRecipientsAsync } = require("./promise.js");
// import { getToRecipientsAsync } from "./promise.js";

function onMessageSendHandler(event) {
  Promise.all([getToRecipientsAsync(), getSenderAsync(), getBodyAsync(), getCCAsync(), getBCCAsync()]).then(
    ([to, sender, body, cc, bcc]) => {
      console.log("To recipients:");
      to.forEach((to) => console.log(to.emailAddress));
      console.log("Sender:" + sender.displayName + " " + sender.emailAddress);
      console.log("CC: " + cc.emailAddress);
      console.log("BCC: " + bcc.emailAddress);
      console.log("Body:" + body);

      //DEBUGGING
      // const message =
      //   "Riciepient: " +
      //   to.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ") +
      //   "\nCC recipients: " +
      //   (cc ? cc.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ") : "None") +
      //   "\nBCC recipients: " +
      //   (bcc
      //     ? bcc.map((recipient) => recipient.emailAddress + " (" + recipient.displayName + ")").join(", ")
      //     : "None") +
      //   "\nSender: " +
      //   sender.displayName +
      //   "\nBody: " +
      //   body;
      // console.error(message);
      // event.completed({ allowEvent: false, errorMessage: message });
      // return;

      //   Office.context.mailbox.item.body.getAsync(
      //     "text",
      //     { asyncContext: event },
      //     getBodyCallback
      //   );
      // }

      // function getBodyCallback(asyncResult){
      //   const event = asyncResult.asyncContext;
      //   let body = "";
      //   if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
      //     body = asyncResult.value;
      //   } else {
      //     const message = "Failed to get body text";
      //     console.error(message);
      //     event.completed({ allowEvent: false, errorMessage: message });
      //     return;
      //   }

      const banner = getBannerFromBody(body);
      // Check if the banner is null error
      bannerNullHandler(banner, event);

      const bannerMarkings = parseBannerMarkings(banner);

      // const message = "Banner Markings: " + bannerMarkings;
      // console.error(message);
      // event.completed({ allowEvent: false, errorMessage: message });
      // return;

      if (bannerMarkings.message !== "") {
        errorPopupHandler(bannerMarkings.message, event);
      }

      //CHANGE
      // checkRecipientClassification(toRecipients,bannerMarkings.banner[0])
      // .then((allowEvent) => {
      //   if (!allowEvent) {
      //     // Prevent sending the email
      //     console.log("Prevent sending email");
      //     event.completed({ allowEvent: false });
      //     Office.context.mailbox.item.notificationMessages.addAsync(
      //       "unauthorizedSending",
      //       {
      //         type: Office.MailboxEnums.ItemNotificationMessageType
      //           .ErrorMessage,
      //         message: "You are not authorized to send this email",
      //       },
      //       (result) => {
      //         console.log(result);
      //       }
      //     );
      //   } else {
      //     // Allow sending the email
      //     event.completed({ allowEvent: true });
      //   }
      // })
      // .catch((error) => {
      //   console.error(
      //     "Error occurred while checking recipient classification: " + error
      //   );
      // });
    }
  );
}

function getBannerFromBody(body) {
  const banner_regex = /^(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im;

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

function bannerNullHandler(banner, event) {
  if (banner == null) {
    event.completed({
      allowEvent: false,
      cancelLabel: "Ok",
      commandId: "msgComposeOpenPaneButton",
      contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
      errorMessage: "Please enter a banner, banner error detected.",
      //     //underneath with enable the user to press send anyways, might need later
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
    return;
  }
}

function hasMatches(body) {
  if (body == null || body == "") {
    return false;
  }

  const arrayOfTerms = ["send", "picture", "document", "attachment"];
  for (let index = 0; index < arrayOfTerms.length; index++) {
    const term = arrayOfTerms[index].trim();
    const regex = RegExp(term, "i");
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
      errorMessage:
        "Looks like the body of your message includes an image or an inline file. Would you like to attach a copy of it to the message?",
      cancelLabel: "Attach a copy",
      commandId: "msgComposeOpenPaneButton",
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });
  } else {
    event.completed({
      allowEvent: false,
      errorMessage: "Looks like you're forgetting to include an attachment blayh blah testing.",
      cancelLabel: "Add an attachment",
      commandId: "msgComposeOpenPaneButton",
    });
  }
}

/**
 * Gets the 'to' from email.
 */
function getToRecipientsAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.to.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get 'To' recipients. Error: " + JSON.stringify(result.error));
        reject("Failed to get 'To' recipients.");
      } else {
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'sender' from email.
 */
function getSenderAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.from.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("unable to get sender");
        reject("Failed to get sender. " + JSON.stringify(result.error));
      } else {
        //const msgFrom = result.value;
        //console.log("Message from: " + msgFrom.displayName + " (" + msgFrom.emailAddress + ")");
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'body' from email.
 */
function getBodyAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.log("unable to get body");
        reject("Failed to get body. " + JSON.stringify(result.error));
      } else {
        //console.log("this worked");
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'CC' from email.
 */
function getCCAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.cc.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get 'CC' recipients. Error: " + JSON.stringify(result.error));
        reject("Failed to get 'CC' recipients.");
      } else {
        resolve(result.value);
      }
    });
  });
}

/**
 * Gets the 'BCC' from email.
 */
function getBCCAsync() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.bcc.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error("Failed to get 'BCC' recipients. Error: " + JSON.stringify(result.error));
        reject("Failed to get 'BCC' recipients.");
      } else {
        resolve(result.value);
      }
    });
  });
}

function parseBannerMarkings(banner) {
  // const cat1_regex = "TOP[\s]*SECRET|TS|(TS)|SECRET|S|(S)|CONFIDENTIAL|C|(C)|UNCLASSIFIED|U|(U)";
  // const cat4_regex = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS";
  // const cat7_regex = "ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  // const cat4_and_cat7 = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  const cat1_regex = /TOP\s*SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U/gi;
  const cat4_regex = /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS/gi;
  const cat7_regex =
    /ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
  const cat4_and_cat7 =
    /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;

  const Categories = banner.split("//");
  console.log(Categories);
  let Category_1 = Category(Categories[0], cat1_regex, 1);
  let Category_4 = null;
  let Category_7 = null;
  if (Categories[1]) {
    if (Categories[1].toUpperCase().match(cat7_regex)) {
      // If the second parse matches the regex for category 7, then we need to make category 4 null and run category 7
      console.log("second category matches category 7");
      Category_4 = null;
      Category_7 = Category(Categories[1], cat7_regex, 7);
    } else {
      // If the second parse doesnt match, run each category with its corresponding regex
      console.log("second category doesnt match category 7, running normal program");
      Category_4 = Category(Categories[1], cat4_regex, 4);
      Category_7 = Category(Categories[2], cat7_regex, 7);
    }
  } else {
    console.log("second category returned null");
  }

  const Together = [Category_1, Category_4, Category_7];
  //CHANGE
  let errMsg = checkDisseminations(Category_1, Category_7);
  //add Zach's stuff after testing

  //return Together;
  //CHANGE
  return {
    banner: Together,
    message: errMsg,
  };
}

/**
 * returns the submarkings of the category. if there is one category, then it returns null
 * @param { string } category
 * @returns { array } || null
 */
function getSubMarkings(category) {
  let submarkings = category.split("/");
  if (submarkings.length <= 1) {
    console.log("There is only one submarking");
    return null;
  }
  console.log(submarkings);
  return submarkings;
}

/**
 * function that uses regex to match the input category string, if no match is found it returns null
 * @param { String } category
 * @param { String } regex
 * @param { int } categoryNum
 */
function Category(category, regex, categoryNum) {
  if (!category) {
    console.log("Category " + categoryNum + " string returned null");
    return null;
  } else if (category.toUpperCase().match(regex)) {
    console.log("returning category " + categoryNum);
    console.log(category.toUpperCase());
    return category.toUpperCase();
  }
  console.log("String did not match category " + categoryNum + "'s regex");
  return null;
}

/**
 * given a string it validates that the first marking is classified or not
 * returns a true or false value depending on if its valid or not
 * @param {string} banner
 */
function ValidateClassification(banner) {
  regex = /TS|S|C|U/gi;
  if (banner.match(regex)) {
    return true;
  }
  return false;
}
function validateSCI(classification, sci, dissemination) {
  let valid = 0;
  let msg = "";
  let subBanner = sci.split("/");
  subBanner.ForEach((marking) => {
    /**
     * May be used only with
     * TOP SECRET, SECRET, or CONFIDENTIAL. NOFORN is required.
     *
     */
    if (marking.match(/HCS/gi)) {
      if (classification.includes("U") || classification.includes("UNCLASSIFIED")) {
        valid = 1;
        msg += "CANNOT USE HCS with UNCLASSIFIED. ";
      }

      if (dissemination.includes("NOFORN") || dissemination.includes("NOT RELEASABLE TO FOREIGN NATIONALS")) {
      } else {
        valid = 1;
        msg += "HCS MUST USE NOFORN. ";
      }
    }

    /**
     * May be used only with
     * TOP SECRET, SECRET, or CONFIDENTIAL.
     *
     */
    if (marking.match(/SI/gi)) {
      if (classification.includes("U") || classification.includes("UNCLASSIFIED")) {
        valid = 1;
        msg += "CANNOT USE SI with UNCLASSIFIED. ";
      }
    }

    /**
     * May be used only with
     * TOP SECRET. Requires SI and ORCON
     *
     */
    if (marking.match(/-G/gi)) {
      if (!classification.includes("TS")) {
        valid = 1;
        msg += "CANNOT USE -G with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      } else if (!classification.includes("TOP SECRET")) {
        valid = 1;
        msg += "CANNOT USE -G with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      }

      if (!sci.includes("SI")) {
        valid = 1;
        msg += "MUST USE -G with SI. ";
      } else if (!sci.includes("COMINT")) {
        valid = 1;
        msg += "MUST USE -G with SI. ";
      }

      if (!sci.includes("ORCON")) {
        valid = 1;
        msg += "MUST USE -G with ORCON. ";
      } else if (!sci.includes("ORIGINATOR CONTROLLED")) {
        valid = 1;
        msg += "MUST USE -G with ORCON. ";
      }
    }

    /**
     * May be used only with
     * TOP SECRET. Requires SI
     *
     */
    if (marking.match(/-ECI/gi)) {
      if (!classification.includes("TS")) {
        valid = 1;
        msg += "CANNOT USE -ECI with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      } else if (!classification.includes("TOP SECRET")) {
        valid = 1;
        msg += "CANNOT USE -ECI with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      }

      if (!sci.includes("SI")) {
        valid = 1;
        msg += "MUST USE -ECI with SI. ";
      } else if (!sci.includes("COMINT")) {
        valid = 1;
        msg += "MUST USE -ECI with SI. ";
      }
    }

    /**
     * May be used only with
     * TOP SECRET or SECRET. May require RSEN for imagery product
     *
     */
    if (marking.match(/TK/gi)) {
      if (!classification.includes("TS")) {
        valid = 1;
        msg += "CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. ";
      } else if (!classification.includes("TOP SECRET")) {
        valid = 1;
        msg += "CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. ";
      } else {
        if (!classification.includes("S")) {
          valid = 1;
          msg += "CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. ";
        } else if (!classification.includes("SECRET")) {
          valid = 1;
          msg += "CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. ";
        }
      }
    }
  });

  return [valid, msg];
}

/**
 * @param {String} classification
 * @param {String} dissemination
 */
function checkDisseminations(classification, dissemination) {
  console.log("CLASSIFICATION: " + classification + "\n");
  console.log("DISSEM: " + dissemination + "\n");

  let errorMsg = "";

  let dissParts = dissemination.split("/");

  let dissPartsArray = [];

  for (let i = 0; i < dissParts.length; i++) {
    dissPartsArray.push(dissParts[i]);
  }

  let NOFORNEncountered = false;
  let EYESONLYEncountered = false;
  let RELIDOEncountered = false;
  let RELTOEncountered = false;

  //check disseminations
  for (let i = 0; i < dissPartsArray.length; i++) {
    //FOR OFFICIAL USE ONLY (FOUO): cannot be used with classified information.
    if (dissPartsArray[i] === "FOUO" && classification !== "UNCLASSIFIED") {
      errorMsg = "Cannot use FOUO with classified information.";
    }

    //ORIGINATOR CONTROLLED (ORCON): can only be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    if (dissPartsArray[i] === "ORCON" && classification === "UNCLASSIFIED") {
      errorMsg = "Cannot use ORCON with unclassified information.";
    }

    //CONTROLLED IMAGERY (IMCON): can only be used with SECRET. May require NOFORN.
    if (dissPartsArray[i] === "IMCON" && classification !== "SECRET") {
      errorMsg = "IMCON can only be used with SECRET information.";
    }

    //SOURCES AND METHODS (SAMI): can only be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    //Can be used with REL TO or RELIDO.
    if (dissPartsArray[i] === "SAMI" && classification === "UNCLASSIFIED") {
      errorMsg = "Cannot use SAMI with unclassified information.";
    }

    /** BIG CODE CHUNK THAT HANDLES NOFORN, REL TO, RELIDO, EYES ONLY **/
    //NOT RELEASABLE TO FOREIGN NATIONALS (NOFORN): can only be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    //Cannot be used with REL TO, RELIDO, or EYES ONLY.
    if (dissPartsArray[i] === "NOFORN") {
      NOFORNEncountered = true;
      if (classification === "UNCLASSIFIED") {
        errorMsg = "Cannot use NOFORN with unclassified information.";
      }
    }
    //EYES ONLY: can only be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    //Cannot be used with NOFORN or REL TO. Can be used wth RELIDO.
    else if (dissPartsArray[i].includes("EYES ONLY")) {
      if (dissPartsArray[i].match(/[A-Z]{3}\sEYES ONLY/g)) {
        EYESONLYEncountered = true;
      } else {
        errorMsg = "Wrong formatting of EYES ONLY.";
      }
      if (classification === "UNCLASSIFIED") {
        errorMsg = "EYES ONLY cannot be used with unclassified information.";
      }
    }
    //RELEASABLE BY INFORMATION DISCLOSURE OFFICIAL (RELIDO): may be used independently or with REL TO.
    //Cannot be used with NOFORN.
    else if (dissPartsArray[i] === "RELIDO") {
      RELIDOEncountered = true;
    }
    //AUTHORIZED FOR RELEASE TO (REL TO): can only be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    //May be used with RELIDO. Cannot be used with NOFORN or EYES ONLY.
    else if (dissPartsArray[i].includes("REL TO")) {
      if (dissPartsArray[i].match(/REL TO\s[A-Z]{3}/g)) {
        RELTOEncountered = true;
      } else {
        errorMsg = "Wrong formatting of REL TO.";
      }
      if (classification === "UNCLASSIFIED") {
        errorMsg = "Cannot use REL TO with unclassified information.";
      }
    }

    if (NOFORNEncountered && dissPartsArray[i] === "EYES ONLY") {
      errorMsg = "NOFORN cannot be used with EYES ONLY.";
    } else if (EYESONLYEncountered && dissPartsArray[i] === "NOFORN") {
      errorMsg = "EYES ONLY cannot be used with NOFORN.";
    } else if (NOFORNEncountered && dissPartsArray[i] === "RELIDO") {
      errorMsg = "NOFORN cannot be used with RELIDO.";
    } else if (RELIDOEncountered && dissPartsArray[i] === "NOFORN") {
      errorMsg = "RELIDO cannot be used with NOFORN.";
    } else if (NOFORNEncountered && dissPartsArray[i].includes("REL TO")) {
      errorMsg = "NOFORN cannot be used with REL TO.";
    } else if (RELTOEncountered && dissPartsArray[i] === "NOFORN") {
      errorMsg = "REL TO cannot be used with NOFORN.";
    } else if (EYESONLYEncountered && dissPartsArray[i].includes("REL TO")) {
      errorMsg = "EYES ONLY cannot be used with REL TO.";
    } else if (RELTOEncountered && dissPartsArray === "EYES ONLY") {
      errorMsg = "REL TO cannot be used with EYES ONLY.";
    }

    //CAUTION PROPRIETARY INFORMATION INVOLVED (PROPIN): can be used with all classifications.
    //No error checking needed because there are no restrictions.

    /** BIG CODE CHUNK FOR RD/FRD AND CNDWI/SG **/
    //RESTRICTED DATA (RD): can be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    //FORMERLY RESTRICTED DATA (RD): can be used with TOP SECRET, SECRET, or CONFIDENTIAL.
    if (dissPartsArray[i].includes("RD") || dissPartsArray[i].includes("FRD")) {
      if (dissPartsArray[i] === "RD" || dissPartsArray[i] === "FRD") {
        if (classification === "UNCLASSIFIED") {
          errorMsg = "Cannot use RD or FRD with unclassified information.";
        }

        //-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION (-CNWDI): can be used with TOP SECRET or SECRET.
        //Requires RD or FRD.
      } else if (dissPartsArray[i].match(/(RD|FRD)-CNWDI/g)) {
        if (classification === "CONFIDENTIAL" || classification === "UNCLASSIFIED") {
          errorMsg = "-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED.";
        }

        //-SIGMA[#] (-SG[#]): may be used with TOP SECRET, SECRET, or CONDFIDENTIAL.
        //Requires RD or FRD. [#] represents the SIGMA number, ranges from 1-99.
      } else if (dissPartsArray[i].match(/(RD|FRD)-SG\[(?:[1-9]|[1-9][0-9]|99)\]/g)) {
        if (classification === "UNCLASSIFIED") {
          errorMsg = "-SG cannot be used with UNCLASSIFIED information.";
        }
      } else {
        errorMsg = "Wrong format of banner of RD and FRD.";
      }
    } else if (dissPartsArray[i].includes("-CNWDI")) {
      if (dissPartsArray[i].match(/(RD|FRD)-CNWDI/g)) {
        if (classification === "CONFIDENTIAL" || classification === "UNCLASSIFIED") {
          errorMsg = "-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED.";
        }
      } else {
        errorMsg = "RD or FRD is required for -CNWDI.";
      }
    } else if (dissPartsArray[i].includes("-SG")) {
      if (dissPartsArray[i].match(/(RD|FRD)-SG\[(?:[1-9]|[1-9][0-9]|99)\]/g)) {
        if (classification === "UNCLASSIFIED") {
          errorMsg = "-SG cannot be used with UNCLASSIFIED information.";
        }
      } else {
        errorMsg = "RD or FRD is required for -SG[#].";
      }
    }

    //DOD or DOE CONTROLLED NUCLEAR INFORMATION (DOD UCNI or DOE UCNI): can only be used with UNCLASSIFIED.
    if (dissPartsArray[i] === "DOD UCNI" || dissPartsArray[i] === "DOE UCNI") {
      if (classification !== "UNCLASSIFIED") {
        errorMsg = "DOD/DOE UCNI can only be used with unclassified information.";
      }
    }

    //DEA SENSITIVE (DSEN): can only be used with unclassified.
    if (dissPartsArray[i] === "DSEN" && classification !== "UNCLASSIFIED") {
      errorMsg = "DSEN can only be used with unclassified.";
    }

    //FOREIGN INTELLIGENCE SURVEILLANCE ACT (FISA): does not have any restrictions.
    //No error checking needed.

    // console.log(errorMsg);
  }

  //CHANGE
  return errorMsg;
}

//CHANGE:
function errorPopupHandler(errorMsg, event) {
  event.completed({
    allowEvent: false,
    cancelLabel: "Ok",
    commandId: "msgComposeOpenPaneButton",
    contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
    errorMessage: errorMsg,
    //sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser
  });
  return;
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null) {
  Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
}
