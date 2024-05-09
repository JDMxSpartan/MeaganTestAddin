/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const CSVdata = `Email,Authorization,Agency,Country
zachary.sierra@my.utsa.edu,TOP SECRET,NSA,USA
ria.james@my.utsa.edu,TOP SECRET,NSA,USA
eric.lopardi@my.utsa.edu,TOP SECRET,NSA,USA
meagan.mueller@my.utsa.edu,TOP SECRET,NSA,USA
kevin.dehoyos@my.utsa.edu,TOP SECRET,NSA,USA
madison.mueller@my.utsa.edu,TOP SECRET,NSA,USA
jonathan.gregerson@my.utsa.edu,SECRET,NSA,USA
john.doe@my.utsa.edu,TOP SECRET,NSA,NZL
jane.doe@my.utsa.edu,SECRET,FBI,NZL
alex.smith@my.utsa.edu,CONFIDENTIAL,NSA,AUS
emily.johnson@my.utsa.edu,UNCLASSIFIED,FBI,USA
michael.wang@my.utsa.edu,SECRET,NSA,USA
sarah.khan@my.utsa.edu,CONFIDENTIAL,FBI,AUS
chris.davis@my.utsa.edu,TOP SECRET,NSA,USA
lisa.nguyen@my.utsa.edu,UNCLASSIFIED,FBI,NZL
ryan.chen@my.utsa.edu,SECRET,NSA,USA
jessica.wilson@my.utsa.edu,CONFIDENTIAL,FBI,AUS
adam.thomas@my.utsa.edu,UNCLASSIFIED,NSA,NZL
amanda.jackson@my.utsa.edu,TOP SECRET,FBI,USA
kevin.murphy@my.utsa.edu,SECRET,NSA,USA
olivia.white@my.utsa.edu,CONFIDENTIAL,FBI,USA
daniel.lee@my.utsa.edu,UNCLASSIFIED,NSA,AUS
grace.martin@my.utsa.edu,TOP SECRET,FBI,NZL
samuel.rodriguez@my.utsa.edu,SECRET,NSA,USA
natalie.anderson@my.utsa.edu,CONFIDENTIAL,FBI,AUS
eric.garcia@my.utsa.edu,UNCLASSIFIED,NSA,NZL`;

function parseCSV(csvData) {
  const rows = csvData.split('\n');
  const headers = rows[0].split(',');
  const data = [];
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i].split(',');
    const rowData = {};
    for (let j = 0; j < headers.length; j++) {
      rowData[headers[j]] = row[j];
    }
    data.push(rowData);
  }
  return data;
}

function onMessageSendHandler(event) {
  Promise.all([getToRecipientsAsync(), getSenderAsync(), getBodyAsync(), getCCAsync(), getBCCAsync()]).then(
    ([to, sender, body, cc, bcc]) => {
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

      const banner = getBannerFromBody(body);
      bannerNullHandler(banner, event);
      bannerMatchHandler(banner, event);
      const bannerMarkings = parseBannerMarkings(banner);

      //DEBUGGING
      // const message = "Banner Markings: " + bannerMarkings;
      // event.completed({ allowEvent: false, errorMessage: message });
      // return;

      if (bannerMarkings.message !== "") {
        errorPopupHandler(bannerMarkings.message, event);
      }

      //message = "reach here, banner mark is" + bannerMarkings.banner[0];
      //errorPopupHandler(message, event);
      const csvData = parseCSV(CSVdata);

      // try {
      //   const CSVdata = parseCSV(csvData);
      //   const headers = Object.keys(CSVdata[0]); // Get the headers
      //   const firstRow = CSVdata[0]; // Get the first row of data
      //   message = "reach here, csv is " + CSVdata + headers + firstRow;
      //   errorPopupHandler(message, event);
      // } catch (error) {
      //   message = "reach here, csv error is " + error;
      //   errorPopupHandler(message, event);
      // }

      Promise.all([
        //checkRecipientClassification(sender, 'sender', bannerMarkings.banner[0]),
        checkRecipientClassification(to, "to", bannerMarkings.banner[0], event, csvData),
        checkRecipientClassification(cc, "CC", bannerMarkings.banner[0], event, csvData),
        checkRecipientClassification(bcc, "BCC", bannerMarkings.banner[0], event, csvData),
      ]).then(([recipientCheck, ccCheck, bccCheck]) => {
        let authChecksPassed = true;
        let countryChecksPassed = false;
        let hasCheckedCountry = false;

        for (const [authorized, email] of recipientCheck) {
          if (!authorized) {
            message = `Recipient ${email} is NOT AUTHORIZED to view this email`;
            authChecksPassed = false;
            errorPopupHandler(message, event);
          }
        }
  
        for (const [authorized, email] of ccCheck) {
          if (!authorized) {
            message = `CC'd user ${email} is NOT AUTHORIZED to view this email`;
            authChecksPassed = false;
            errorPopupHandler(message, event);
          }
        }
  
        for (const [authorized, email] of bccCheck) {
          if (!authorized) {
            console.log(
              `BCC'd user ${email} is NOT AUTHORIZED to view this email`
            );
            authChecksPassed = false;
            errorPopupHandler(message, event);
          }
        }
      
      dissemination = bannerMarkings.banner[2];
      if (dissemination != null) {
        let dissParts = dissemination.split("/");
        let dissPartsArray = [];
        let hasNOFORN = false;
        for (let i = 0; i < dissParts.length; i++) {
          dissPartsArray.push(dissParts[i]);
          if (dissPartsArray[i] === "NOFORN") {
            hasNOFORN = true;
          }
        }

        if (hasNOFORN === true) {
          //if dissPartsArray does not contain noforn at all then has checked is true
          for (let i = 0; i < dissPartsArray.length; i++) {
            //what other banenrs do we need to check
            //REL TO with the country authorized
            //
            if (dissPartsArray[i] === "NOFORN") {
              //NOFORNEncountered = true;
              Promise.all([
                checkCountryForRecipients("to", to, csvData),
                checkCountryForRecipients("CC", cc, csvData),
                checkCountryForRecipients("BCC", bcc, csvData),
              ]).then(([recipientCheck, ccCheck, bccCheck]) => {
                console.log("LOGGING NOFORN AUTHORIZATION:");
                console.log("To check: " + recipientCheck);
                console.log("CC check: " + ccCheck);
                console.log("BCC check: " + bccCheck);
                let message = "";

                for (const [authorized, email] of recipientCheck) {
                  if (!authorized) {
                    message = `Recipient(s) ${email} is a Foreign National and NOT AUTHORIZED to view this email`;
                    authChecksPassed = false;
                    errorPopupHandler(message, event);
                  }
                }

                for (const [authorized, email] of ccCheck) {
                  if (!authorized) {
                    message = `CC'd user(s) ${email} is a Foreign National and NOT AUTHORIZED to view this email`;
                    authChecksPassed = false;
                    errorPopupHandler(message, event);
                  }
                }

                for (const [authorized, email] of bccCheck) {
                  if (!authorized) {
                    console.log(
                      `BCC'd use(s)r ${email} is a Foreign National and NOT AUTHORIZED to view this email`
                    );
                    authChecksPassed = false;
                    errorPopupHandler(message, event);
                  }
                }

                console.log(
                  "Authorization checks passed is: " + authChecksPassed
                );
                console.log("Country checks passed is: " + countryChecksPassed);
                if (countryChecksPassed && authChecksPassed) {
                  event.completed({
                    allowEvent: true,
                  });
                }
              });
            }
          }
        } else {
          //should maybe add something here after gone through loop and didnt find noforn
          console.log(" dissemination is not noforn");
          hasCheckedCountry = true;
        }
      } else {
        console.log(" no disseminations found so dont need to check country");
        hasCheckedCountry = true;
      }

      //AUTH CHECKS PASSED THEN ALLOW EVENT ////////////////////////////////////////////
      console.log("Authorization checks passed is: " + authChecksPassed);
      console.log("hasCheckedCountry is: " + hasCheckedCountry);

      if (authChecksPassed && hasCheckedCountry) {
        event.completed({
          allowEvent: true,
        });
      }
    });
    //./Promise 1
  });
  //./Promise 2
}


/**
 * CHECKS THE CLASSIFICATION LEVEL OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @param {String} documentClassification The classification level of the email
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 * above needs to return not only a boolean but email adress too
 */
function checkRecipientClassification(
  recipients,
  recipientType,
  documentClassification,
  event,
  csvData
) {
  // console.log(`Checking ${recipientType} recipients classification`);

  
  return Promise.all(
    recipients.map((recipient) => {
      const emailAddress = recipient.emailAddress;
      // console.log(`${recipientType} Email Address: ${emailAddress}`);
      if (!emailAddress) {
        // console.log("No recipients for: " + recipientType + " type returned " + recipients.emailAddress);
        return [true, null];
      }
      return userMeetsSecurityClearance(
        csvData,
        documentClassification,
        emailAddress,
        event
      )
        .then((isClearance) => {
          if (!isClearance) {
            //console.log(`${emailAddress} is NOT AUTHORIZED to view this email`);
            return [false, emailAddress];
          } else {
            //console.log(`${recipientType} is cleared`);
            return [true, emailAddress];
          }
        })
        .catch((error) => {
          // console.error(
          //   `Error while checking ${recipientType} clearance: `,
          //   error
          // );
          return [false, emailAddress];
        });
    })
  ).then((results) => {
    return results; // Return authorization status and email address for each recipient
  });
}


async function userMeetsSecurityClearance(csvData, documentClassification, email, event) {
  return new Promise((resolve, reject) => {
    // message = "reach here, csv is " + csvData;
    // errorPopupHandler(message, event);
    let accessGranted = false;

    for (const row of csvData) {
      message = "reach here, csv row is " + row + JSON.stringify(row) + row["Email"];
      errorPopupHandler(message, event);
      if (row["Email"] === email.toLowerCase()) {
        const userClearance = row["Authorization"];
        if (canUserAccess(documentClassification, userClearance)) {
          accessGranted = true;
          resolve(accessGranted);
          return;
        }
      }
    }

    resolve(accessGranted);
  });
}


function canUserAccess(documentClassification, userClearance) {
  console.log("canUserAccess Function")
  const levels = ['confidential', 'secret', 'top secret'];
  const documentIndex = levels.indexOf(documentClassification.trim().toLowerCase());
  const userIndex = levels.indexOf(userClearance.trim().toLowerCase());

  return userIndex >= documentIndex;
}

/**
 * CHECKS THE NOFORN STATUS OF A TO, CCs, OR BBCs.
 * @param {array} recipients An array of recipients, CCs, or BCCs
 * @param {String} recipientType The type of recipient ('to', 'cc', or 'bcc')
 * @returns {Promise<boolean>} Returns a promise resolving to true if all recipients are permitted to view the email
 */
function checkCountryForRecipients(recipientType, recipients, csvData) {
  // console.log(`Checking ${recipientType} country`);

  return Promise.all(
    recipients.map((recipient) => {
      const emailAddress = recipient.emailAddress;
      // console.log(`${recipientType} Email Address: ${emailAddress}`);
      return check_NOFORN_Access(csvData, emailAddress)
        .then((isNOFORN) => {
          // console.log(`isNOFORN for ${recipientType} ${emailAddress} returned: ${isNOFORN}`);
          if (!isNOFORN) {
            console.log(
              `${emailAddress} is a Foreign National and NOT AUTHORIZED to view this email`
            );
            return [false, emailAddress];
          } else {
            console.log(`${recipientType} user is Cleared as USA`);
            return [true, emailAddress];
          }
        })
        .catch((error) => {
          console.error(
            `Error while checking isNOFORN for ${recipientType} ${emailAddress}: `,
            error
          );
          return [false, emailAddress];
        });
    })
  ).then((results) => {
    return results; // Return true if all recipients are cleared
  });
}

async function check_NOFORN_Access(csvData, email) {
  return new Promise((resolve, reject) => {
    // message = "reach here, csv is " + csvData;
    // errorPopupHandler(message, event);
    let accessGranted = false;

    for (const row of csvData) {
      message = "reach here, csv row is " + row + JSON.stringify(row) + row["Email"];
      errorPopupHandler(message, event);
      if (row["Email"] === email.toLowerCase()) {
        const userCountry = row["Country"];
        if (userCountry == "USA") {
          accessGranted = true;
          console.log("AccessGranted = true");
          resolve(accessGranted); 
          return; 
        }
      }
    }

    resolve(accessGranted);
  });
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

function bannerMatchHandler(banner, event) {
  if (banner == "DONT MATCH") {
    console.log("banners don't match errMsgUI");

    //type can either be errorMessage or informationalMessage
    // mailboxItem.notificationMessages.addAsync("NoSend", {
    // type: "errorMessage",
    // message: "Please enter a banner marking for this email.",
    // });

    //event.completed({ allowEvent: false });

    event.completed({
      allowEvent: false,
      cancelLabel: "Ok",
      commandId: "msgComposeOpenPaneButton",
      contextData: JSON.stringify({ a: "aValue", b: "bValue" }),
      errorMessage: "Please match banners.",
      //     //underneath with enable the user to press send anyways, might need later
      sendModeOverride: Office.MailboxEnums.SendModeOverride.PromptUser,
    });

    return;
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

// this file contains the banner etraction logic from the body as
// well as parsing the banner and retrieving the categories

/**
 * function to extract banner from message body
 * parameter is the message body contents
 * returns the banner from the body
 * @param { String } body
 */

function getBannerFromBody(body) {
  const classification_regex = /^(CLASSIFICATION:)/im;

  // Find the index of the first occurrence of classification banner
  const firstOccurrenceIndex = body.search(classification_regex);

  // Remove the classification line from the first banner
  const banner1WithoutClassification1 = body.slice(firstOccurrenceIndex + 15); // Adding 15 to skip the "CLASSIFICATION:" part
  const banner1WithoutClassification2 = banner1WithoutClassification1.replace(
    classification_regex,
    ""
  );

  // Find the index of the second occurrence, starting the search immediately after the first occurrence
  const secondOccurrenceIndex = body.indexOf(
    "CLASSIFICATION:",
    firstOccurrenceIndex + 1
  );

  // Remove the classification line from the second banner
  const banner2WithoutClassification1 = body.slice(secondOccurrenceIndex + 15); // Adding 15 to skip the "CLASSIFICATION:" part
  const banner2WithoutClassification2 = banner2WithoutClassification1.replace(
    classification_regex,
    ""
  );

  const banner_regex =
    /^\s*(TOP *SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U)((\/\/)?(.*)?(\/\/)((.*)*))?/im;
  console.log("Searching for Banner...");
  const banner1 = banner1WithoutClassification2.match(banner_regex);

  console.log("Searching for Banner 2...");
  const banner2 = banner2WithoutClassification2.match(banner_regex);

  // Check if both banners are found
  if (banner1 && banner2) {
    console.log("Banners Found!");
    console.log([banner1[0], banner2[0]]);
    if (banner2[0] == banner1[0]) {
      return banner1[0];
    } else if (banner2[0] != banner1[0]) {
      console.log("BANNER EXTRACT: Banners Don't match.....");
      return "DONT MATCH";
    }
  } else {
    console.log("Banner Null :(");
    return null;
  }
}

/**
 * function to parse banner markings
 * parameter is the banner
 * returns an array of each category being array[1] is cat1 and on for 1, 4 and 7
 * @param { String } banner
 */
function parseBannerMarkings(banner) {
  // const cat1_regex = "TOP[\s]*SECRET|TS|(TS)|SECRET|S|(S)|CONFIDENTIAL|C|(C)|UNCLASSIFIED|U|(U)";
  // const cat4_regex = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS";
  // const cat7_regex = "ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  // const cat4_and_cat7 = "COMINT|-GAMMA|\/|TALENT[\s]*KEYHOLE|SI-G\/TK|HCS|GCS|ORIGINATOR[\s]*CONTROLLED|ORCON|NOT[\s]*RELEASABLE[\s]*TO[\s]*FOREIGN[\s]*NATIONALS|NOFORN|AUTHORIZED[\s]*FOR[\s]*RELEASE[\s]*TO[\s]*USA,[\s]*AUZ,[\s]*NZL|REL[\s]*TO[\s]*USA,[\s]*AUS,[\s]*NZL|CAUTION-PROPERIETARY INFORMATION INVOLVED|PROPIN";
  const cat1_regex = /TOP\s*SECRET|TS|SECRET|S|CONFIDENTIAL|C|UNCLASSIFIED|U/gi;
  const cat4_regex =
    /COMINT|SI|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G|SI-G-[A-Z]{3}|TK|HCS|GCS/gi;
  const cat7_regex =
    /FOR\s*OFFICIAL\s*USE\s*ONLY|FOUO|ORIGINATOR\s*CONTROLLED|ORCON|CONTROLLED\s*IMAGERY|IMCON|SOURCES\s*AND\s*METHODS|SAMI|INFORMATION|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|CAUTION\s*-\s*PROPRIETARY\s*INFORMATION\s*INVOLVED|PROPIN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|RELEASABLE\s*BY\s*INFORMATION\s*DISCLOSURE\s*OFFICIAL|RELIDO|RESTRICTED\s*DATA|RD|-\s*CRITICAL\s*NUCLEAR\s*WEAPON\s*DESIGN\s*INFORMATION|-CNWDI|-SIGMA\s*[0-9]+|FORMERLY\s*RESTRICTED\s*DATA|FRD|DOD\s*CONTROLLED\s*NUCLEAR\s*INFORMATION|DOD\s*UCNI|DOE\s*CONTROLLED\s*NUCLEAR\s*INFORMATION|DOE\s*UCNI|DEA\s*SENSITIVE|FOREIGN\s*INTELLIGENCE\s*SURVEILLANCE\s*ACT/gi;

  //const cat4_and_cat7 = /COMINT|-GAMMA|\/|TALENT\s*KEYHOLE|SI-G|TK|HCS|GCS|ORIGINATOR\s*CONTROLLED|ORCON|NOT\s*RELEASABLE\s*TO\s*FOREIGN\s*NATIONALS|NOFORN|AUTHORIZED\s*FOR\s*RELEASE\s*TO\s*((USA|AUS|NZL)(,)?( *))*|REL\s*TO\s*((USA|AUS|NZL)(,)?( *))*|CAUTION-PROPERIETARY\s*INFORMATION\s*INVOLVED|PROPIN/gi;
  console.log("BEFORE Banner: " + banner);
  const Categories = banner.split("//");

  Categories[0] = Categories[0].trim();

  //console.log(Categories);
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
      console.log(
        "second category doesnt match category 7, running normal program"
      );
      Category_4 = Category(Categories[1], cat4_regex, 4);
      Category_7 = Category(Categories[2], cat7_regex, 7);
    }
  } else {
    console.log("SCI: " + Category_4);
    //console.log("second category returned null");
  }

  console.log("CAT 1: " + Category_1);
  console.log("CAT 4: " + Category_4);
  console.log("CAT 7: " + Category_7);

  const Together = [Category_1, Category_4, Category_7];
  console.log("AFTER Banner: " + banner);
  console.log("TOGETHER: " + Together);

  // if you want to use whats returned from the function you have to do catAbbreviations.classification, etc.
  let catAbbreviations = convertCatToAbrev(Category_4, Category_7);
  console.log("Abbreviated sci: " + catAbbreviations.sci);
  console.log("Abbreviated dissemination: " + catAbbreviations.dissemination);

  //CHANGE
  //KEVIN - If dissem is null then returns "" err msg from checkdissem func. If there is an error with this later on, then maybe err handle before function is called if there is no dissem
  let errMsg = checkDisseminations(Category_1, catAbbreviations.dissemination);

  let val;
  if (Category_4 !== null) {
    val = validateSCI(
      Category_1,
      catAbbreviations.sci,
      catAbbreviations.dissemination
    );
    if ((val[0] = 1)) {
      errMsg += val[1];
    }
  }

  console.log(errMsg);

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
  submarkings = category.split("/");
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
    //console.log("returning category " + categoryNum);
    //console.log(category.toUpperCase());
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
function convertCatToAbrev(sci, dissemination) {
  let abbrevSCI = "";
  let abbrevDissemination = "";

  // Abbreviate SCI
  if (sci != null) {
    if (sci.includes("/")) {
      const sciArray = sci.split("/");
      const abbreviatedSCIs = sciArray.map((sciItem) => {
        // Abbreviate each SCI individually
        if (sciItem === "HCS") {
          return "HCS";
        } else if (sciItem === "COMINT" || sciItem === "SI") {
          return "SI";
        } else if (sciItem.startsWith("COMINT-GAMMA")) {
          return sciItem.replace("COMINT-GAMMA", "SI-G");
        } else if (/^COMINT-ECI [A-Z]+$/.test(sciItem)) {
          return sciItem.replace("COMINT", "SI");
        } else if (/^COMINT-GAMMA-ECI [A-Z]+$/.test(sciItem)) {
          return sciItem.replace("COMINT-GAMMA", "SI-G");
        } else if (/^SI-G-ECI [A-Z]+$/.test(sciItem)) {
          return sciItem;
        } else if (sciItem === "TALENT KEYHOLE" || sciItem === "TK") {
          return "TK";
        }
        // If the SCI abbreviation is not recognized, keep it unchanged
        else {
          return sciItem;
        }
      });
      // Join the abbreviated SCIs back into a single string
      abbrevSCI = abbreviatedSCIs.join("/");
    } else {
      // If there's no "/", abbreviate the single SCI as before
      if (sci === "HCS") {
        abbrevSCI = "HCS";
      } else if (sci === "COMINT" || sci === "SI") {
        abbrevSCI = "SI";
      } else if (sci.startsWith("COMINT-GAMMA")) {
        abbrevSCI = sci.replace("COMINT", "SI");
        abbrevSCI = sci.replace("GAMMA", "G");
      } else if (sci === "SI-G") {
        abbrevSCI = "SI-G";
      } else if (/^COMINT-ECI [A-Z]+$/.test(sci)) {
        abbrevSCI = sci.replace("COMINT", "SI");
      } else if (/^COMINT-GAMMA-ECI [A-Z]+$/.test(sci)) {
        abbrevSCI = sci.replace("COMINT", "SI");
        abbrevSCI = sci.replace("GAMMA", "G");
      } else if (/^SI-G-ECI [A-Z]+$/.test(sci)) {
        abbrevSCI = sci;
      } else if (sci === "TALENT KEYHOLE" || sci === "TK") {
        abbrevSCI = "TK";
      }
    }
  }
  // Abbreviate dissemination
  if (dissemination != null) {
    if (dissemination.includes("/")) {
      const disseminationArray = dissemination.split("/");
      const abbreviatedDisseminations = disseminationArray.map(
        (disseminationItem) => {
          // Abbreviate each dissemination individually
          if (
            disseminationItem === "FOR OFFICIAL USE ONLY" ||
            disseminationItem === "FOUO"
          ) {
            return "FOUO";
          } else if (
            disseminationItem === "ORIGINATOR CONTROLLED" ||
            disseminationItem === "ORCON"
          ) {
            return "ORCON";
          } else if (
            disseminationItem === "CONTROLLED IMAGERY" ||
            disseminationItem === "IMCON"
          ) {
            return "IMCON";
          } else if (
            disseminationItem === "SOURCES AND METHODS" ||
            disseminationItem === "SAMI"
          ) {
            return "SAMI";
          } else if (
            disseminationItem === "NOT RELEASABLE TO FOREIGN NATIONALS" ||
            disseminationItem === "NOFORN"
          ) {
            return "NOFORN";
          } else if (
            disseminationItem === "CAUTION-PROPRIETARY INFORMATION INVOLVED" ||
            disseminationItem === "PROPIN"
          ) {
            return "PROPIN";
          } else if (disseminationItem.startsWith("REL TO ")) {
            return disseminationItem;
          } else if (
            disseminationItem.startsWith("AUTHORIZED FOR RELEASE TO ")
          ) {
            // Extract the additional text after "AUTHORIZED FOR RELEASE TO "
            const additionalText = disseminationItem.substring(
              "AUTHORIZED FOR RELEASE TO ".length
            );
            return "REL TO " + additionalText;
          } else if (disseminationItem.startsWith("REL TO")) {
            console.log("STARTS WITH REL TO");
            return disseminationItem;
          } else if (
            disseminationItem ===
              "RELEASABLE BY INFORMATION DISCLOSURE OFFICIAL" ||
            disseminationItem === "RELIDO"
          ) {
            return "RELIDO";
          } else if (
            /FORMERLY RESTRICTED DATA-SIGMA\s\d{1,2}/.test(disseminationItem)
          ) {
            return disseminationItem.replace(
              "FORMERLY RESTRICTED DATA-SIGMA",
              "FRD-SG"
            );
          } else if (/RESTRICTED DATA-SIGMA\s\d{1,2}/.test(disseminationItem)) {
            return disseminationItem.replace("RESTRICTED DATA-SIGMA", "RD");
          } else if (
            disseminationItem ===
            "FORMERLY RESTRICTED DATA-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION"
          ) {
            return "FRD-CNWDI";
          } else if (
            disseminationItem ===
            "RESTRICTED DATA-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION"
          ) {
            return "RD-CNWDI";
          } else if (
            disseminationItem === "RESTRICTED DATA" ||
            disseminationItem === "RD"
          ) {
            return "RD";
          } else if (
            disseminationItem === "FORMERLY RESTRICTED DATA" ||
            disseminationItem === "FRD"
          ) {
            return "FRD";
          } else if (
            disseminationItem === "DOD CONTROLLED NUCLEAR INFORMATION" ||
            disseminationItem === " "
          ) {
            return " ";
          } else if (
            disseminationItem === "DOE CONTROLLED NUCLEAR INFORMATION" ||
            disseminationItem === " "
          ) {
            return " ";
          } else if (
            disseminationItem === "DEA SENSITIVE" ||
            disseminationItem === "DSEN"
          ) {
            return "DEA SENSITIVE";
          } else if (
            disseminationItem === "FOREIGN INTELLIGENCE SURVEILLANCE ACT" ||
            disseminationItem === "FISA"
          ) {
            return "FISA";
          }
          // If the dissemination abbreviation is not recognized, keep it unchanged
          else {
            return disseminationItem;
          }
        }
      );
      // Join the abbreviated disseminations back into a single string
      abbrevDissemination = abbreviatedDisseminations.join("/");
    } else {
      if (
        dissemination === "FOR OFFICIAL USE ONLY" ||
        dissemination === "FOUO"
      ) {
        abbrevDissemination = "FOUO";
      } else if (
        dissemination === "ORIGINATOR CONTROLLED" ||
        dissemination === "ORCON"
      ) {
        abbrevDissemination = "ORCON";
      } else if (
        dissemination === "CONTROLLED IMAGERY" ||
        dissemination === "IMCON"
      ) {
        abbrevDissemination = "IMCON";
      } else if (
        dissemination === "SOURCES AND METHODS" ||
        dissemination === "SAMI"
      ) {
        abbrevDissemination = "SAMI";
      } else if (
        dissemination === "NOT RELEASABLE TO FOREIGN NATIONALS" ||
        dissemination === "NOFORN"
      ) {
        abbrevDissemination = "NOFORN";
      } else if (
        dissemination === "CAUTION-PROPRIETARY INFORMATION INVOLVED" ||
        dissemination === "PROPIN"
      ) {
        abbrevDissemination = "PROPIN";
      } else if (dissemination.startsWith("REL TO ")) {
        abbrevDissemination = dissemination;
      } else if (dissemination.startsWith("AUTHORIZED FOR RELEASE TO ")) {
        // Extract the additional text after "AUTHORIZED FOR RELEASE TO "
        const additionalText = dissemination.substring(
          "AUTHORIZED FOR RELEASE TO ".length
        );
        abbrevDissemination = "REL TO " + additionalText;
      } else if (
        dissemination === "RELEASABLE BY INFORMATION DISCLOSURE OFFICIAL" ||
        dissemination === "RELIDO"
      ) {
        abbrevDissemination = "RELIDO";
      } else if (
        /FORMERLY RESTRICTED DATA-SIGMA\s\d{1,2}/.test(dissemination)
      ) {
        abbrevDissemination = dissemination.replace(
          "FORMERLY RESTRICTED DATA-SIGMA",
          "FRD-SG"
        );
      } else if (/RESTRICTED DATA-SIGMA\s\d{1,2}/.test(dissemination)) {
        abbrevDissemination = dissemination.replace(
          "RESTRICTED DATA-SIGMA",
          "RD-SG"
        );
      } else if (
        dissemination ===
        "FORMERLY RESTRICTED DATA-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION"
      ) {
        abbrevDissemination = "FRD-CNWDI";
      } else if (
        dissemination ===
        "RESTRICTED DATA-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION"
      ) {
        abbrevDissemination = "RD-CNWDI";
      } else if (
        dissemination === "RESTRICTED DATA" ||
        dissemination === "RD"
      ) {
        abbrevDissemination = "RD";
      } else if (
        dissemination === "FORMERLY RESTRICTED DATA" ||
        dissemination === "FRD"
      ) {
        abbrevDissemination = "FRD";
      } else if (
        dissemination === "DOD CONTROLLED NUCLEAR INFORMATION" ||
        dissemination === "DOD UCNI"
      ) {
        abbrevDissemination = "DOD UCNI";
      } else if (
        dissemination === "DOE CONTROLLED NUCLEAR INFORMATION" ||
        dissemination === "DOE UCNI"
      ) {
        abbrevDissemination = "DOE UCNI";
      } else if (
        dissemination === "DEA SENSITIVE" ||
        dissemination === "DSEN"
      ) {
        abbrevDissemination = "DEA SENSITIVE";
      } else if (
        dissemination === "FOREIGN INTELLIGENCE SURVEILLANCE ACT" ||
        dissemination === "FISA"
      ) {
        abbrevDissemination = "FISA";
      }
    }
  }
  return {
    sci: abbrevSCI,
    dissemination: abbrevDissemination,
  };
}

function ValidateClassification(banner) {
  regex = /TS|TOP *SECRET|S|SECRET|C|CONFIDENTIAL|U|UNCLASSIFIED/gi;
  if (banner.match(regex)) {
    return true;
  }
  return false;
}
function validateSCI(classification, sci, dissemination) {
  let isDissemination = true;
  if (dissemination === null) {
    dissemination = "";
    isDissemination = false;
  }
  console.log(classification + " " + sci + " " + dissemination);
  let valid = 0;
  let msg = "";
  let subBanner = null;
  if (sci) {
    subBanner = sci.split("/");
  }
  if (subBanner === null) {
    subBanner = sci;
  }

  subBanner.forEach((marking) => {
    /**
     * May be used only with
     * TOP SECRET, SECRET, or CONFIDENTIAL. NOFORN is required.
     *
     */
    if (marking.match(/HCS/gi)) {
      if (
        classification.includes("U") ||
        classification.includes("UNCLASSIFIED")
      ) {
        valid = 1;
        msg += "CANNOT USE HCS with UNCLASSIFIED. ";
      }

      if (isDissemination) {
        if (
          dissemination.includes("NOFORN") ||
          dissemination.includes("NOT RELEASABLE TO FOREIGN NATIONALS")
        ) {
        } else {
          valid = 1;
          msg += "HCS MUST USE NOFORN. ";
        }
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
      if (
        classification.includes("U") ||
        classification.includes("UNCLASSIFIED")
      ) {
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
      if (classification.match(/TS|TOP SECRET/gi)) {
      } else {
        valid = 1;
        msg += "CANNOT USE -G with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      }

      if (sci.match(/SI|COMINT/gi)) {
      } else {
        valid = 1;
        msg += "MUST USE SI with -G. ";
      }

      if (isDissemination) {
        if (dissemination.match(/ORCON|ORIGINATOR CONTROLLED/gi)) {
        } else {
          valid = 1;
          msg += "MUST USE ORCON with -G. ";
        }
      } else {
        valid = 1;
        msg += "MUST USE ORCON with -G. ";
      }
    }

    /**
     * May be used only with
     * TOP SECRET. Requires SI
     *
     */
    if (marking.match(/-ECI/gi)) {
      if (classification.match(/TS|TOP SECRET/gi)) {
      } else {
        valid = 1;
        msg += "CANNOT USE -ECI with UNCLASSIFIED, CONFIDENTIAL, or SECRET. ";
      }

      if (sci.match(/SI|COMINT/gi)) {
      } else {
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
      if (classification.match(/^C|CONFIDENTIAL|^U|UNCLASSIFIED/gi)) {
        valid = 1;
        msg += "CANNOT USE TK with UNCLASSIFIED, CONFIDENTIAL. ";
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

  //KEVIN - Trying to fix split error when dissem is null - To remove my changes just remove the if statement from around the whole code
  // KEVIN - Edit UPDATE - 4/18/24 - I think it fixed the error....
  if (dissemination != null) {
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
        if (dissPartsArray[i].match(/(?:[A-Z]{3,4}(?:\/)?)+ EYES ONLY/g)) {
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
        if (dissPartsArray[i].match(/REL TO USA(?:, [A-Z]{3,4})*$/g)) {
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
      if (
        dissPartsArray[i].includes("RD") ||
        dissPartsArray[i].includes("FRD")
      ) {
        if (dissPartsArray[i] === "RD" || dissPartsArray[i] === "FRD") {
          if (classification === "UNCLASSIFIED") {
            errorMsg = "Cannot use RD or FRD with unclassified information.";
          }

          //-CRITICAL NUCLEAR WEAPON DESIGN INFORMATION (-CNWDI): can be used with TOP SECRET or SECRET.
          //Requires RD or FRD.
        } else if (dissPartsArray[i].match(/(RD|FRD)-CNWDI/g)) {
          if (
            classification === "CONFIDENTIAL" ||
            classification === "UNCLASSIFIED"
          ) {
            errorMsg =
              "-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED.";
          }

          //-SIGMA[#] (-SG[#]): may be used with TOP SECRET, SECRET, or CONDFIDENTIAL.
          //Requires RD or FRD. [#] represents the SIGMA number, ranges from 1-99.
        } else if (dissPartsArray[i].match(/(RD|FRD)-SG\s\d{1,2}/g)) {
          if (classification === "UNCLASSIFIED") {
            errorMsg = "-SG cannot be used with UNCLASSIFIED information.";
          }
        } else {
          errorMsg = "Wrong format of banner of RD and FRD.";
        }
      } else if (dissPartsArray[i].includes("-CNWDI")) {
        if (dissPartsArray[i].match(/(RD|FRD)-CNWDI/g)) {
          if (
            classification === "CONFIDENTIAL" ||
            classification === "UNCLASSIFIED"
          ) {
            errorMsg =
              "-CNWDI cannot be used with CONFIDENTIAL or UNCLASSIFIED.";
          }
        } else {
          errorMsg = "RD or FRD is required for -CNWDI.";
        }
      } else if (dissPartsArray[i].includes("-SG")) {
        if (dissPartsArray[i].match(/(RD|FRD)-SG\s\d{1,2}/g)) {
          if (classification === "UNCLASSIFIED") {
            errorMsg = "-SG cannot be used with UNCLASSIFIED information.";
          }
        } else {
          errorMsg = "RD or FRD is required for -SG[#].";
        }
      }

      //DOD or DOE CONTROLLED NUCLEAR INFORMATION (DOD UCNI or DOE UCNI): can only be used with UNCLASSIFIED.
      if (
        dissPartsArray[i] === "DOD UCNI" ||
        dissPartsArray[i] === "DOE UCNI"
      ) {
        if (classification !== "UNCLASSIFIED") {
          errorMsg =
            "DOD/DOE UCNI can only be used with unclassified information.";
        }
      }

      //DEA SENSITIVE (DSEN): can only be used with unclassified.
      if (
        dissPartsArray[i] === "DEA SENSITIVE" &&
        classification !== "UNCLASSIFIED"
      ) {
        errorMsg = "DEA SENSITIVE can only be used with unclassified.";
      }

      //FOREIGN INTELLIGENCE SURVEILLANCE ACT (FISA): does not have any restrictions.
      //No error checking needed.

      // console.log(errorMsg);
    }

    //CHANGE
    return errorMsg;
  }
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
