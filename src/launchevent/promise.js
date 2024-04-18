/**
I could not get this to work for the life of me
I added the script to the commands.html and to the web
pack config as a wbpack plugin, tried exporting and importing
too but nothing worked so Im not sure how to break up the files
 */

/**
 * Gets the 'to' from email.
 */
function getToRecipients() {
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

export { getToRecipients as getToRecipientsAsync };
