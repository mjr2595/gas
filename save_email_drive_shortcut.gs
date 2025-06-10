/**
 * @OnlyCurrentDoc
 *
 * This script searches for unread emails under a specified Gmail label,
 * extracts links to Google Docs, creates shortcuts to those docs in a
 * specified Google Drive folder, and marks the emails as read.
 */

// --- CONFIGURATION ---
// Replace this with the exact name of your Gmail label.
const GMAIL_LABEL = "GMAIL_LABEL";

// Replace this with the ID of your target Google Drive folder.
const DRIVE_FOLDER_ID = "DRIVE_FOLDER_ID";
// --- END CONFIGURATION ---

/**
 * Main function to process unread emails and create Drive shortcuts.
 * This is the function you should run or set up a trigger for.
 */
function processUnreadDocs() {
  // Check if the configuration variables have been set.
  if (
    GMAIL_LABEL === "YOUR_GMAIL_LABEL" ||
    DRIVE_FOLDER_ID === "YOUR_DRIVE_FOLDER_ID"
  ) {
    Logger.log(
      "ERROR: Please configure the GMAIL_LABEL and DRIVE_FOLDER_ID variables at the top of the script."
    );
    // Display an alert in the script editor if run manually.
    SpreadsheetApp.getUi().alert(
      "Configuration Needed",
      "Please open the script editor and set your Gmail Label and Drive Folder ID.",
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }

  const query = `label:"${GMAIL_LABEL}" is:unread`;
  Logger.log(`Searching for emails with query: ${query}`);

  try {
    const threads = GmailApp.search(query);
    const targetFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    Logger.log(`Found ${threads.length} unread thread(s). Processing...`);

    threads.forEach((thread) => {
      const messages = thread.getMessages();
      messages.forEach((message) => {
        // Process only if the individual message is unread.
        if (message.isUnread()) {
          const body = message.getPlainBody();

          // Regex to find all Google Docs URLs in the email body.
          const urlRegex =
            /https:\/\/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/g;
          let match;
          let linksFound = false;

          while ((match = urlRegex.exec(body)) !== null) {
            linksFound = true;
            const docId = match[1]; // The captured group from the regex is the doc ID.
            Logger.log(`Found Google Doc link with ID: ${docId}`);

            try {
              // Get the Google Doc file by its ID
              const docFile = DriveApp.getFileById(docId);
              const docName = docFile.getName();

              // Create a shortcut to the document in the target folder.
              targetFolder.createShortcut(docId);
              Logger.log(`Successfully created shortcut for: "${docName}"`);
            } catch (e) {
              Logger.log(
                `Error creating shortcut for doc ID ${docId}. It might be a permission issue or an invalid ID. Error: ${e.toString()}`
              );
            }
          }

          if (linksFound) {
            // Mark the specific message as read after processing.
            message.markRead();
            Logger.log(
              `Marked message with subject "${message.getSubject()}" as read.`
            );
          }
        }
      });
    });
    Logger.log("Processing complete.");
  } catch (e) {
    Logger.log(`An error occurred during script execution: ${e.toString()}`);
  }
}
