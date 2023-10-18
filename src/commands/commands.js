/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called
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

  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);

  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

// ##############################################################################
// Creating function that protects the worksheet.
// Toggle logic: sheet.protection.protected, doesn't run until after the sync is complete and the sheet.protection.protected has been assigned the correct value that is fetched from the document, it must come after the await operator ensures sync has completed.
// Specify args parameter to the function

async function toggleProtection(args) {
  try {
    await Excel.run(async (context) => {
      // TODO1: Queue commands to reverse the protection status of the current worksheet.

      // Uses workheet object's protection property in a standard toggle pattern.
      const sheet = context.workbook.worksheets.getActiveWorksheet();

      // TODO2: Queue command to load the sheet's "protection.protected" property from
      // the document and re-synchronize the document and task pane. This is necessary because sheet object is only a proxy object that exists in the task pane's script, unaffected by actual protection state of document, so it's protection.protected property can't have a real value.
      // To avoid an exception error, we must first fetch the protection status from the document and use it to set value of sheet.protection.protected.
      // Fetching process has 3 steps.

      sheet.load("protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }

      await context.sync(); // method that sends the queued commands to the document to be executed.
    });
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }

  args.completed(); // Calling args completed is a requirement for all add-in commands of type 'ExecuteFunction'. It signals to Office client application that the function has finished and the UI cna become responsive again.
}

// Registers the function
Office.actions.associate("toggleProtection", toggleProtection);

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal();

// The add-in command functions need to be available in global scope
g.action = action;
