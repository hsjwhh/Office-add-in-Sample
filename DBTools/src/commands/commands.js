/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

Office.onReady(() => {
  // If needed, Office.js is ready to be called

  Excel.run(function(context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var changeHandler = sheet.onChanged.add(handleChange);
  
    return context.sync().then(function() {
      console.log("Listening for changes...");
    });
  
    function handleChange(eventArgs) {
      var changedRange = eventArgs.getRanges()[0];
      var oldValue = changedRange.oldValues[0][0];
      var newValue = changedRange.values[0][0];
  
      console.log("Cell changed from '" + oldValue + "' to '" + newValue + "'");
    }
  }).catch(function(error) {
    console.log(error);
  });

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

async function toggleProtection(args) {
  try {
    await Excel.run(async (context) => {
      // TODO1: Queue commands to reverse the protection status of the current worksheet.
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      // TODO2: Queue command to load the sheet's "protection.protected" property from
      //        the document and re-synchronize the document and task pane.
      sheet.load("protection/protected");
      await context.sync();

      if (sheet.protection.protected) {
        sheet.protection.unprotect();
      } else {
        sheet.protection.protect();
      }
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  args.completed();
}
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
