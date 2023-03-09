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

Office.initialize = function () {
  // Your add-in's initialization logic, if any, goes here.
};

function trackMessage(event) {
  const buttonId = event.source.id;
  const itemId = Office.context.mailbox.item.id;
  // save this message
  event.completed();
}

function test1234() {
  console.log("test1234");
}

// Link the MyCustomMenu Item with executefucntion to the function
Office.actions.associate("test1234", test1234);

// Register the function with Office.
Office.actions.associate("trackMessage", trackMessage);
