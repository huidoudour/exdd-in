/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global Office */

Office.onReady(() => {
  // If needed, Office.js is ready to be called.
  console.log('Office.js is ready for Excel add-in');
});

/**
 * Handles the show taskpane action for Excel add-in
 * @param event {Office.AddinCommands.Event}
 */
function action(event) {
  // For Excel add-ins, we need to show the taskpane
  Office.addin.showAsTaskpane();
  
  // Be sure to indicate when the add-in command function is complete.
  event.completed();
}

// Register the function with Office.
Office.actions.associate("action", action);