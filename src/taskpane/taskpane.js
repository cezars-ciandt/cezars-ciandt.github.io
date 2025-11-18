/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Initialize Office Add-in - everything runs under Excel context
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
      console.log("Running in Excel context");
    };

  document.getElementById("submitConnection").onclick = saveToken;

});
    
function saveToken() {
  const tokenInput = document.getElementById("connectionInput").value;

  console.log(`Token input value: ${tokenInput}`);

  const token = tokenInput.trim();

  console.log(`Token value: ${token}`);
  

  Office.context.roamingSettings.set("flowToken", token);

  Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Token saved");
        } else {
            console.error("Error saving token:", result.error.message);
        }

  });
};