/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
 
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;

});

export async function run() {
  /**
   * Insert your Outlook code here
   */
  Office.context.mailbox.item.body.setAsync(
    "Hii from Time.Dev",
    {
      coercionType: "html", // Write text as HTML
    },

    // Callback method to check that setAsync succeeded
    function (asyncResult) {
      if (asyncResult.status ==
        Office.AsyncResultStatus.Failed) {
        write(asyncResult.error.message);
      }
    }
  );
}
