/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

// eslint-disable-next-line @typescript-eslint/no-unused-vars
/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    // eslint-disable-next-line no-undef
    const val = (document.getElementById("tbVals") as HTMLInputElement).value;
    console.log("HERE");
    Office.context.ui.messageParent(val);
  } catch (error) {
    console.error(error);
  }
}
