/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  try {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/dialog.html",
      {
        height: 50,
        width: 30,
        displayInIframe: true,
      },
      function (result) {
        const dialog = result.value;
        dialog.addEventHandler(
          Office.EventType.DialogMessageReceived,
          async (msg: { message: string; origin: string | undefined }) => {
            console.log(msg.message);
            try {
              await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.values = [[msg.message]];
                const letters = "0123456789ABCDEF";
                let color = "#";
                for (let i = 0; i < 6; i++) {
                  color += letters[Math.floor(Math.random() * 16)];
                }
                range.format.fill.color = color;
                await context.sync();
                dialog.close();
              });
            } catch (error) {
              console.error(error);
            }
          }
        );
      }
    );
  } catch (error) {
    console.error(error);
  }
}

// export async function run() {
//   try {
//     await Excel.run(async (context) => {
//       /**
//        * Insert your Excel code here
//        */
//       const range = context.workbook.getSelectedRange();

//       // Read the range address
//       range.load("address");

//       // Update the fill color
//       range.format.fill.color = "yellow";

//       await context.sync();
//       console.log(`The range address was ${range.address}.`);
//     });
//   } catch (error) {
//     console.error(error);
//   }
// }
