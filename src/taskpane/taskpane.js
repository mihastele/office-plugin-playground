/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import { getSelectedText, appendTextToHTML } from "../taskpane/util/demoSelect.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("showText").onclick = showText;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */

    // insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    paragraph.font.color = "blue";

    await context.sync();
  });
}

export async function showText() {
  return appendTextToHTML(await getSelectedText());
}
