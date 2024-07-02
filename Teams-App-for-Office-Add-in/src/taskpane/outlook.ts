/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runOutlook;
  }
});

export async function runOutlook() {
  const item = Office.context.mailbox.item;
  const subjectElement = document.getElementById("item-subject");

  // Clear previous content
  subjectElement.textContent = '';

  // Append "Subject:" in bold
  const boldSubject = document.createElement("b");
  boldSubject.textContent = "Subject:";
  subjectElement.appendChild(boldSubject);

  // Append line break
  subjectElement.appendChild(document.createElement("br"));

  // Append the item's subject
  subjectElement.appendChild(document.createTextNode(item.subject));
}
