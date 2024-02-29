/* global Word console */

const insertText = async (text: string) => {
  // Write text to the document.
  try {
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertParagraph(text, Word.InsertLocation.end);
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

let dialog;

export const popupLoginDialog = async () => {
  try {
    Office.context.ui.displayDialogAsync(
      "https://localhost:3000/test.html",
      { height: 30, width: 20 },

      function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
      }
    );
  } catch (error) {
    // Note: In a production add-in, notify the user through your add-in's UI.
    console.error(error);
  }
};

export const processMessage = async (arg: any) => {
  // change font color of selection to be red
  // arg.message is the message from the dialog
  console.log("welcome, " + arg.message);
  dialog.close();

  if (document.getElementById("sideload-msg") != null) {
    document.getElementById("sideload-msg").style.display = "none";
  }
  if (document.getElementById("app-body") != null) {
    document.getElementById("app-body").style.display = "none";
  }
  if (document.getElementById("signin-body") != null) {
    document.getElementById("signin-body").style.display = "flex";
  }
};

export default insertText;
