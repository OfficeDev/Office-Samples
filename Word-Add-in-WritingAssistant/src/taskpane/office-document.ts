/* global Word console */
export const initDocument = async () => {
  try {
    await Word.run(async (context) => {
      const body = context.document.body;
      body.clear();
      body.insertParagraph(
        "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
        "Start"
      );
      body.insertParagraph(
        "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries.",
        "End"
      );
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const insertFile = async (base64String: any) => {
  try {
    await Word.run(async (context) => {
      context.document.insertFileFromBase64(base64String, "Replace", {
        importTheme: true,
        importStyles: true,
        importParagraphSpacing: true,
        importPageColor: true,
        importChangeTrackingMode: true,
        importCustomProperties: true,
        importCustomXmlParts: true,
        importDifferentOddEvenPages: true,
      });
      await context.sync();
    });
    await Word.run(async (context) => {
      context.document.body.paragraphs.getFirst().select();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
};

export const insertAnnotations = async (args: string[]) => {
  // Adds annotations to the selected paragraph.
  await Word.run(async (context) => {
    for (var arg of args) {
      let paragraph = context.document.getParagraphByUniqueLocalId(arg);
      const critique1 = {
        colorScheme: Word.CritiqueColorScheme.red,
        start: 1,
        length: 3,
      };
      const critique2 = {
        colorScheme: Word.CritiqueColorScheme.green,
        start: 6,
        length: 1,
      };
      const critique3 = {
        colorScheme: Word.CritiqueColorScheme.blue,
        start: 10,
        length: 3,
      };
      const critique4 = {
        colorScheme: Word.CritiqueColorScheme.lavender,
        start: 14,
        length: 3,
      };
      const critique5 = {
        colorScheme: Word.CritiqueColorScheme.berry,
        start: 18,
        length: 10,
      };
      const annotationSet: Word.AnnotationSet = {
        critiques: [critique1, critique2, critique3, critique4, critique5],
      };

      const annotationIds = paragraph.insertAnnotations(annotationSet);

      await context.sync();
      console.log("Annotations inserted: " + annotationIds.value);
    }
  });
};

export let allAnnotationIds: any[] = [];

export const insertInitAnnotations = async () => {
  // Adds annotations to the selected paragraph.
  await Word.run(async (context) => {
    const paragraph = context.document.body.paragraphs.getFirst().getNext();
    const critique1 = {
      colorScheme: Word.CritiqueColorScheme.red,
      start: 207,
      length: 10,
    };
    const critique2 = {
      colorScheme: Word.CritiqueColorScheme.green,
      start: 6,
      length: 1,
    };
    const critique3 = {
      colorScheme: Word.CritiqueColorScheme.blue,
      start: 10,
      length: 3,
    };
    const critique4 = {
      colorScheme: Word.CritiqueColorScheme.lavender,
      start: 14,
      length: 3,
    };
    const critique5 = {
      colorScheme: Word.CritiqueColorScheme.berry,
      start: 18,
      length: 10,
    };
    const annotationSet: Word.AnnotationSet = {
      critiques: [critique1, critique2, critique3, critique4, critique5],
    };

    const annotationIds = paragraph.insertAnnotations(annotationSet);

    await context.sync();
    console.log("Annotations inserted: " + annotationIds.value);
    allAnnotationIds = annotationIds.value;
  });
};

export const getAnnotations = async (): Promise<any> => {
  // Gets annotations found in the selected paragraph.
  let outputText = "";
  await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations = paragraph.getAnnotations();
    annotations.load("id,state,critiqueAnnotation");

    await context.sync();

    console.log("Annotations found:");
    outputText += annotations.items.length + " Annotation(s) found:\n";
    for (var i = 0; i < annotations.items.length; i++) {
      const annotation = annotations.items[i];

      console.log(`${annotation.id} - ${annotation.state} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}`);
      outputText += `${annotation.id} - ${annotation.state} - ${JSON.stringify(
        annotation.critiqueAnnotation.critique
      )}\n`;
    }
  });
  return outputText;
};

export const deleteAnnotations = async () => {
  // Deletes all annotations found in the selected paragraph.
  let result = "";
  await Word.run(async (context) => {
    const paragraph = context.document.getSelection().paragraphs.getFirst();
    const annotations = paragraph.getAnnotations();
    annotations.load("id");

    await context.sync();

    const ids = [];
    const length = annotations.items.length;
    for (var i = 0; i < annotations.items.length; i++) {
      const annotation = annotations.items[i];

      ids.push(annotation.id);
      annotation.delete();
    }

    await context.sync();

    console.log("Annotations deleted:", ids);
    if (length === 0) {
      result = "No annotations found.";
    } else {
      result = length + " Annotation(s) deleted. \n" + ids;
    }
  });
  return result;
};

export const acceptAction = async (id: string, expectedString: string) => {
  await Word.run(async (context) => {
    const annotation = context.document.getAnnotationById(id);
    annotation.load("id,state,critiqueAnnotation");
    annotation.critiqueAnnotation.accept();
    await context.sync();
    let range = annotation.critiqueAnnotation.range;
    range.insertText(expectedString, "Replace");
    await context.sync();
  });
};

export const rejectAction = async (id: string) => {
  await Word.run(async (context) => {
    const annotation = context.document.getAnnotationById(id);
    annotation.load("id,state,critiqueAnnotation");
    annotation.critiqueAnnotation.reject();
    await context.sync();
  });
};

export default initDocument;
