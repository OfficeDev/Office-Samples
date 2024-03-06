/* global Word console */

import * as React from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";
import { allAnnotationIds, insertInitAnnotations } from "../office-document";
import NewModal from "./NewModal";
import FileUploader from "./FileUploader";

const useStyles = makeStyles({
  instructions: {
    fontWeight: tokens.fontWeightSemibold,
    marginTop: "20px",
    marginBottom: "10px",
  },
  textPromptAndInsertion: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
  },
  textAreaField: {
    marginLeft: "20px",
    marginTop: "30px",
    marginBottom: "20px",
    marginRight: "20px",
    maxWidth: "50%",
  },
});

const AnnotationComponents: React.FC = () => {
  const styles = useStyles();
  let eventContexts = [];

  const [state, setModalShow] = React.useState({
    show: false,
    eventName: "",
    eventMessage: "",
    annotationId: "",
    paraIds: [""],
  });

  const handleModalShow = (
    show: boolean,
    eventName: string,
    eventMessage: string,
    annotationId: string,
    paraIds: string[]
  ) => {
    setModalShow({
      show: show,
      eventName: eventName,
      eventMessage: eventMessage,
      annotationId: annotationId,
      paraIds: paraIds,
    });
  };

  const handleGrammerChecking = async () => {
    await insertInitAnnotations();
    await registerEventHandlers();
  };

  const registerEventHandlers = async () => {
    // Registers event handlers.
    await Word.run(
      async (context: {
        document: {
          onParagraphAdded: { add: (arg0: (args: Word.ParagraphChangedEventArgs) => Promise<void>) => any };
          onParagraphChanged: { add: (arg0: (args: Word.ParagraphChangedEventArgs) => Promise<void>) => any };
          onAnnotationClicked: { add: (arg0: (args: Word.AnnotationClickedEventArgs) => Promise<void>) => any };
          onAnnotationHovered: { add: (arg0: (args: Word.AnnotationHoveredEventArgs) => Promise<void>) => any };
          onAnnotationInserted: { add: (arg0: (args: Word.AnnotationInsertedEventArgs) => Promise<void>) => any };
          onAnnotationRemoved: { add: (arg0: (args: Word.AnnotationRemovedEventArgs) => Promise<void>) => any };
        };
        sync: () => any;
      }) => {
        eventContexts[0] = context.document.onParagraphAdded.add(paragraphAdded);
        eventContexts[1] = context.document.onParagraphChanged.add(paragraphChanged);

        eventContexts[2] = context.document.onAnnotationClicked.add(onClickedHandler);
        eventContexts[3] = context.document.onAnnotationHovered.add(onHoveredHandler);
        eventContexts[4] = context.document.onAnnotationInserted.add(onInsertedHandler);
        eventContexts[5] = context.document.onAnnotationRemoved.add(onRemovedHandler);

        await context.sync();
      }
    );
    console.log("Event handlers registered.");
  };

  const paragraphAdded = async (args: Word.ParagraphAddedEventArgs) => {
    //let resultString = "";
    await Word.run(async (context) => {
      const results = [];
      for (let id of args.uniqueLocalIds) {
        let para = context.document.getParagraphByUniqueLocalId(id);
        para.load("uniqueLocalId");

        results.push({ para: para, text: para.getText() });
      }

      await context.sync();
      /*
      for (let result of results) {
        resultString += `${args.type}: ${result.para.uniqueLocalId} - ${result.text.value}` + "\n";
      }*/
    });
    handleModalShow(
      true,
      args.type,
      "New paragraph(s) added, do you want to start checking grammers?",
      "",
      args.uniqueLocalIds
    );
  };

  const paragraphChanged = async (args: Word.ParagraphChangedEventArgs) => {
    //let resultString = "";
    await Word.run(
      async (context: { document: { getParagraphByUniqueLocalId: (arg0: any) => any }; sync: () => any }) => {
        const results = [];
        for (let id of args.uniqueLocalIds) {
          let para = context.document.getParagraphByUniqueLocalId(id);
          para.load("uniqueLocalId");

          results.push({ para: para, text: para.getText() });
        }

        await context.sync();

        //for (let result of results) {
        //  resultString += `${args.type}: ${result.para.uniqueLocalId} - ${result.text.value}` + "\n";
        //}
      }
    );
    //handleModalShow(true, "ParagraphChanged", resultString, "", [""]);
  };

  const onClickedHandler = async (args: Word.AnnotationClickedEventArgs) => {
    await Word.run(async (context) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");

      await context.sync();

      console.log(`AnnotationClicked: ${args.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}`);
    });
  };

  const onHoveredHandler = async (args: Word.AnnotationHoveredEventArgs) => {
    let expectedString = "";
    await Word.run(async (context: { document: { getAnnotationById: (arg0: any) => any }; sync: () => any }) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");

      await context.sync();
      for (var i = 0; i < allAnnotationIds.length; i++) {
        if (allAnnotationIds[i] === args.id) {
          switch (i) {
            case 0:
              expectedString = "effective";
              break;
            case 1:
              expectedString = "a";
              break;
            case 2:
              expectedString = "sov";
              break;
            case 3:
              expectedString = " 64";
              break;
            case 4:
              expectedString = "developme";
              break;
            default:
              break;
          }
        }
      }
      // result = `AnnotationHovered: ${args.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}` + "\n";
    });
    handleModalShow(true, "AnnotationHovered", expectedString, args.id, [""]);
  };

  const onInsertedHandler = async (args: Word.AnnotationInsertedEventArgs) => {
    await Word.run(async (context) => {
      const annotations = [];
      for (let i = 0; i < args.ids.length; i++) {
        let annotation = context.document.getAnnotationById(args.ids[i]);
        annotation.load("id,critiqueAnnotation");

        annotations.push(annotation);
      }

      await context.sync();
      for (let annotation of annotations) {
        console.log(`AnnotationInserted: ${annotation.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}`);
      }
    });
  };

  const onRemovedHandler = async (args: Word.AnnotationRemovedEventArgs) => {
    await Word.run(async () => {
      for (let id of args.ids) {
        console.log(`AnnotationRemoved: ${id}`);
      }
    });
  };

  return (
    <div className={styles.textPromptAndInsertion}>
      <NewModal
        show={state.show}
        handleClose={() => handleModalShow(false, "", "", "", [""])}
        eventName={state.eventName}
        eventMessage={state.eventMessage}
        annotationId={state.annotationId}
        paraIds={state.paraIds}
      />
      <Field className={styles.textAreaField} size="large" label="Step 1. Import your document."></Field>
      <FileUploader />
      <br />
      <Field
        className={styles.textAreaField}
        size="large"
        label="Step 2. Click following button to check grammar."
      ></Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleGrammerChecking}>
        Check
      </Button>
      <br />
      <Field
        className={styles.textAreaField}
        size="large"
        label="Step 3. Hover the annotations to see details."
      ></Field>
    </div>
  );
};

export default AnnotationComponents;
