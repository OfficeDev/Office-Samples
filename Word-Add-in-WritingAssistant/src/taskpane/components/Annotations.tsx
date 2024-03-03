/* global Word console */

import * as React from "react";
import { Button, Field, tokens, makeStyles } from "@fluentui/react-components";
import { initDocument, insertAnnotations } from "../office-document";
import NewModal from "./NewModal";

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
  });

  const handleModalShow = (show: boolean, eventName: string, eventMessage: string, annotationId: string) => {
    setModalShow({ show: show, eventName: eventName, eventMessage: eventMessage, annotationId: annotationId });
  };

  const handleGrammerChecking = async () => {
    await insertAnnotations();
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
    let resultString = "";
    await Word.run(async (context) => {
      const results = [];
      for (let id of args.uniqueLocalIds) {
        let para = context.document.getParagraphByUniqueLocalId(id);
        para.load("uniqueLocalId");

        results.push({ para: para, text: para.getText() });
      }

      await context.sync();

      for (let result of results) {
        resultString += `${args.type}: ${result.para.uniqueLocalId} - ${result.text.value}` + "\n";
      }
    });
    handleModalShow(true, args.type, resultString, "");
  };

  const paragraphChanged = async (args: Word.ParagraphChangedEventArgs) => {
    let resultString = "";
    await Word.run(
      async (context: { document: { getParagraphByUniqueLocalId: (arg0: any) => any }; sync: () => any }) => {
        const results = [];
        for (let id of args.uniqueLocalIds) {
          let para = context.document.getParagraphByUniqueLocalId(id);
          para.load("uniqueLocalId");

          results.push({ para: para, text: para.getText() });
        }

        await context.sync();

        for (let result of results) {
          resultString += `${args.type}: ${result.para.uniqueLocalId} - ${result.text.value}` + "\n";
        }
      }
    );
    handleModalShow(true, "ParagraphChanged", resultString, "");
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
    let result = "";
    await Word.run(async (context: { document: { getAnnotationById: (arg0: any) => any }; sync: () => any }) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");

      await context.sync();

      result = `AnnotationHovered: ${args.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}` + "\n";
    });
    handleModalShow(true, "AnnotationHovered", result, args.id);
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
        handleClose={() => handleModalShow(false, "", "", "")}
        eventName={state.eventName}
        eventMessage={state.eventMessage}
        annotationId={state.annotationId}
      />
      <Field
        className={styles.textAreaField}
        size="large"
        label="Let us start with inserting some init text into the document. Or you can click the button directly to insert init text."
      ></Field>
      <Button appearance="primary" disabled={false} size="large" onClick={initDocument}>
        Insert initial text and register events.
      </Button>
      <br />
      <Button appearance="primary" disabled={false} size="large" onClick={handleGrammerChecking}>
        Check Grammers.
      </Button>
    </div>
  );
};

export default AnnotationComponents;
