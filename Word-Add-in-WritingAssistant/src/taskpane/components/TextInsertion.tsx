import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import insertText from "../office-document";
import {
  initDocument,
  insertAnnotations,
  getAnnotations,
  acceptFirst,
  rejectLast,
  deleteAnnotations,
} from "../office-document";

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

const TextInsertion: React.FC = () => {
  const [outputText, setOutputText] = useState("");

  let eventContexts = [];

  const registerEventHandlers = async () => {
    // Registers event handlers.
    await Word.run(async (context) => {
      eventContexts[0] = context.document.onParagraphAdded.add(paragraphChanged);
      eventContexts[1] = context.document.onParagraphChanged.add(paragraphChanged);

      eventContexts[2] = context.document.onAnnotationClicked.add(onClickedHandler);
      eventContexts[3] = context.document.onAnnotationHovered.add(onHoveredHandler);
      eventContexts[4] = context.document.onAnnotationInserted.add(onInsertedHandler);
      eventContexts[5] = context.document.onAnnotationRemoved.add(onRemovedHandler);

      await context.sync();
    });
    return "Event handlers registered.\n";
  };

  const paragraphChanged = async (args: Word.ParagraphChangedEventArgs) => {
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
        resultString += `${args.type}: ${result.para.uniqueLocalId} - ${result.text.value}`;
      }
    });
    setOutputText((prevText) => prevText + resultString);
  };

  const onClickedHandler = async (args: Word.AnnotationClickedEventArgs) => {
    let result = "";
    await Word.run(async (context) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");

      await context.sync();

      result = `AnnotationClicked: ${args.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}`;
    });
    setOutputText((prevText) => prevText + result);
  };

  const onHoveredHandler = async (args: Word.AnnotationHoveredEventArgs) => {
    let result = "";
    await Word.run(async (context) => {
      const annotation = context.document.getAnnotationById(args.id);
      annotation.load("critiqueAnnotation");

      await context.sync();

      result = `AnnotationHovered: ${args.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}`;
    });
    setOutputText((prevText) => prevText + result);
  };

  const onInsertedHandler = async (args: Word.AnnotationInsertedEventArgs) => {
    let result = "";
    await Word.run(async (context) => {
      const annotations = [];
      for (let i = 0; i < args.ids.length; i++) {
        let annotation = context.document.getAnnotationById(args.ids[i]);
        annotation.load("id,critiqueAnnotation");

        annotations.push(annotation);
      }

      await context.sync();

      for (let annotation of annotations) {
        result +=
          `AnnotationInserted: ${annotation.id} - ${JSON.stringify(annotation.critiqueAnnotation.critique)}` + "\n";
      }
    });
    setOutputText((prevText) => prevText + result);
  };

  const onRemovedHandler = async (args: Word.AnnotationRemovedEventArgs) => {
    let result = "";
    await Word.run(async () => {
      for (let id of args.ids) {
        result += `AnnotationRemoved: ${id}` + "\n";
      }
    });
    setOutputText((prevText) => prevText + result);
  };

  const deregisterEventHandlers = async () => {
    // Deregisters event handlers.
    await Word.run(async (context) => {
      for (let i = 0; i < eventContexts.length; i++) {
        await Word.run(eventContexts[i].context, async () => {
          eventContexts[i].remove();
        });
      }

      await context.sync();

      eventContexts = [];
    });
    return "Removed event handlers.\n";
  };

  const [text, setText] = useState<string>("");

  const handleTextInsertion = async () => {
    if (text === "" || text === "Some text.") {
      await initDocument(); // If the text is empty, insert the initial text.
    } else {
      await insertText(text);
    }
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();

  const handleOutputTextChange = (event) => {
    setOutputText((prevText) => prevText + event.target.value);
  };

  const handleGetAnnotations = async () => {
    const annotations = await getAnnotations();
    setOutputText((prevText) => prevText + annotations);
  };

  const handleDeleteAnnotations = async () => {
    const result = await deleteAnnotations();
    setOutputText((prevText) => prevText + result);
  };

  const handleClick = async (func) => {
    const result = await func();
    setOutputText((prevText) => prevText + result);
  };

  const clearOutputText = async () => {
    setOutputText("");
  };

  return (
    <div className={styles.textPromptAndInsertion}>
      <Field
        className={styles.textAreaField}
        size="large"
        label="Let us start with inserting some init text into the document. Or you can click the button directly to insert init text."
      ></Field>
      <Button appearance="primary" disabled={false} size="large" onClick={handleTextInsertion}>
        Insert initial text
      </Button>
      <Field className={styles.instructions} size="large" label="Register/Deregister event handlers." />
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <Button
          appearance="primary"
          disabled={false}
          size="large"
          onClick={() => handleClick(registerEventHandlers)}
          style={{ marginRight: "10px" }}
        >
          Register
        </Button>
        <Button appearance="primary" disabled={false} size="large" onClick={() => handleClick(deregisterEventHandlers)}>
          Deregister
        </Button>
      </div>
      <br />
      <Field className={styles.textAreaField} size="large" label="To begin, let's start to insert annotations." />
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <Button appearance="primary" disabled={false} size="large" onClick={() => handleClick(insertAnnotations)}>
          Insert Annotations
        </Button>
        <Button
          appearance="primary"
          disabled={false}
          size="large"
          onClick={handleGetAnnotations}
          style={{ marginRight: "10px" }}
        >
          Get All
        </Button>
        <Button appearance="primary" disabled={false} size="large" onClick={() => handleClick(handleDeleteAnnotations)}>
          Delete All
        </Button>
      </div>
      <Field className={styles.instructions} size="large" label="Accept or reject annotations." />
      <div style={{ display: "flex", justifyContent: "space-between" }}>
        <Button
          appearance="primary"
          disabled={false}
          size="large"
          onClick={acceptFirst}
          style={{ marginRight: "10px" }}
        >
          Accept first.
        </Button>
        <Button appearance="primary" disabled={false} size="large" onClick={rejectLast}>
          Reject last.
        </Button>
      </div>
      <Field className={styles.textAreaField} size="large" label="Output logs"></Field>
      <Button
        style={{ alignSelf: "flex-end" }}
        appearance="primary"
        disabled={false}
        size="large"
        onClick={clearOutputText}
      >
        Clear
      </Button>
      <Textarea
        style={{ width: "100%", height: "300px" }}
        size="large"
        value={outputText}
        onChange={handleOutputTextChange}
        readOnly
      />
    </div>
  );
};

export default TextInsertion;
