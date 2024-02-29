import * as React from "react";
import { useState } from "react";
import { Button, Field, Textarea, tokens, makeStyles } from "@fluentui/react-components";
import insertText from "../office-document";

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
  const [text, setText] = useState<string>("Some text.");

  const handleTextInsertion = async () => {
    await insertText(text);
  };

  const handleTextChange = async (event: React.ChangeEvent<HTMLTextAreaElement>) => {
    setText(event.target.value);
  };

  const styles = useStyles();
  let dialog;

  const popupLoginDialog = async () => {
    try {
      Office.context.ui.displayDialogAsync(
        "https://localhost:3000/login.html",
        { height: 30, width: 20 },

        function (asyncResult) {
          dialog = asyncResult.value;
          dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
        }
      );
    } catch (error) {
      console.error(error);
    }
  };

  const handleLogout = async () => {
    setIsLoggedIn(false);
  };

  const [currentEmail, setCurrentEmail] = useState(""); // Add this line

  const processMessage = async (arg: any) => {
    setCurrentEmail(arg.message);
    setIsLoggedIn(true);
    dialog.close();
  };
  const [isLoggedIn, setIsLoggedIn] = useState(false);

  return (
    <div className={styles.textPromptAndInsertion}>
      {isLoggedIn ? (
        <div>
          <div>Welcome, you {currentEmail} are logged in!</div>
          <Button appearance="primary" size="large" onClick={handleLogout}>
            Logout
          </Button>
        </div>
      ) : (
        <Button appearance="primary" size="large" onClick={popupLoginDialog}>
          Login
        </Button>
      )}
    </div>
  );
};

export default TextInsertion;
