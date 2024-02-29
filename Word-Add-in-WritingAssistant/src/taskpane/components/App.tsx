import * as React from "react";
import Header from "./Header";
import AnnotationComponents from "./Annotations";
import { Field, makeStyles } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    minHeight: "10vh",
  },
  welcome__header: {
    display: "flex",
    flexDirection: "column",
    alignItems: "center",
    justifyContent: "center",
  },
});

const App = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header logo="assets/logo-filled.png" title={props.title} message="Welcome" />
      <div className={styles.welcome__header}>
        <Field size="large" label="Discover what the add-in can do for you."></Field>
      </div>
      <AnnotationComponents />
    </div>
  );
};

export default App;
