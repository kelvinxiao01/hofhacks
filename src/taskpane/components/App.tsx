import * as React from "react";
import Header from "./Header";
import Chat from "./Chat";
import { makeStyles } from "@fluentui/react-components";

interface AppProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    height: "100vh",
    backgroundColor: "#FFFFFF",
  },
  content: {
    flex: 1,
    overflow: "hidden",
  },
});

const App: React.FC<AppProps> = (props: AppProps) => {
  const styles = useStyles();

  return (
    <div className={styles.root}>
      <Header title={props.title} />
      <div className={styles.content}>
        <Chat />
      </div>
    </div>
  );
};

export default App;
