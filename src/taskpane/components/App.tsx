import * as React from "react";
import Header from "./Header";
import Chat from "./Chat";
import { makeStyles } from "@fluentui/react-components";
import { PDFContent } from "../services/pdfService";

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
  const [pdfContent, setPdfContent] = React.useState<PDFContent | null>(null);

  const handlePDFProcessed = (content: PDFContent) => {
    setPdfContent(content);
    // You can pass the PDF content to the Chat component or handle it as needed
    console.log('PDF processed successfully:', content);
  };

  return (
    <div className={styles.root}>
      <Header 
        title={props.title} 
        onPDFProcessed={handlePDFProcessed}
      />
      <div className={styles.content}>
        <Chat pdfContent={pdfContent} />
      </div>
    </div>
  );
};

export default App;
