import * as React from "react";
import { makeStyles, Title3, Button } from "@fluentui/react-components";
import { DocumentPdfRegular } from "@fluentui/react-icons";

interface HeaderProps {
  title: string;
}

const useStyles = makeStyles({
  root: {
    padding: "12px 16px",
    borderBottom: "1px solid #E1E1E1",
    backgroundColor: "#FFFFFF",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
  },
  title: {
    margin: 0,
    fontWeight: 600,
  },
  uploadButton: {
    minWidth: "120px",
    height: "32px",
    borderRadius: "6px",
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const styles = useStyles();
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleUploadClick = () => {
    fileInputRef.current?.click();
  };

  const handleFileChange = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file && file.type === "application/pdf") {
      // TODO: Handle PDF file upload
      console.log("PDF file selected:", file.name);
    }
    // Reset the input value to allow selecting the same file again
    event.target.value = "";
  };

  return (
    <header className={styles.root}>
      <Title3 className={styles.title}>{props.title}</Title3>
      <div>
        <input
          type="file"
          ref={fileInputRef}
          onChange={handleFileChange}
          accept=".pdf"
          style={{ display: "none" }}
        />
        <Button
          appearance="primary"
          icon={<DocumentPdfRegular />}
          onClick={handleUploadClick}
          className={styles.uploadButton}
          title="Upload PDF"
        >
          Upload PDF
        </Button>
      </div>
    </header>
  );
};

export default Header;
