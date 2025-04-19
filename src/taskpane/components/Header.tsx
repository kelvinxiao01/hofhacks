import * as React from "react";
import { makeStyles, Title3, Button, Spinner } from "@fluentui/react-components";
import { DocumentPdfRegular } from "@fluentui/react-icons";
import { PDFService, PDFContent } from "../services/pdfService";

interface HeaderProps {
  title: string;
  onPDFProcessed?: (content: PDFContent) => void;
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
  buttonContent: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
});

const Header: React.FC<HeaderProps> = (props: HeaderProps) => {
  const styles = useStyles();
  const fileInputRef = React.useRef<HTMLInputElement>(null);
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  const handleUploadClick = () => {
    setError(null);
    fileInputRef.current?.click();
  };

  const handleFileChange = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file && file.type === "application/pdf") {
      try {
        setIsProcessing(true);
        const content = await PDFService.parsePDF(file);
        props.onPDFProcessed?.(content);
      } catch (err) {
        setError(err instanceof Error ? err.message : 'Failed to process PDF');
        console.error('Error processing PDF:', err);
      } finally {
        setIsProcessing(false);
      }
    } else {
      setError('Please select a valid PDF file');
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
          icon={isProcessing ? <Spinner size="tiny" /> : <DocumentPdfRegular />}
          onClick={handleUploadClick}
          className={styles.uploadButton}
          title="Upload PDF"
          disabled={isProcessing}
        >
          <span className={styles.buttonContent}>
            {isProcessing ? "Processing..." : "Upload PDF"}
          </span>
        </Button>
        {error && (
          <div style={{ color: 'red', fontSize: '12px', marginTop: '4px' }}>
            {error}
          </div>
        )}
      </div>
    </header>
  );
};

export default Header;
