import * as React from "react";
import { makeStyles, Button, Label } from "@fluentui/react-components";

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    gap: "10px",
    padding: "20px",
  },
  fileInput: {
    display: "none",
  },
  status: {
    marginTop: "8px",
    fontSize: "14px",
  },
  button: {
    borderRadius: "4px",
    width: "auto",
    alignSelf: "flex-start",
    marginBottom: "16px",
  },
});

type UploadStatus = "idle" | "success" | "error" | "uploading";
type StatusMessage = "" | "Please select a PDF file" | "Please select a file first" | "Uploading..." | "File uploaded successfully!" | "Upload failed. Please try again." | "Error uploading file. Please try again.";

const PDFUpload: React.FC = () => {
  const styles = useStyles();
  const [selectedFile, setSelectedFile] = React.useState<File | null>(null);
  const [uploadStatus, setUploadStatus] = React.useState<UploadStatus>("idle");
  const [statusMessage, setStatusMessage] = React.useState<StatusMessage>("");
  const fileInputRef = React.useRef<HTMLInputElement>(null);

  const handleFileSelect = async (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (file && file.type === "application/pdf") {
      setSelectedFile(file);
      setUploadStatus("idle");
      setStatusMessage("");
      
      // Automatically upload the file when selected
      await uploadFile(file);
    } else {
      setUploadStatus("error");
      setStatusMessage("Please select a PDF file");
    }
  };

  const uploadFile = async (file: File) => {
    try {
      setUploadStatus("uploading");
      setStatusMessage("Uploading...");
      
      // Create FormData
      const formData = new FormData();
      formData.append("pdf", file);

      // Dummy endpoint - replace with your actual backend endpoint
      const response = await fetch("http://localhost:3000/api/upload-pdf", {
        method: "POST",
        body: formData,
      });

      if (response.ok) {
        setUploadStatus("success");
        setStatusMessage("File uploaded successfully!");
        setSelectedFile(null);
        if (fileInputRef.current) {
          fileInputRef.current.value = "";
        }
      } else {
        setUploadStatus("error");
        setStatusMessage("Upload failed. Please try again.");
      }
    } catch (error) {
      setUploadStatus("error");
      setStatusMessage("Error uploading file. Please try again.");
      console.error("Upload error:", error);
    }
  };

  return (
    <div className={styles.root}>
      <input
        type="file"
        accept=".pdf"
        onChange={handleFileSelect}
        className={styles.fileInput}
        ref={fileInputRef}
      />
      <Button
        appearance="primary"
        onClick={() => fileInputRef.current?.click()}
        disabled={uploadStatus === "uploading"}
        className={styles.button}
      >
        {uploadStatus === "uploading" ? "Uploading..." : "Select PDF"}
      </Button>
      {statusMessage && <div className={styles.status}>{statusMessage}</div>}
    </div>
  );
};

export default PDFUpload; 