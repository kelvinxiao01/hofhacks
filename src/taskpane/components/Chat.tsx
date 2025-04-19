import * as React from "react";
import { makeStyles, Input, Button, Text, Spinner } from "@fluentui/react-components";
import { Send24Regular, Document24Regular, Save24Regular } from "@fluentui/react-icons";
import { AgentService } from "../services/AgentService";

interface Message {
  id: string;
  content: string;
  sender: 'user' | 'assistant';
  timestamp: Date;
  action?: {
    type: 'WRITE_CELL' | 'WRITE_RANGE' | 'READ_RANGE';
    data?: any;
  };
}

const useStyles = makeStyles({
  chatContainer: {
    display: "flex",
    flexDirection: "column",
    height: "100%",
    padding: "16px",
    gap: "16px",
  },
  messagesContainer: {
    flex: 1,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "12px",
  },
  message: {
    padding: "12px",
    borderRadius: "8px",
    maxWidth: "80%",
    whiteSpace: "pre-wrap",
  },
  userMessage: {
    backgroundColor: "#0078D4",
    color: "white",
    alignSelf: "flex-end",
  },
  assistantMessage: {
    backgroundColor: "#F0F0F0",
    color: "black",
    alignSelf: "flex-start",
  },
  inputContainer: {
    display: "flex",
    gap: "8px",
    padding: "8px",
    borderTop: "1px solid #E1E1E1",
  },
  input: {
    flex: 1,
  },
  actionIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
    fontSize: "12px",
    color: "#666",
    marginTop: "4px",
  },
  dataPreview: {
    maxHeight: "200px",
    overflow: "auto",
    backgroundColor: "#F8F8F8",
    padding: "8px",
    borderRadius: "4px",
    marginTop: "8px",
    fontSize: "12px",
  },
  statusBar: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    padding: "8px",
    backgroundColor: "#F8F8F8",
    borderRadius: "4px",
    fontSize: "12px",
  },
});

const Chat: React.FC = () => {
  const styles = useStyles();
  const [messages, setMessages] = React.useState<Message[]>([]);
  const [inputValue, setInputValue] = React.useState("");
  const [isProcessing, setIsProcessing] = React.useState(false);
  const [status, setStatus] = React.useState("Ready");
  const agentService = AgentService.getInstance();

  const handleSendMessage = async () => {
    if (!inputValue.trim() || isProcessing) return;

    const userMessage: Message = {
      id: Date.now().toString(),
      content: inputValue,
      sender: 'user',
      timestamp: new Date(),
    };

    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");
    setIsProcessing(true);
    setStatus("Processing...");

    try {
      const response = await agentService.processMessage(inputValue);
      
      const assistantMessage: Message = {
        id: (Date.now() + 1).toString(),
        content: response.message,
        sender: 'assistant',
        timestamp: new Date(),
        action: response.action
      };

      setMessages((prev) => [...prev, assistantMessage]);
      setStatus("Ready");
    } catch (error) {
      const errorMessage: Message = {
        id: (Date.now() + 1).toString(),
        content: "Sorry, I encountered an error while processing your request. Please try again.",
        sender: 'assistant',
        timestamp: new Date(),
      };
      setMessages((prev) => [...prev, errorMessage]);
      setStatus("Error occurred");
    } finally {
      setIsProcessing(false);
    }
  };

  const handleKeyPress = (event: React.KeyboardEvent) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      handleSendMessage();
    }
  };

  const renderActionIndicator = (message: Message) => {
    if (!message.action) return null;

    let icon = null;
    let actionText = "";

    switch (message.action.type) {
      case 'WRITE_CELL':
        icon = <Save24Regular />;
        actionText = `Written to cell ${message.action.data.address}`;
        break;
      case 'WRITE_RANGE':
        icon = <Save24Regular />;
        actionText = `Written to range ${message.action.data.address}`;
        break;
      case 'READ_RANGE':
        icon = <Document24Regular />;
        actionText = `Read from ${message.action.data.address}`;
        break;
    }

    return (
      <div className={styles.actionIndicator}>
        {icon}
        <span>{actionText}</span>
      </div>
    );
  };

  const renderDataPreview = (message: Message) => {
    if (!message.action || !message.action.data) return null;

    return (
      <div className={styles.dataPreview}>
        <pre>{JSON.stringify(message.action.data, null, 2)}</pre>
      </div>
    );
  };

  return (
    <div className={styles.chatContainer}>
      <div className={styles.statusBar}>
        <span>{status}</span>
        {isProcessing && <Spinner size="tiny" />}
      </div>
      
      <div className={styles.messagesContainer}>
        {messages.map((message) => (
          <div key={message.id}>
            <div
              className={`${styles.message} ${
                message.sender === 'user' ? styles.userMessage : styles.assistantMessage
              }`}
            >
              <Text>{message.content}</Text>
            </div>
            {renderActionIndicator(message)}
            {renderDataPreview(message)}
          </div>
        ))}
      </div>
      
      <div className={styles.inputContainer}>
        <Input
          className={styles.input}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyPress={handleKeyPress}
          placeholder="Type your message... (e.g., 'Read the current worksheet' or 'Write value 42 to cell A1')"
          disabled={isProcessing}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          onClick={handleSendMessage}
          disabled={isProcessing}
        />
      </div>
    </div>
  );
};

export default Chat; 