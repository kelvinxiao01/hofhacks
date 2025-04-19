import * as React from "react";
import { makeStyles, Input, Button, Text, Spinner } from "@fluentui/react-components";
import { Send24Regular, Document24Regular, Save24Regular } from "@fluentui/react-icons";
import { AgentService } from "../services/AgentService";
import { AIAgentResponse, ExcelAction } from "../services/ExcelActionProtocol";

interface Message {
  id: string;
  content: string;
  sender: 'user' | 'assistant';
  timestamp: Date;
  action?: {
    type: 'WRITE_CELL' | 'WRITE_RANGE' | 'READ_RANGE';
    data?: any;
  };
  aiResponse?: AIAgentResponse;
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
  },
  input: {
    flex: 1,
  },
  statusContainer: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    fontSize: "12px",
    color: "#666",
  },
  actionIndicator: {
    display: "flex",
    alignItems: "center",
    gap: "4px",
    fontSize: "12px",
    color: "#0078D4",
    marginTop: "4px",
  },
  dataPreview: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: "#E6E6E6",
    borderRadius: "4px",
    fontSize: "12px",
    maxHeight: "100px",
    overflow: "auto",
  },
  actionsList: {
    marginTop: "8px",
    padding: "8px",
    backgroundColor: "#E6E6E6",
    borderRadius: "4px",
    fontSize: "12px",
  },
  actionItem: {
    marginBottom: "4px",
    padding: "4px",
    backgroundColor: "#D0D0D0",
    borderRadius: "4px",
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

    console.log('Sending user message:', inputValue);
    setMessages((prev) => [...prev, userMessage]);
    setInputValue("");
    setIsProcessing(true);
    setStatus("Processing...");

    try {
      // First, try to process the message locally
      console.log('Processing message locally...');
      const localResponse = await agentService.processMessage(inputValue);
      console.log('Local response:', localResponse);
      
      // For read operations, we don't need to send to the AI agent
      const isReadOperation = inputValue.toLowerCase().includes('read') || 
                             inputValue.toLowerCase().includes('show') || 
                             inputValue.toLowerCase().includes('what');
      
      let aiResponse;
      if (!isReadOperation) {
        // Then, send the message to the AI agent
        console.log('Sending message to AI agent...');
        aiResponse = await agentService.sendMessageToAIAgent(inputValue);
        console.log('AI agent response:', aiResponse);
        
        // Process the AI agent response
        console.log('Processing AI agent response...');
        await agentService.processAIAgentResponse(aiResponse);
      } else {
        // For read operations, use the local response as the AI response
        aiResponse = {
          message: localResponse.message,
          actions: []
        };
      }
      
      const assistantMessage: Message = {
        id: (Date.now() + 1).toString(),
        content: localResponse.message || aiResponse.message,
        sender: 'assistant',
        timestamp: new Date(),
        action: localResponse.action,
        aiResponse: aiResponse
      };

      console.log('Adding assistant message:', assistantMessage);
      setMessages((prev) => [...prev, assistantMessage]);
      setStatus("Ready");
    } catch (error) {
      console.error('Error processing message:', error);
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
    if (event.key === "Enter" && !event.shiftKey) {
      event.preventDefault();
      handleSendMessage();
    }
  };

  const renderActionIndicator = (message: Message) => {
    if (!message.action) return null;

    let actionText = "";
    switch (message.action.type) {
      case "WRITE_CELL":
        actionText = `Written to cell ${message.action.data.address}`;
        break;
      case "WRITE_RANGE":
        actionText = `Written to range ${message.action.data.address}`;
        break;
      case "READ_RANGE":
        actionText = `Read from range ${message.action.data.address}`;
        break;
      default:
        actionText = `Action: ${message.action.type}`;
    }

    return (
      <div className={styles.actionIndicator}>
        <Document24Regular />
        <span>{actionText}</span>
      </div>
    );
  };

  const renderDataPreview = (message: Message) => {
    if (!message.action || !message.action.data) return null;

    // For read operations, we don't need to show the raw data preview
    // since the message content already contains the formatted data
    if (message.action.type === 'READ_RANGE') {
      return null;
    }

    return (
      <div className={styles.dataPreview}>
        <pre>{JSON.stringify(message.action.data, null, 2)}</pre>
      </div>
    );
  };

  const renderActionsList = (message: Message) => {
    if (!message.aiResponse || !message.aiResponse.actions || message.aiResponse.actions.length === 0) {
      return null;
    }

    return (
      <div className={styles.actionsList}>
        <Text weight="semibold">Actions to perform:</Text>
        {message.aiResponse.actions.map((action, index) => (
          <div key={index} className={styles.actionItem}>
            <Text weight="semibold">{action.type}</Text>
            {action.description && <div>{action.description}</div>}
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className={styles.chatContainer}>
      <div className={styles.messagesContainer}>
        {messages.map((message) => (
          <div
            key={message.id}
            className={`${styles.message} ${
              message.sender === "user" ? styles.userMessage : styles.assistantMessage
            }`}
          >
            {message.content}
            {message.sender === "assistant" && renderActionIndicator(message)}
            {message.sender === "assistant" && renderDataPreview(message)}
            {message.sender === "assistant" && renderActionsList(message)}
          </div>
        ))}
      </div>
      <div className={styles.inputContainer}>
        <Input
          className={styles.input}
          value={inputValue}
          onChange={(e) => setInputValue(e.target.value)}
          onKeyPress={handleKeyPress}
          placeholder="Ask me to help with Excel..."
          disabled={isProcessing}
        />
        <Button
          appearance="primary"
          icon={<Send24Regular />}
          onClick={handleSendMessage}
          disabled={isProcessing || !inputValue.trim()}
        />
      </div>
      <div className={styles.statusContainer}>
        <Text>{status}</Text>
        {isProcessing && <Spinner size="tiny" />}
      </div>
    </div>
  );
};

export default Chat; 