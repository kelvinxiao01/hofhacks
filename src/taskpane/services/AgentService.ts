import { ExcelService, ExcelRange, ExcelWorksheet } from './ExcelService';
import { ExcelActionExecutor } from './ExcelActionExecutor';
import { AIAgentResponse, ExcelAction, ExcelActionType, WriteCellData, WriteRangeData } from './ExcelActionProtocol';

export interface AgentResponse {
  message: string;
  action?: {
    type: 'WRITE_CELL' | 'WRITE_RANGE' | 'READ_RANGE';
    data?: any;
  };
}

export class AgentService {
  private static instance: AgentService;
  private excelService: ExcelService;
  private actionExecutor: ExcelActionExecutor;

  private constructor() {
    this.excelService = ExcelService.getInstance();
    this.actionExecutor = ExcelActionExecutor.getInstance();
  }

  public static getInstance(): AgentService {
    if (!AgentService.instance) {
      AgentService.instance = new AgentService();
    }
    return AgentService.instance;
  }

  /**
   * Process a message from the user and return a response
   * @param message The message from the user
   * @returns A promise that resolves to an AgentResponse
   */
  public async processMessage(message: string): Promise<AgentResponse> {
    // Convert message to lowercase for easier matching
    const lowerMessage = message.toLowerCase();
    console.log('Processing message:', message);

    // Check for read commands
    if (lowerMessage.includes('read') || lowerMessage.includes('show') || lowerMessage.includes('what')) {
      console.log('Detected read command');
      try {
        console.log('Getting current worksheet...');
        const worksheet = await this.excelService.getCurrentWorksheet();
        console.log('Worksheet retrieved:', worksheet.name);
        console.log('Worksheet data:', JSON.stringify(worksheet.ranges[0].values));
        
        return {
          message: `Here's what I found in the worksheet "${worksheet.name}":\n${JSON.stringify(worksheet.ranges[0].values, null, 2)}`,
          action: {
            type: 'READ_RANGE',
            data: worksheet.ranges[0]
          }
        };
      } catch (error) {
        console.error('Error reading worksheet:', error);
        return {
          message: `Error reading worksheet: ${error.message}`
        };
      }
    }

    // Check for write commands
    if (lowerMessage.includes('write') || lowerMessage.includes('put') || lowerMessage.includes('set')) {
      // Extract cell reference and value from message
      const cellMatch = message.match(/[A-Z]+\d+/);
      const valueMatch = message.match(/value\s+(.+)$/i);

      if (cellMatch && valueMatch) {
        const cellAddress = cellMatch[0];
        const value = valueMatch[1].trim();

        try {
          await this.excelService.writeToCell(cellAddress, value);

          return {
            message: `I've written the value "${value}" to cell ${cellAddress}`,
            action: {
              type: 'WRITE_CELL',
              data: { address: cellAddress, value }
            }
          };
        } catch (error) {
          return {
            message: `Error writing to cell: ${error.message}`
          };
        }
      }
    }

    // Check for range write commands
    if (lowerMessage.includes('range') && (lowerMessage.includes('write') || lowerMessage.includes('fill'))) {
      const rangeMatch = message.match(/range\s+([A-Z]+\d+:[A-Z]+\d+)/i);
      const valuesMatch = message.match(/values\s+\[(.*)\]/i);

      if (rangeMatch && valuesMatch) {
        const rangeAddress = rangeMatch[1];
        try {
          // Parse the values string into a 2D array
          const valuesString = valuesMatch[1];
          const values = this.parseValuesString(valuesString);
          
          await this.excelService.writeToRange(rangeAddress, values);

          return {
            message: `I've written values to range ${rangeAddress}`,
            action: {
              type: 'WRITE_RANGE',
              data: { address: rangeAddress, values }
            }
          };
        } catch (error) {
          return {
            message: `Error writing to range: ${error.message}`
          };
        }
      }
    }

    // Check for selected range commands
    if (lowerMessage.includes('selected') || lowerMessage.includes('selection')) {
      try {
        const range = await this.excelService.getSelectedRange();
        return {
          message: `Here's what's in the selected range ${range.address}:\n${JSON.stringify(range.values, null, 2)}`,
          action: {
            type: 'READ_RANGE',
            data: range
          }
        };
      } catch (error) {
        return {
          message: `Error reading selected range: ${error.message}`
        };
      }
    }

    // Default response for unrecognized commands
    return {
      message: "I can help you read and write to Excel. Try commands like:\n" +
               "- 'Read the current worksheet'\n" +
               "- 'Write value 42 to cell A1'\n" +
               "- 'Write range A1:B3 values [[1,2],[3,4],[5,6]]'\n" +
               "- 'Show me what's in the selected range'"
    };
  }

  /**
   * Process a response from the AI agent and execute the actions
   * @param response The response from the AI agent
   * @returns A promise that resolves when all actions have been executed
   */
  public async processAIAgentResponse(response: AIAgentResponse): Promise<void> {
    try {
      // Execute all actions in the response
      await this.actionExecutor.executeActions(response.actions);
    } catch (error) {
      console.error('Error executing AI agent actions:', error);
      throw error;
    }
  }

  /**
   * Process a message from the user and send it to the AI agent
   * @param message The message from the user
   * @returns A promise that resolves to an AIAgentResponse
   */
  public async sendMessageToAIAgent(message: string): Promise<AIAgentResponse> {
    // In a real implementation, this would send the message to the AI agent
    // and wait for a response. For now, we'll simulate a response.
    
    // Example response for a pivot table request
    if (message.toLowerCase().includes('pivot table')) {
      return this.createPivotTableResponse();
    }
    
    // Example response for a chart request
    if (message.toLowerCase().includes('chart')) {
      return this.createChartResponse();
    }
    
    // Example response for a formula request
    if (message.toLowerCase().includes('formula')) {
      return this.createFormulaResponse();
    }
    
    // Default response
    return {
      message: "I understand you want to work with Excel. I'll help you with that.",
      actions: []
    };
  }

  /**
   * Create a sample pivot table response
   * @returns A sample AIAgentResponse for a pivot table
   */
  private createPivotTableResponse(): AIAgentResponse {
    return {
      message: "I'll create a pivot table for you based on the data in the current worksheet.",
      actions: [
        {
          type: ExcelActionType.CREATE_PIVOT_TABLE,
          description: "Create a pivot table from the data in the current worksheet",
          data: {
            sourceRange: "A1:D10",
            destinationRange: "F1",
            rows: ["Category"],
            columns: ["Region"],
            values: [
              {
                field: "Sales",
                function: "sum"
              }
            ]
          }
        }
      ]
    };
  }

  /**
   * Create a sample chart response
   * @returns A sample AIAgentResponse for a chart
   */
  private createChartResponse(): AIAgentResponse {
    return {
      message: "I'll create a chart for you based on the data in the current worksheet.",
      actions: [
        {
          type: ExcelActionType.CREATE_CHART,
          description: "Create a column chart from the data in the current worksheet",
          data: {
            type: "column",
            title: "Sales by Region",
            dataRange: "A1:B5",
            destinationRange: "D1:H10"
          }
        }
      ]
    };
  }

  /**
   * Create a sample formula response
   * @returns A sample AIAgentResponse for a formula
   */
  private createFormulaResponse(): AIAgentResponse {
    return {
      message: "I'll insert a formula for you in the current worksheet.",
      actions: [
        {
          type: ExcelActionType.INSERT_FORMULA,
          description: "Insert a SUM formula to calculate the total sales",
          data: {
            address: "B10",
            formula: "=SUM(B2:B9)"
          }
        }
      ]
    };
  }

  private parseValuesString(valuesString: string): any[][] {
    // Simple parser for values in format [[1,2],[3,4]]
    try {
      // Replace single quotes with double quotes for valid JSON
      const jsonString = valuesString.replace(/'/g, '"');
      return JSON.parse(`[${jsonString}]`);
    } catch (error) {
      throw new Error(`Invalid values format: ${error.message}`);
    }
  }

  public async getSelectedRange(): Promise<ExcelRange> {
    return this.excelService.getSelectedRange();
  }
} 