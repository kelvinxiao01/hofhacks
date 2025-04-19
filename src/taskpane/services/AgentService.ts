import { ExcelService, ExcelRange, ExcelWorksheet } from './ExcelService';

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

  private constructor() {
    this.excelService = ExcelService.getInstance();
  }

  public static getInstance(): AgentService {
    if (!AgentService.instance) {
      AgentService.instance = new AgentService();
    }
    return AgentService.instance;
  }

  public async processMessage(message: string): Promise<AgentResponse> {
    // Convert message to lowercase for easier matching
    const lowerMessage = message.toLowerCase();

    // Check for read commands
    if (lowerMessage.includes('read') || lowerMessage.includes('show') || lowerMessage.includes('what')) {
      try {
        const worksheet = await this.excelService.getCurrentWorksheet();
        return {
          message: `Here's what I found in the worksheet "${worksheet.name}":\n${JSON.stringify(worksheet.ranges[0].values, null, 2)}`,
          action: {
            type: 'READ_RANGE',
            data: worksheet.ranges[0]
          }
        };
      } catch (error) {
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