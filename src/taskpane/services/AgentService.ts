import { ExcelService, ExcelRange, ExcelWorksheet } from './ExcelService';
import { ExcelActionExecutor } from './ExcelActionExecutor';
import { PivotTableService } from './PivotTableService';
import { AIAgentResponse, ExcelAction, ExcelActionType, WriteCellData, WriteRangeData } from './ExcelActionProtocol';

export interface AgentResponse {
  message: string;
  action?: {
    type: 'WRITE_CELL' | 'WRITE_RANGE' | 'READ_RANGE' | 'CREATE_PIVOT_TABLE';
    data?: any;
  };
  body?: string;
}

export class AgentService {
  private static instance: AgentService;
  private excelService: ExcelService;
  private actionExecutor: ExcelActionExecutor;
  private pivotTableService: PivotTableService;

  private constructor() {
    this.excelService = ExcelService.getInstance();
    this.actionExecutor = ExcelActionExecutor.getInstance();
    this.pivotTableService = PivotTableService.getInstance();
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

    // Check for pivot table commands
    if (lowerMessage.includes('pivot table') || lowerMessage.includes('pivot')) {
      console.log('Detected pivot table command');
      try {
        // Extract source range
        const sourceRangeMatch = lowerMessage.match(/from ([A-Za-z]+\d+:[A-Za-z]+\d+)/i);
        const sourceRange = sourceRangeMatch ? sourceRangeMatch[1] : "A1:D10"; // Default range if not specified
        
        // Extract destination range
        const destRangeMatch = lowerMessage.match(/to ([A-Za-z]+\d+)/i);
        const destinationRange = destRangeMatch ? destRangeMatch[1] : "F1"; // Default range if not specified
        
        // Extract row fields - fixed regex to properly capture the row field
        const rowMatch = lowerMessage.match(/rows as ([^,]+?)(?:\s+and|\s+with|\s+to|$)/i);
        const rows = rowMatch ? [rowMatch[1].trim()] : ["Category"]; // Default row field if not specified
        
        // Extract column fields - fixed regex to properly capture the column field
        const columnMatch = lowerMessage.match(/columns as ([^,]+?)(?:\s+and|\s+with|\s+to|$)/i);
        const columns = columnMatch ? [columnMatch[1].trim()] : ["Region"]; // Default column field if not specified
        
        // Extract value fields - fixed regex to properly capture the value field
        const valueMatch = lowerMessage.match(/sum the ([^,]+?)(?:\s+values|\s+and|\s+with|\s+to|$)/i);
        const values = valueMatch ? [{ field: valueMatch[1].trim(), function: "sum" }] : [{ field: "Sales", function: "sum" }]; // Default value field if not specified
        
        console.log('Pivot table parameters:', { sourceRange, destinationRange, rows, columns, values });
        
        // Create a response with instructions
        const instructions = `I'll create a pivot table with the following configuration:
- Source Range: ${sourceRange}
- Destination Range: ${destinationRange}
- Rows: ${rows.join(', ')}
- Columns: ${columns.join(', ')}
- Values: ${values.map(v => `${v.field} (${v.function})`).join(', ')}`;
        
        return {
          message: 'I can help you create a pivot table. Here are the step-by-step instructions:',
          body: instructions,
          action: {
            type: 'CREATE_PIVOT_TABLE',
            data: {
              sourceRange,
              destinationRange,
              rows,
              columns,
              values
            }
          }
        };
      } catch (error) {
        console.error('Error processing pivot table command:', error);
        return {
          message: 'I encountered an error while processing your pivot table request. Please try again with a clearer command.',
          body: 'Example: "create a pivot table from A1:D10 to F1 with rows: Product, columns: Region, values: Sales sum"'
        };
      }
    }

    // Check for read commands
    if (lowerMessage.includes('read') || lowerMessage.includes('show') || lowerMessage.includes('what')) {
      console.log('Detected read command');
      try {
        console.log('Getting current worksheet...');
        const worksheet = await this.excelService.getCurrentWorksheet();
        console.log('Worksheet retrieved:', worksheet.name);
        console.log('Worksheet data:', JSON.stringify(worksheet.ranges[0].values));
        
        // Format the output in a nicer way
        const formattedOutput = this.formatWorksheetOutput(worksheet.ranges[0].values);
        
        return {
          message: `Here's what I found in the worksheet "${worksheet.name}":`,
          action: {
            type: 'READ_RANGE',
            data: worksheet.ranges[0]
          },
          body: formattedOutput
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

    // Check if this is a pivot table command
    const pivotTableRegex = /create.*pivot.*table/i;
    if (pivotTableRegex.test(message)) {
      // Extract parameters using regex
      const sourceRangeMatch = message.match(/from\s+([A-Z0-9:]+)/i);
      const destinationRangeMatch = message.match(/to\s+([A-Z0-9:]+)/i);
      const rowsMatch = message.match(/rows?\s*:\s*([^,]+)/i);
      const columnsMatch = message.match(/columns?\s*:\s*([^,]+)/i);
      const valuesMatch = message.match(/values?\s*:\s*([^,]+)/i);

      const sourceRange = sourceRangeMatch ? sourceRangeMatch[1] : '';
      const destinationRange = destinationRangeMatch ? destinationRangeMatch[1] : '';
      const rows = rowsMatch ? rowsMatch[1].split(',').map(r => r.trim()) : [];
      const columns = columnsMatch ? columnsMatch[1].split(',').map(c => c.trim()) : [];
      const values = valuesMatch ? valuesMatch[1].split(',').map(v => v.trim()) : [];

      // Create instructions for manual pivot table creation
      const instructions = [
        'To create this pivot table manually:',
        '1. Select your data range: ' + sourceRange,
        '2. Go to Insert > PivotTable',
        '3. Choose "New Worksheet" or select a location',
        '4. In the PivotTable Fields pane:',
        rows.length > 0 ? `   - Drag these fields to Rows: ${rows.join(', ')}` : '',
        columns.length > 0 ? `   - Drag these fields to Columns: ${columns.join(', ')}` : '',
        values.length > 0 ? `   - Drag these fields to Values: ${values.join(', ')}` : '',
        '5. Adjust the layout and formatting as needed'
      ].filter(Boolean).join('\n');

      return {
        message: 'I can help you create a pivot table. Here are the step-by-step instructions:',
        body: instructions,
        action: {
          type: 'CREATE_PIVOT_TABLE',
          data: {
            sourceRange,
            destinationRange,
            rows,
            columns,
            values
          }
        }
      };
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
    
    // Check for read commands
    if (message.toLowerCase().includes('read') || message.toLowerCase().includes('show') || message.toLowerCase().includes('what')) {
      try {
        const worksheet = await this.excelService.getCurrentWorksheet();
        const formattedOutput = this.formatWorksheetOutput(worksheet.ranges[0].values);
        
        // Check if the worksheet has data
        const hasData = !formattedOutput.includes("The worksheet is empty");
        
        return {
          message: hasData ? 
            `Found data in worksheet "${worksheet.name}":` : 
            `The worksheet is currently empty. Try adding some data first!`,
          actions: [],
          body: formattedOutput
        };
      } catch (error) {
        console.error('Error reading worksheet:', error);
        return {
          message: `Error reading worksheet: ${error.message}`,
          actions: []
        };
      }
    }
    
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

  /**
   * Wrap text to fit within a specified width
   * @param text The text to wrap
   * @param width The maximum width
   * @returns An array of wrapped lines
   */
  private wrapText(text: string, width: number): string[] {
    if (text.length <= width) {
      return [text];
    }
    
    const lines: string[] = [];
    let currentLine = '';
    
    // Split by spaces to preserve word boundaries
    const words = text.split(' ');
    
    for (const word of words) {
      // If adding this word would exceed the width, start a new line
      if (currentLine.length + word.length + 1 > width) {
        if (currentLine.length > 0) {
          lines.push(currentLine);
        }
        currentLine = word;
      } else {
        // Add the word to the current line
        if (currentLine.length === 0) {
          currentLine = word;
        } else {
          currentLine += ' ' + word;
        }
      }
    }
    
    // Add the last line if it's not empty
    if (currentLine.length > 0) {
      lines.push(currentLine);
    }
    
    return lines;
  }

  /**
   * Format the worksheet output in a nicer way
   * @param values The values from the worksheet
   * @returns A formatted string representation of the worksheet
   */
  private formatWorksheetOutput(values: any[][]): string {
    if (!values || values.length === 0) {
      return "The worksheet is empty.";
    }

    let output = "";
    let hasContent = false;
    const maxLineLength = 40; // Maximum line length before wrapping

    // Process each row
    values.forEach((row, rowIndex) => {
      // Check if the row has any non-empty cells
      const hasNonEmptyCells = row.some(cell => cell !== null && cell !== undefined && cell !== "");
      
      if (hasNonEmptyCells) {
        hasContent = true;
        
        // Format the row header
        const rowHeader = `Row ${rowIndex + 1}: `;
        output += rowHeader;
        
        // Process each cell in the row
        let currentLine = "";
        let isFirstCell = true;
        
        // Process each cell in the row
        row.forEach((cell, colIndex) => {
          if (cell !== null && cell !== undefined && cell !== "") {
            // Convert column index to letter (0 = A, 1 = B, etc.)
            const colLetter = this.getColumnLetter(colIndex);
            const rowNumber = rowIndex + 1;
            
            // Format the cell value
            let cellValue = cell;
            if (typeof cellValue === 'object') {
              cellValue = JSON.stringify(cellValue);
            }
            
            const cellContent = `${colLetter}${rowNumber}=${cellValue}`;
            
            // Check if adding this cell would exceed the line length
            if (currentLine.length + cellContent.length + (isFirstCell ? 0 : 2) > maxLineLength) {
              // If we're in the middle of a line, start a new line
              if (currentLine.length > 0) {
                output += currentLine + "\n";
                currentLine = "";
                isFirstCell = true;
              }
              
              // If the cell content itself is longer than maxLineLength, wrap it
              if (cellContent.length > maxLineLength) {
                const wrappedLines = this.wrapText(cellContent, maxLineLength);
                wrappedLines.forEach((line, i) => {
                  if (i === 0) {
                    currentLine = line;
                  } else {
                    output += currentLine + "\n";
                    currentLine = line;
                  }
                });
              } else {
                currentLine = cellContent;
              }
            } else {
              // Add the cell to the current line
              if (isFirstCell) {
                currentLine = cellContent;
                isFirstCell = false;
              } else {
                currentLine += ", " + cellContent;
              }
            }
          }
        });
        
        // Add the last line if it's not empty
        if (currentLine.length > 0) {
          output += currentLine;
        }
        
        output += "\n";
      }
    });

    if (!hasContent) {
      return "The worksheet is empty.";
    }

    return output;
  }

  /**
   * Convert a column index to a letter (0 = A, 1 = B, etc.)
   * @param index The column index
   * @returns The column letter
   */
  private getColumnLetter(index: number): string {
    let letter = '';
    let temp = index;
    
    while (temp >= 0) {
      letter = String.fromCharCode(65 + (temp % 26)) + letter;
      temp = Math.floor(temp / 26) - 1;
    }
    
    return letter;
  }
} 