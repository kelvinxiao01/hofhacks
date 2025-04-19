# Excel AI Agent

This project implements an AI agent that can interact with Excel through a taskpane add-in. The agent can perform various Excel operations based on natural language instructions from the user.

## Architecture

The project consists of the following components:

1. **Frontend (Excel Add-in)**: A React-based taskpane that provides a chat interface for users to interact with the AI agent.

2. **Agent Service**: A service that processes user messages and communicates with the backend AI agent.

3. **Excel Service**: A service that interacts with Excel using the ExcelJS library.

4. **Excel Action Protocol**: A protocol for communication between the backend AI agent and the Excel frontend.

5. **Excel Action Executor**: A service that executes Excel actions based on the protocol.

## Protocol

The Excel Action Protocol defines a standardized way for the backend AI agent to communicate with the Excel frontend. The protocol consists of the following components:

### AIAgentResponse

The `AIAgentResponse` interface defines the structure of a response from the AI agent:

```typescript
export interface AIAgentResponse {
  /**
   * A message to display to the user
   */
  message: string;
  
  /**
   * A list of actions to perform
   */
  actions: ExcelAction[];
  
  /**
   * Optional metadata about the response
   */
  metadata?: {
    /**
     * Whether the actions were successful
     */
    success: boolean;
    
    /**
     * Any error messages
     */
    errors?: string[];
    
    /**
     * Any additional data
     */
    [key: string]: any;
  };
}
```

### ExcelAction

The `ExcelAction` interface defines a single action to be performed in Excel:

```typescript
export interface ExcelAction {
  /**
   * The type of action to perform
   */
  type: ExcelActionType;
  
  /**
   * The data required for the action
   */
  data: any;
  
  /**
   * Optional description of what this action does
   */
  description?: string;
}
```

### ExcelActionType

The `ExcelActionType` enum defines the types of actions that can be performed:

```typescript
export enum ExcelActionType {
  // Cell operations
  WRITE_CELL = 'WRITE_CELL',
  READ_CELL = 'READ_CELL',
  FORMAT_CELL = 'FORMAT_CELL',
  
  // Range operations
  WRITE_RANGE = 'WRITE_RANGE',
  READ_RANGE = 'READ_RANGE',
  FORMAT_RANGE = 'FORMAT_RANGE',
  
  // Worksheet operations
  CREATE_WORKSHEET = 'CREATE_WORKSHEET',
  DELETE_WORKSHEET = 'DELETE_WORKSHEET',
  RENAME_WORKSHEET = 'RENAME_WORKSHEET',
  
  // Formula operations
  INSERT_FORMULA = 'INSERT_FORMULA',
  
  // Chart operations
  CREATE_CHART = 'CREATE_CHART',
  
  // Pivot table operations
  CREATE_PIVOT_TABLE = 'CREATE_PIVOT_TABLE',
  
  // Filter operations
  APPLY_FILTER = 'APPLY_FILTER',
  
  // Conditional formatting
  APPLY_CONDITIONAL_FORMATTING = 'APPLY_CONDITIONAL_FORMATTING',
  
  // Data validation
  APPLY_DATA_VALIDATION = 'APPLY_DATA_VALIDATION',
  
  // Custom operations
  CUSTOM = 'CUSTOM'
}
```

## Action Data Types

The protocol defines several data types for specific actions:

### WriteCellData

```typescript
export interface WriteCellData {
  /**
   * The cell address (e.g., 'A1')
   */
  address: string;
  
  /**
   * The value to write
   */
  value: any;
  
  /**
   * Optional formatting options
   */
  formatting?: {
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      size?: number;
      name?: string;
    };
    fill?: {
      type: 'pattern' | 'gradient';
      color?: string;
      pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
    };
    border?: {
      style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted';
      color?: string;
    };
    alignment?: {
      horizontal?: 'left' | 'center' | 'right';
      vertical?: 'top' | 'middle' | 'bottom';
      wrapText?: boolean;
    };
    numFmt?: string;
  };
}
```

### WriteRangeData

```typescript
export interface WriteRangeData {
  /**
   * The range address (e.g., 'A1:B5')
   */
  address: string;
  
  /**
   * The values to write (2D array)
   */
  values: any[][];
  
  /**
   * Optional formatting options (applied to the entire range)
   */
  formatting?: {
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      size?: number;
      name?: string;
    };
    fill?: {
      type: 'pattern' | 'gradient';
      color?: string;
      pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
    };
    border?: {
      style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted';
      color?: string;
    };
    alignment?: {
      horizontal?: 'left' | 'center' | 'right';
      vertical?: 'top' | 'middle' | 'bottom';
      wrapText?: boolean;
    };
    numFmt?: string;
  };
}
```

### CreatePivotTableData

```typescript
export interface CreatePivotTableData {
  /**
   * The source range address (e.g., 'A1:D10')
   */
  sourceRange: string;
  
  /**
   * The destination range address (e.g., 'F1')
   */
  destinationRange: string;
  
  /**
   * The rows to include in the pivot table
   */
  rows: string[];
  
  /**
   * The columns to include in the pivot table
   */
  columns: string[];
  
  /**
   * The values to include in the pivot table
   */
  values: {
    /**
     * The field to use for the value
     */
    field: string;
    
    /**
     * The function to use (e.g., 'sum', 'count', 'average')
     */
    function: 'sum' | 'count' | 'average' | 'max' | 'min' | 'product' | 'stdDev' | 'stdDevP' | 'var' | 'varP';
  }[];
  
  /**
   * Optional filters to apply
   */
  filters?: string[];
}
```

## Usage

To use the Excel AI Agent:

1. Open the taskpane in Excel.
2. Type a message in the chat input, such as "Create a pivot table for the data in A1:D10".
3. The agent will process your message and perform the requested actions in Excel.

## Example Messages

Here are some example messages you can try:

- "Create a pivot table for the data in A1:D10"
- "Create a chart for the data in A1:B5"
- "Insert a formula to calculate the sum of B2:B9 in cell B10"
- "Write the value 42 to cell A1"
- "Write the values [[1,2],[3,4],[5,6]] to range A1:B3"
- "Read the current worksheet"
- "Show me what's in the selected range"

## Development

### Prerequisites

- Node.js
- npm or yarn
- Excel (desktop or web)

### Setup

1. Clone the repository.
2. Install dependencies: `npm install` or `yarn install`.
3. Start the development server: `npm start` or `yarn start`.
4. Sideload the add-in in Excel.

### Building for Production

To build the add-in for production:

```
npm run build
```

or

```
yarn build
```

## License

This project is licensed under the MIT License - see the LICENSE file for details. 