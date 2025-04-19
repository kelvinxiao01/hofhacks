/**
 * ExcelActionProtocol.ts
 * 
 * This file defines the protocol for communication between the backend AI agent
 * and the Excel frontend. The protocol is designed to be flexible and extensible,
 * allowing for a wide range of Excel operations to be performed.
 */

import { ExcelRange } from './ExcelService';

/**
 * Represents a single Excel action to be performed
 */
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

/**
 * Types of Excel actions that can be performed
 */
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

/**
 * Response from the AI agent containing actions to perform
 */
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
   * Optional formatted body content to display in the chat
   * This can be used for rich text, formatted tables, or other structured content
   */
  body?: string;
  
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

/**
 * Data for writing to a cell
 */
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
    /**
     * Font options
     */
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
      size?: number;
      name?: string;
    };
    
    /**
     * Fill options
     */
    fill?: {
      type: 'pattern' | 'gradient';
      color?: string;
      pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
    };
    
    /**
     * Border options
     */
    border?: {
      style?: 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted';
      color?: string;
    };
    
    /**
     * Alignment options
     */
    alignment?: {
      horizontal?: 'left' | 'center' | 'right';
      vertical?: 'top' | 'middle' | 'bottom';
      wrapText?: boolean;
    };
    
    /**
     * Number format
     */
    numFmt?: string;
  };
}

/**
 * Data for writing to a range
 */
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

/**
 * Data for creating a pivot table
 */
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

/**
 * Data for inserting a formula
 */
export interface InsertFormulaData {
  /**
   * The cell address (e.g., 'A1')
   */
  address: string;
  
  /**
   * The formula to insert
   */
  formula: string;
}

/**
 * Data for creating a chart
 */
export interface CreateChartData {
  /**
   * The type of chart to create
   */
  type: 'column' | 'bar' | 'line' | 'pie' | 'scatter' | 'area' | 'doughnut' | 'radar';
  
  /**
   * The title of the chart
   */
  title: string;
  
  /**
   * The data range for the chart (e.g., 'A1:B10')
   */
  dataRange: string;
  
  /**
   * The destination range for the chart (e.g., 'D1:H10')
   */
  destinationRange: string;
  
  /**
   * Optional series configuration
   */
  series?: {
    /**
     * The name of the series
     */
    name: string;
    
    /**
     * The x-axis values range
     */
    xValues?: string;
    
    /**
     * The y-axis values range
     */
    yValues: string;
  }[];
}

/**
 * Data for applying conditional formatting
 */
export interface ApplyConditionalFormattingData {
  /**
   * The range to apply conditional formatting to (e.g., 'A1:A10')
   */
  range: string;
  
  /**
   * The type of conditional formatting
   */
  type: 'cellIs' | 'containsText' | 'colorScale' | 'dataBar' | 'iconSet' | 'topBottom' | 'uniqueValues';
  
  /**
   * The criteria for the conditional formatting
   */
  criteria: any;
  
  /**
   * The formatting to apply when the condition is met
   */
  formatting: {
    font?: {
      bold?: boolean;
      italic?: boolean;
      underline?: boolean;
      color?: string;
    };
    fill?: {
      type: 'pattern' | 'gradient';
      color?: string;
      pattern?: 'solid' | 'darkGray' | 'mediumGray' | 'lightGray' | 'none';
    };
  };
} 