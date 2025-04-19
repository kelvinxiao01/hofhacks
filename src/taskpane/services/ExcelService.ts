import ExcelJS from 'exceljs';
import { readCell, readRange } from '../taskpane';

// Office.js is available globally in the Excel add-in environment
declare const Excel: any;

export interface ExcelRange {
  address: string;
  values: any[][];
}

export interface ExcelWorksheet {
  name: string;
  ranges: ExcelRange[];
}

export class ExcelService {
  private static instance: ExcelService;
  private workbook: ExcelJS.Workbook;
  private currentFilePath: string | null = null;
  private autoSaveInterval: number | null = null;
  private autoSaveDelay: number = 5000; // 5 seconds
  private isInitialized: boolean = false;
  
  private constructor() {
    this.workbook = new ExcelJS.Workbook();
    // We'll initialize the workbook when the Office.js API is ready
    this.initializeWorkbook();
  }

  public static getInstance(): ExcelService {
    if (!ExcelService.instance) {
      ExcelService.instance = new ExcelService();
    }
    return ExcelService.instance;
  }

  private async initializeWorkbook(): Promise<void> {
    try {
      console.log('Initializing workbook...');
      // Wait for Office.js to be ready
      await this.waitForOfficeJs();
      console.log('Office.js is ready');
      
      // Create a new workbook
      this.workbook = new ExcelJS.Workbook();
      const worksheet = this.workbook.addWorksheet('Sheet1');
      this.isInitialized = true;
      console.log('Created new workbook');
    } catch (error) {
      console.error('Error initializing workbook:', error);
      // Create a new workbook as fallback
      this.workbook = new ExcelJS.Workbook();
      const worksheet = this.workbook.addWorksheet('Sheet1');
      this.isInitialized = true;
      console.log('Created new workbook after error');
    }
  }

  private waitForOfficeJs(): Promise<void> {
    return new Promise((resolve) => {
      if (typeof Office !== 'undefined') {
        resolve();
      } else {
        // Wait for Office.js to be available
        const checkInterval = setInterval(() => {
          if (typeof Office !== 'undefined') {
            clearInterval(checkInterval);
            resolve();
          }
        }, 100);
      }
    });
  }

  public async getCurrentWorksheet(): Promise<ExcelWorksheet> {
    if (!this.isInitialized) {
      console.log('Workbook not initialized, initializing now...');
      await this.initializeWorkbook();
    }

    try {
      console.log('Getting current worksheet using Office.js API...');
      
      // Use the readRange function to get the worksheet data
      const worksheetData = await readRange('A1:Z100');
      console.log('Worksheet data retrieved:', worksheetData);
      
      // Create a worksheet object
      const worksheet = this.workbook.getWorksheet(1) || this.workbook.addWorksheet('Sheet1');
      
      // Update the worksheet with the data from Office.js
      this.updateWorksheetFromOfficeJsData(worksheet, worksheetData || []);
      
      // Get the used range
      const usedRange = {
        address: this.getWorksheetRange(worksheet),
        values: [] as any[][]
      };
      
      // Read values from the worksheet
      console.log('Reading values from worksheet...');
      worksheet.eachRow((row, rowNumber) => {
        console.log(`Reading row ${rowNumber}, cell count: ${row.cellCount}`);
        const rowValues = [] as any[];
        row.eachCell((cell) => {
          console.log(`Cell ${cell.address}: ${cell.value}`);
          rowValues.push(cell.value);
        });
        usedRange.values.push(rowValues);
      });
      
      console.log('Worksheet data:', JSON.stringify(usedRange.values));
      return {
        name: worksheet.name,
        ranges: [usedRange]
      };
    } catch (error) {
      console.error('Error getting current worksheet:', error);
      throw error;
    }
  }
  
  private updateWorksheetFromOfficeJsData(worksheet: ExcelJS.Worksheet, data: any[][]): void {
    // Clear the worksheet
    worksheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.value = null;
      });
    });
    
    // Update the worksheet with the data from Office.js
    data.forEach((row, rowIndex) => {
      row.forEach((value, colIndex) => {
        const cell = worksheet.getCell(rowIndex + 1, colIndex + 1);
        cell.value = value;
      });
    });
  }

  public async writeToCell(address: string, value: any): Promise<void> {
    try {
      console.log(`Writing value "${value}" to cell ${address} using Office.js API...`);
      
      // Use Office.js API to write to the cell
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.values = [[value]];
        await context.sync();
      });
      
      console.log(`Successfully wrote value "${value}" to cell ${address}`);
      
      // Update the local workbook
      if (!this.isInitialized) {
        await this.initializeWorkbook();
      }
      
      const worksheet = this.workbook.getWorksheet(1);
      if (!worksheet) {
        throw new Error('No worksheet found');
      }
      
      const cell = worksheet.getCell(address);
      cell.value = value;
    } catch (error) {
      console.error(`Error writing to cell ${address}:`, error);
      throw error;
    }
  }

  public async writeToRange(address: string, values: any[][]): Promise<void> {
    try {
      console.log(`Writing values to range ${address} using Office.js API...`);
      
      // Use Office.js API to write to the range
      await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(address);
        range.values = values;
        await context.sync();
      });
      
      console.log(`Successfully wrote values to range ${address}`);
      
      // Update the local workbook
      if (!this.isInitialized) {
        await this.initializeWorkbook();
      }
      
      const worksheet = this.workbook.getWorksheet(1);
      if (!worksheet) {
        throw new Error('No worksheet found');
      }
      
      const [startCell, endCell] = address.split(':');
      const startCol = worksheet.getColumn(startCell.replace(/[0-9]/g, ''));
      const startRow = parseInt(startCell.replace(/[A-Z]/g, ''));
      
      values.forEach((row, rowIndex) => {
        row.forEach((value, colIndex) => {
          const cell = worksheet.getCell(startRow + rowIndex, startCol.number + colIndex);
          cell.value = value;
        });
      });
    } catch (error) {
      console.error(`Error writing to range ${address}:`, error);
      throw error;
    }
  }

  public async getSelectedRange(): Promise<ExcelRange> {
    try {
      console.log('Getting selected range using Office.js API...');
      
      // Use Office.js API to get the selected range
      return new Promise((resolve) => {
        Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.load('address, values');
          
          await context.sync();
          console.log('Office.js API: Retrieved selected range:', range.address, range.values);
          
          resolve({
            address: range.address,
            values: range.values
          });
        });
      });
    } catch (error) {
      console.error('Error getting selected range:', error);
      throw error;
    }
  }

  private getWorksheetRange(worksheet: ExcelJS.Worksheet): string {
    const dimensions = worksheet.dimensions;
    const minRow = dimensions.top;
    const maxRow = dimensions.bottom;
    const minCol = dimensions.left;
    const maxCol = dimensions.right;
    
    const startCell = worksheet.getCell(minRow, minCol).address;
    const endCell = worksheet.getCell(maxRow, maxCol).address;
    
    return `${startCell}:${endCell}`;
  }

  public dispose(): void {
    // Clear the autosave interval when the service is disposed
    if (this.autoSaveInterval) {
      clearInterval(this.autoSaveInterval);
      this.autoSaveInterval = null;
    }
  }

  /**
   * Create a pivot table in the current worksheet
   * @param options The options for creating the pivot table
   */
  public async createPivotTable(options: {
    name: string;
    sourceRange: string;
    destinationRange: string;
    rows?: string[];
    columns?: string[];
    values?: Array<{ field: string; function?: string }>;
  }): Promise<void> {
    try {
      console.log('Creating pivot table with options:', options);
      
      // Use Excel.run to create the pivot table
      await Excel.run(async (context) => {
        // Get the active worksheet
        const worksheet = context.workbook.worksheets.getActiveWorksheet();
        
        // Create the pivot table
        const pivotTable = worksheet.pivotTables.add(
          options.name,
          options.sourceRange,
          options.destinationRange
        );
        
        // Add row fields
        if (options.rows && options.rows.length > 0) {
          for (const row of options.rows) {
            try {
              const hierarchy = pivotTable.hierarchies.getItem(row);
              pivotTable.rowHierarchies.add(hierarchy);
            } catch (error) {
              console.warn(`Could not add row field: ${row}`, error);
            }
          }
        }
        
        // Add column fields
        if (options.columns && options.columns.length > 0) {
          for (const column of options.columns) {
            try {
              const hierarchy = pivotTable.hierarchies.getItem(column);
              pivotTable.columnHierarchies.add(hierarchy);
            } catch (error) {
              console.warn(`Could not add column field: ${column}`, error);
            }
          }
        }
        
        // Add value fields
        if (options.values && options.values.length > 0) {
          for (const value of options.values) {
            try {
              const hierarchy = pivotTable.hierarchies.getItem(value.field);
              const dataHierarchy = pivotTable.dataHierarchies.add(hierarchy);
              
              // Set the summarization function (sum, count, average, etc.)
              if (value.function) {
                // Map the function string to the appropriate Excel.AggregationFunction
                let aggregationFunction = Excel.AggregationFunction.sum;
                if (value.function.toLowerCase() === 'count') {
                  aggregationFunction = Excel.AggregationFunction.count;
                } else if (value.function.toLowerCase() === 'average') {
                  aggregationFunction = Excel.AggregationFunction.average;
                } else if (value.function.toLowerCase() === 'max') {
                  aggregationFunction = Excel.AggregationFunction.max;
                } else if (value.function.toLowerCase() === 'min') {
                  aggregationFunction = Excel.AggregationFunction.min;
                }
                
                dataHierarchy.summarizeBy = aggregationFunction;
              }
            } catch (error) {
              console.warn(`Could not add value field: ${value.field}`, error);
            }
          }
        }
        
        await context.sync();
        console.log('Pivot table created successfully');
      });
    } catch (error) {
      console.error('Error creating pivot table:', error);
      throw error;
    }
  }
} 