import ExcelJS from 'exceljs';

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
      // Wait for Office.js to be ready
      await this.waitForOfficeJs();
      
      // Get the current document data
      const documentData = await this.getCurrentDocumentData();
      if (documentData) {
        // Load the workbook from the document data
        await this.loadWorkbookFromData(documentData);
        this.startAutoSave();
        this.isInitialized = true;
        console.log('Workbook initialized successfully');
      } else {
        // If we can't get document data, create a new workbook
        this.workbook = new ExcelJS.Workbook();
        const worksheet = this.workbook.addWorksheet('Sheet1');
        this.isInitialized = true;
        this.startAutoSave();
        console.log('Created new workbook');
      }
    } catch (error) {
      console.error('Error initializing workbook:', error);
      // Create a new workbook as fallback
      this.workbook = new ExcelJS.Workbook();
      const worksheet = this.workbook.addWorksheet('Sheet1');
      this.isInitialized = true;
      this.startAutoSave();
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

  private getCurrentDocumentData(): Promise<ArrayBuffer | null> {
    return new Promise((resolve) => {
      if (Office && Office.context && Office.context.document) {
        Office.context.document.getFileAsync(Office.FileType.Compressed, (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            const file = result.value;
            file.getSliceAsync(0, (sliceResult) => {
              if (sliceResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve(sliceResult.value.data);
              } else {
                console.warn('Failed to get file slice:', sliceResult.error);
                resolve(null);
              }
            });
          } else {
            console.warn('Failed to get file:', result.error);
            resolve(null);
          }
        });
      } else {
        console.warn('Office.js API not available');
        resolve(null);
      }
    });
  }

  private async loadWorkbookFromData(data: ArrayBuffer): Promise<void> {
    try {
      // Load the workbook from the ArrayBuffer
      await this.workbook.xlsx.load(data);
      console.log('Workbook loaded from document data');
    } catch (error) {
      console.error('Error loading workbook from data:', error);
      throw error;
    }
  }

  private startAutoSave(): void {
    // Clear any existing interval
    if (this.autoSaveInterval) {
      clearInterval(this.autoSaveInterval);
    }

    // Set up a new interval for autosaving
    this.autoSaveInterval = window.setInterval(() => {
      this.autoSave();
    }, this.autoSaveDelay);
  }

  private async autoSave(): Promise<void> {
    if (this.isInitialized) {
      try {
        await this.saveWorkbookToOffice();
        console.log('Workbook autosaved successfully');
      } catch (error) {
        console.error('Error autosaving workbook:', error);
      }
    }
  }

  private async saveWorkbookToOffice(): Promise<void> {
    try {
      // Get the workbook as a buffer
      const buffer = await this.workbook.xlsx.writeBuffer();
      
      // Save the buffer back to the document
      await this.saveBufferToOffice(buffer);
    } catch (error) {
      console.error('Error saving workbook to Office:', error);
      throw error;
    }
  }

  private saveBufferToOffice(buffer: ArrayBuffer): Promise<void> {
    return new Promise((resolve, reject) => {
      if (Office && Office.context && Office.context.document) {
        // Convert ArrayBuffer to base64
        const bytes = new Uint8Array(buffer);
        let binary = '';
        for (let i = 0; i < bytes.byteLength; i++) {
          binary += String.fromCharCode(bytes[i]);
        }
        const base64 = btoa(binary);
        
        // Save the workbook data using the document's save method
        Office.context.document.settings.set('workbookData', base64);
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error('Failed to save workbook'));
          }
        });
      } else {
        reject(new Error('Office.js API not available'));
      }
    });
  }

  public async loadWorkbook(filePath: string): Promise<void> {
    // This method is kept for compatibility but will use Office.js instead
    try {
      const documentData = await this.getCurrentDocumentData();
      if (documentData) {
        await this.loadWorkbookFromData(documentData);
        this.currentFilePath = filePath;
        this.startAutoSave();
        this.isInitialized = true;
      } else {
        throw new Error('Could not get document data');
      }
    } catch (error) {
      console.error('Error loading workbook:', error);
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

  public async getCurrentWorksheet(): Promise<ExcelWorksheet> {
    if (!this.isInitialized) {
      await this.initializeWorkbook();
    }

    const worksheet = this.workbook.getWorksheet(1); // Get first worksheet
    if (!worksheet) {
      throw new Error('No worksheet found');
    }

    const usedRange = {
      address: this.getWorksheetRange(worksheet),
      values: [] as any[][]
    };

    worksheet.eachRow((row) => {
      const rowValues = [] as any[];
      row.eachCell((cell) => {
        rowValues.push(cell.value);
      });
      usedRange.values.push(rowValues);
    });

    return {
      name: worksheet.name,
      ranges: [usedRange]
    };
  }

  public async writeToCell(address: string, value: any): Promise<void> {
    if (!this.isInitialized) {
      await this.initializeWorkbook();
    }

    const worksheet = this.workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error('No worksheet found');
    }

    const cell = worksheet.getCell(address);
    cell.value = value;
    
    // Trigger autosave after writing
    this.autoSave();
  }

  public async writeToRange(address: string, values: any[][]): Promise<void> {
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
    
    // Trigger autosave after writing
    this.autoSave();
  }

  public async getSelectedRange(): Promise<ExcelRange> {
    if (!this.isInitialized) {
      await this.initializeWorkbook();
    }

    const worksheet = this.workbook.getWorksheet(1);
    if (!worksheet) {
      throw new Error('No worksheet found');
    }

    // In ExcelJS, we'll get the dimensions of the worksheet as the selected range
    const range = {
      address: this.getWorksheetRange(worksheet),
      values: [] as any[][]
    };

    worksheet.eachRow((row) => {
      const rowValues = [] as any[];
      row.eachCell((cell) => {
        rowValues.push(cell.value);
      });
      range.values.push(rowValues);
    });

    return range;
  }

  public async saveWorkbook(filePath: string): Promise<void> {
    try {
      await this.saveWorkbookToOffice();
      this.currentFilePath = filePath;
    } catch (error) {
      console.error('Error saving workbook:', error);
      throw error;
    }
  }
  
  public dispose(): void {
    // Clear the autosave interval when the service is disposed
    if (this.autoSaveInterval) {
      clearInterval(this.autoSaveInterval);
      this.autoSaveInterval = null;
    }
  }
} 