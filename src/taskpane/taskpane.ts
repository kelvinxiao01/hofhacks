/* global Excel console */

export async function insertText(text: string) {
  // Write text to the top left cell.
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange("A1");
      range.values = [[text]];
      range.format.autofitColumns();
      await context.sync();
    });
  } catch (error) {
    console.log("Error: " + error);
  }
}

export async function readCell(address: string): Promise<string> {
  // Read text from a cell.
  try {
    let result: string = "";
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("formulas");
      await context.sync();
      result = range.formulas[0][0];
      console.log(`Read formula from cell ${address}:`, result);
    });
    return result;
  } catch (error) {
    console.error(`Error reading formula from cell ${address}:`, error);
    throw error;
  }
}

export async function readRange(address: string): Promise<any[][]> {
  // Read values from a range.
  try {
    let result: any[][] = [];
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("formulas");
      await context.sync();
      result = range.formulas;
      console.log(`Read formulas from range ${address}:`, result);
    });
    return result;
  } catch (error) {
    console.error(`Error reading formulas from range ${address}:`, error);
    throw error;
  }
}
