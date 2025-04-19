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

export async function readCell(address: string) {
  // Read text from a cell.
  try {
    let result;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("values");
      await context.sync();
      result = range.values[0][0];
      console.log(`Value in ${address}: ${result}`);
    });
    return result;
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}

export async function readRange(address: string) {
  // Read values from a range.
  try {
    let result;
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      const range = sheet.getRange(address);
      range.load("values");
      await context.sync();
      result = range.values;
      console.log(`Values in ${address}:`, result);
    });
    return result;
  } catch (error) {
    console.log("Error: " + error);
    return null;
  }
}
