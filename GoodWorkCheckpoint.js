const ExcelJS = require('exceljs');
const fs = require('fs');

async function insertValuesBelowPlaceholders(filePath, placeholders, valuesArray) {
  let workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  let worksheet = workbook.getWorksheet(1); // Assuming the first sheet

  const placeholdersInfo = {};

  // Find placeholder columns and their styles
  placeholders.forEach((placeholder) => {
    worksheet.eachRow((row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.value === placeholder) {
          placeholdersInfo[placeholder] = {
            cellAddress: cell.address,
            style: cell.style,
            isMerged: cell.isMerged,
            mergeRange: cell.isMerged ? cell.master.address : null
          };
        }
      });
    });
  });

  // Determine starting row for inserting values
  const startRow = Object.values(placeholdersInfo).reduce((minRow, info) => {
    const rowNumber = parseInt(info.cellAddress.match(/\d+/)[0]);
    return Math.max(minRow, rowNumber + 1);
  }, 1);

  // Insert the number of rows needed
  worksheet.spliceRows(startRow, 0, ...Array(valuesArray.length).fill([]));

  // Commit the changes after inserting rows
  await workbook.xlsx.writeFile(filePath);

  // Ensure the file is closed before reopening
  await new Promise((resolve, reject) => {
    fs.close(fs.openSync(filePath, 'r'), (err) => {
      if (err) reject(err);
      else resolve();
    });
  });

  // Reopen the workbook to continue with populating data
  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  worksheet = workbook.getWorksheet(1);

  // Insert values below placeholders
  valuesArray.forEach((values, rowIndex) => {
    const currentRow = startRow + rowIndex;

    Object.keys(placeholdersInfo).forEach((placeholder, index) => {
      const info = placeholdersInfo[placeholder];
      const column = info.cellAddress.match(/[A-Z]+/)[0];
      const mergeRange = info.mergeRange ? info.mergeRange.match(/[A-Z]+/)[0] : null;

      if (info.isMerged) {
        const mergeRangeStart = column + currentRow;
        const mergeRangeEnd = mergeRange + currentRow;
        worksheet.unMergeCells(mergeRangeStart, mergeRangeEnd);
        worksheet.mergeCells(`${mergeRangeStart}:${mergeRangeEnd}`);

        // Apply style to merged cells
        const startCell = worksheet.getCell(mergeRangeStart);
        const endCell = worksheet.getCell(mergeRangeEnd);
        applyStyleToRange(startCell, endCell, info.style);
      }

      const cell = worksheet.getCell(column + currentRow);
      cell.value = values[index];
      cell.style = info.style;
    });
  });

  await workbook.xlsx.writeFile(filePath);
  console.log('Values inserted successfully');
}

// Function to apply style to a range of cells
function applyStyleToRange(startCell, endCell, style) {
  const startCol = columnToNumber(startCell.address.replace(/\d+/, ''));
  const endCol = columnToNumber(endCell.address.replace(/\d+/, ''));
  const startRow = parseInt(startCell.address.match(/\d+/)[0]);
  const endRow = parseInt(endCell.address.match(/\d+/)[0]);

  for (let row = startRow; row <= endRow; row++) {
    for (let col = startCol; col <= endCol; col++) {
      const cell = startCell.worksheet.getCell(row, col);
      cell.style = style;
    }
  }
}

// Function to convert column letters to numbers
function columnToNumber(column) {
  let number = 0;
  for (let i = 0; i < column.length; i++) {
    number = number * 26 + (column.charCodeAt(i) - ('A'.charCodeAt(0) - 1));
  }
  return number;
}

const filePath = 'template.xlsx';
const placeholders = ['[p1]', '[p2]', '[p3]']; // Replace with your actual placeholders
const valuesArray = [
  ['Value1', 'Value2', 'Value3'], // Values for the first row
  ['Value4', 'Value5', 'Value6'], // Values for the second row
  ['Value7', 'Value8', 'Value9'], // Values for the third row
];

insertValuesBelowPlaceholders(filePath, placeholders, valuesArray);
