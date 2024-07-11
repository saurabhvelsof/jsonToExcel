const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// Load JSON data
const jsonFilePath = path.join(__dirname, "output.json");
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf-8"));

// Load the Excel template
const templateFilePath = path.join(__dirname, "FloodStateCR.xlsx");

const loadTemplateAndPopulateData = async () => {
  console.time("Report Generation Time"); // Start the timer

  let workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templateFilePath);
  let worksheet = workbook.getWorksheet(1); // Assuming data goes into the first sheet

  // Find the merge range for a given cell address
  const findMergeRange = (cell) => {
    for (const key in worksheet._merges) {
      const merge = worksheet._merges[key].model;
      if (
        cell.row >= merge.top &&
        cell.row <= merge.bottom &&
        cell.col >= merge.left &&
        cell.col <= merge.right
      ) {
        return merge;
      }
    }
    return null;
  };

  // Find placeholders and their positions
  const placeholders = {};
  worksheet.eachRow((row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      const cellValue = cell.value;
      if (
        cellValue &&
        typeof cellValue === "string" &&
        cellValue.startsWith("[") &&
        cellValue.endsWith("]")
      ) {
        placeholders[cellValue] = {
          cellAddress: cell.address,
          style: cell.style,
          isMerged: cell.isMerged,
          mergeRange: cell.isMerged ? cell.master.address : null,
        };
      }
    });
  });

  // Update placeholders' positions after row insertion
  const updatePlaceholders = (startRow, insertedRow) => {
    for (const key in placeholders) {
      const pos = placeholders[key];
      const currentRow = parseInt(pos.cellAddress.match(/\d+/)[0]);
      if (currentRow >= startRow) {
        const newRow = currentRow + insertedRow;
        const newCellAddress = pos.cellAddress.replace(/\d+/, newRow);
        placeholders[key].cellAddress = newCellAddress;
      }
    }
  };

  // Calculate the number of rows required for each repeated value
  const calculateRowsForRepeatedValues = (dataArray) => {
    let totalRows = 0;
    dataArray.forEach((item) => {
      let maxRowsForValue = 1; // At least one row is required for each value
      for (const key in item) {
        if (key.startsWith("RepeatedValues_") || key.startsWith("RepeatedValues")) {
          const nestedDataArray = item[key];
          const nestedRows = calculateRowsForRepeatedValues(nestedDataArray);
          maxRowsForValue = Math.max(maxRowsForValue, nestedRows);
        }
      }
      totalRows += maxRowsForValue;
    });
    return totalRows;
  };

  // Insert empty rows based on calculated row counts
  const insertEmptyRows = (startRow, dataArray) => {
    const numberOfRows = calculateRowsForRepeatedValues(dataArray);
    worksheet.spliceRows(startRow, 0, ...Array(numberOfRows).fill([]));
    updatePlaceholders(startRow, numberOfRows);
    return numberOfRows;
  };

  // Populate single value placeholders
  const populateSingleValuePlaceholders = (data) => {
    for (const key in data) {
      if (
        !key.startsWith("RepeatedValues_") &&
        !key.startsWith("RepeatedValues")
      ) {
        const pos = placeholders[key];
        if (pos) {
          const cellAddress = pos.cellAddress;
          const cell = worksheet.getCell(cellAddress);
          cell.value = data[key];
          cell.style = pos.style;
        }
      }
    }
  };

  // Populate multiple value placeholders recursively
  const populateMultipleValuePlaceholders = (startRow, dataArray) => {
    dataArray.forEach((item, index) => {
      const newRow = worksheet.getRow(startRow + index);
      for (const key in item) {
        if (
          key.startsWith("RepeatedValues_") ||
          key.startsWith("RepeatedValues")
        ) {
          const nestedDataArray = item[key];
          const nestedPlaceholder = Object.keys(nestedDataArray[0])[0];
          const nestedStartRow = placeholders[nestedPlaceholder].row + 1;
          populateMultipleValuePlaceholders(nestedStartRow, nestedDataArray);
        }
      }

      for (const key in item) {
        const pos = placeholders[key];
        if (pos) {
          const cellAddress = `${String.fromCharCode(64 + pos.col)}${
            startRow + index
          }`;
          const newCell = worksheet.getCell(cellAddress);
          const originalCell = worksheet.getCell(
            `${String.fromCharCode(64 + pos.col)}${pos.row}`
          );

          const mergeInfo = findMergeRange(originalCell);
          if (mergeInfo) {
            const startColChar = String.fromCharCode(64 + mergeInfo.left);
            const endColChar = String.fromCharCode(64 + mergeInfo.right);
            const newMergeRange = `${startColChar}${startRow + index}:${endColChar}${startRow + index}`;
            const startCell = `${startColChar}${startRow + index}`;
            const endCell = `${endColChar}${startRow + index}`;
            if (!isRangeAlreadyMerged(startCell, endCell)) {
              worksheet.mergeCells(newMergeRange);
            }
          }
          newCell.value = item[key];
          newCell.style = originalCell.style;
        }
      }
    });
  };

  // First, insert all the empty rows
  const sheetData = jsonData[0].Sheet1;
  for (const dataKey in sheetData) {
    if (dataKey.startsWith("RepeatedValues_")) {
      const firstPlaceholder = Object.keys(sheetData[dataKey][0])[0];
      if (placeholders[firstPlaceholder]) {
        const startRow = parseInt(placeholders[firstPlaceholder].cellAddress.match(/\d+/)[0]) + 1;
        const rowsInserted = insertEmptyRows(startRow, sheetData[dataKey]);
        console.log(`Inserted ${rowsInserted} rows for ${dataKey}`);
      }
    }
  }

  // Save the workbook to disk
  const tempFilePath = path.join(__dirname, "temp_report.xlsx");
  await workbook.xlsx.writeFile(tempFilePath);

  // Close the workbook
  await new Promise((resolve, reject) => {
    fs.close(fs.openSync(tempFilePath, "r"), (err) => {
      if (err) reject(err);
      else resolve();
    });
  });

  // Reopen the workbook to continue with populating data
  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(tempFilePath);
  worksheet = workbook.getWorksheet(1); // Assuming data goes into the first sheet

  // Populate the values
  for (const dataKey in sheetData) {
    if (dataKey.startsWith("RepeatedValues_")) {
      const firstPlaceholder = Object.keys(sheetData[dataKey][0])[0];
      if (placeholders[firstPlaceholder]) {
        const startRow = parseInt(placeholders[firstPlaceholder].cellAddress.match(/\d+/)[0]) + 1;
        populateMultipleValuePlaceholders(startRow, sheetData[dataKey]);
      }
    } else {
      populateSingleValuePlaceholders(sheetData);
    }
  }

  // Save the final workbook with populated values
  const outputFilePath = path.join(__dirname, "updated_report.xlsx");
  await workbook.xlsx.writeFile(outputFilePath);
  console.log("File saved.");

  console.timeEnd("Report Generation Time"); // End the timer
};

loadTemplateAndPopulateData().catch((error) => {
  console.error("Error processing template:", error);
});
