const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// Load JSON data
const jsonFilePath = path.join(__dirname, "./shortJsonData/section3.json");
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
  worksheet.eachRow((row, rowNumber) => {
    row.eachCell((cell, colNumber) => {
      const cellValue = cell.value;
      if (
        typeof cellValue === "string" &&
        cellValue.startsWith("[") &&
        cellValue.endsWith("]")
      ) {
        const mergeRange = findMergeRange(cell);
        if (mergeRange) {
          // If the cell is merged, use the starting cell of the merge range
          placeholders[cellValue] = {
            row: mergeRange.top,
            col: mergeRange.left,
          };
        } else {
          // If the cell is not merged, use the current cell's position
          placeholders[cellValue] = { row: rowNumber, col: colNumber };
        }
      }
    });
  });

  // Insert empty rows based on the count of `RepeatedValues_$`
  const getMaxRepeatedValuesCount = (data) => {
    let maxCount = 0;
    for (const key in data) {
      if (key.startsWith("RepeatedValues_")) {
        maxCount = Math.max(maxCount, data[key].length);
      }
    }
    return maxCount;
  };

  const maxCount = getMaxRepeatedValuesCount(jsonData[0].Sheet1);
  const startRow = Object.values(placeholders)[0].row + 1;

  // Insert the number of rows needed
  worksheet.spliceRows(startRow, 0, ...Array(maxCount).fill([]));

  // Save the workbook to a buffer
  const buffer = await workbook.xlsx.writeBuffer();

  // Clear the original workbook from memory (optional, to aid garbage collection)
  workbook = null;

  // Reopen the workbook from the buffer
  workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  worksheet = workbook.getWorksheet(1);

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

  // Copy styles from the placeholder cell to the new cell
  const copyCellStyles = (sourceCell, targetCell) => {
    targetCell.style = { ...sourceCell.style };
  };

  // Check if a range is already merged
  const isRangeAlreadyMerged = (startCell, endCell) => {
    for (const key in worksheet._merges) {
      const merge = worksheet._merges[key].model;
      const mergeStartCell = `${String.fromCharCode(64 + merge.left)}${
        merge.top
      }`;
      const mergeEndCell = `${String.fromCharCode(64 + merge.right)}${
        merge.bottom
      }`;
      if (
        (startCell >= mergeStartCell && startCell <= mergeEndCell) ||
        (endCell >= mergeStartCell && endCell <= mergeEndCell)
      ) {
        return true;
      }
    }
    return false;
  };

  // Update placeholders' positions after row insertion
  const updatePlaceholders = (insertedRow) => {
    for (const key in placeholders) {
      if (placeholders[key].row >= insertedRow) {
        placeholders[key].row += 1;
      }
    }
  };

  // Populate single value placeholders
  const populateSingleValuePlaceholders = (data) => {
    for (const key in data) {
      if (
        !key.startsWith("RepeatedValues_") &&
        !key.startsWith("RepeatedValues")
      ) {
        const { row, col } = placeholders[key];
        worksheet.getRow(row).getCell(col).value = data[key];
      }
    }
  };

  // Populate multiple value placeholders recursively
  const populateMultipleValuePlaceholders = (startRow, dataArray) => {
    // Reverse the dataArray to populate in the correct order
    dataArray.reverse().forEach((item, index) => {
      const newRow = worksheet.insertRow(startRow, [], "i");
      const originalRow = worksheet.getRow(startRow - 1);

      if (
        originalRow._cells &&
        newRow._cells &&
        originalRow._cells.length === newRow._cells.length
      ) {
        for (let i = 1; i < originalRow._cells.length; i++) {
          if (originalRow._cells[i]._mergeCount !== undefined) {
            newRow._cells[i]._mergeCount = originalRow._cells[i]._mergeCount;
          }
        }
      }

      updatePlaceholders(startRow);
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
          const cellAddress = `${String.fromCharCode(64 + pos.col)}${startRow}`;
          const newCell = worksheet.getCell(cellAddress);
          const originalCell = worksheet.getCell(
            `${String.fromCharCode(64 + pos.col)}${pos.row}`
          );

          const mergeInfo = findMergeRange(originalCell);
          if (mergeInfo) {
            const startColChar = String.fromCharCode(64 + mergeInfo.left);
            const endColChar = String.fromCharCode(64 + mergeInfo.right);
            const newMergeRange = `${startColChar}${startRow}:${endColChar}${startRow}`;
            const startCell = `${startColChar}${startRow}`;
            const endCell = `${endColChar}${startRow}`;
            if (!isRangeAlreadyMerged(startCell, endCell)) {
              worksheet.mergeCells(newMergeRange);
            }
            applyStyleToRange(startCell, endCell, originalCell.style);
          }

          newCell.value = item[key];
          newCell.style = originalCell.style;
        }
      }
    });
  };

  // Traverse the JSON data and call respective functions
  const sheetData = jsonData[0].Sheet1;
  for (const dataKey in sheetData) {
    if (dataKey.startsWith("RepeatedValues_")) {
      const firstPlaceholder = Object.keys(sheetData[dataKey][0])[0];
      if (placeholders[firstPlaceholder]) {
        const startRow = placeholders[firstPlaceholder].row + 1;
        populateMultipleValuePlaceholders(startRow, sheetData[dataKey]);
      }
    } else {
      populateSingleValuePlaceholders(sheetData);
    }
  }

  // Save the updated workbook
  const outputFilePath = path.join(__dirname, "updated_report.xlsx");
  await workbook.xlsx.writeFile(outputFilePath);
  console.log("File saved.");

  console.timeEnd("Report Generation Time"); // End the timer
};

loadTemplateAndPopulateData().catch((error) => {
  console.error("Error processing template:", error);
});
