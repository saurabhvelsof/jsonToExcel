const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// Load JSON data
const jsonFilePath = path.join(__dirname, "output.json");
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf-8"));

// Load the Excel template
const templateFilePath = path.join(__dirname, "FloodStateCR.xlsx");

const loadTemplateAndPopulateData = async () => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(templateFilePath);
  const worksheet = workbook.getWorksheet(1); // Assuming data goes into the first sheet

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
      //   worksheet.duplicateRow(startRow - 1, (amount = 1), (insert = true));
      // Ensure both rows have the _cells array and they have the same length
      if (
        originalRow._cells &&
        newRow._cells &&
        originalRow._cells.length === newRow._cells.length
      ) {
        for (let i = 1; i < originalRow._cells.length; i++) {
          // Ensure both cells have the _mergeCount property
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
          // copyCellStyles(originalCell, newCell);

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
          }
          newCell.value = item[key];
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
};

loadTemplateAndPopulateData().catch((error) => {
  console.error("Error processing template:", error);
});
