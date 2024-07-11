const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

// Load JSON data
const jsonFilePath = path.join(__dirname, "./shortJsonData/section3.json");
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, "utf-8"));

// Load the Excel template
const templateFilePath = path.join(__dirname, "template.xlsx");

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

  // Function to convert column letters to numbers
  function columnToNumber(column) {
    let number = 0;
    for (let i = 0; i < column.length; i++) {
      number = number * 26 + (column.charCodeAt(i) - ("A".charCodeAt(0) - 1));
    }
    return number;
  }

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
        if (
          key.startsWith("RepeatedValues_") ||
          key.startsWith("RepeatedValues")
        ) {
          const nestedDataArray = item[key];
          const nestedRows = calculateRowsForRepeatedValues(nestedDataArray);
          maxRowsForValue = Math.max(maxRowsForValue, nestedRows);
        }
      }
      totalRows += maxRowsForValue;
    });
    return totalRows;
  };

  // Calculate the number of rows required for each value in an value of RepeatedValues_$
  const calculateRowsForValues = (item) => {
    let totalRows = 0;
    let maxRowsForValue = 1; // At least one row is required for each value
    for (const key in item) {
      if (
        key.startsWith("RepeatedValues_") ||
        key.startsWith("RepeatedValues")
      ) {
        const nestedDataArray = item[key];
        const nestedRows = calculateRowsForRepeatedValues(nestedDataArray);
        maxRowsForValue = Math.max(maxRowsForValue, nestedRows);
      }
    }
    totalRows += maxRowsForValue;

    return totalRows;
  };

  // Insert empty rows based on calculated row counts
  const insertEmptyRows = (startRow, dataArray) => {
    const numberOfRows = calculateRowsForRepeatedValues(dataArray);
    worksheet.spliceRows(startRow, 0, ...Array(numberOfRows).fill([]));
    updatePlaceholders(startRow, numberOfRows);
    return numberOfRows;
  };

  // Function to apply style to a range of cells
  function applyStyleToRange(startCell, endCell, style) {
    const startCol = columnToNumber(startCell.address.replace(/\d+/, ""));
    const endCol = columnToNumber(endCell.address.replace(/\d+/, ""));
    const startRow = parseInt(startCell.address.match(/\d+/)[0]);
    const endRow = parseInt(endCell.address.match(/\d+/)[0]);

    for (let row = startRow; row <= endRow; row++) {
      for (let col = startCol; col <= endCol; col++) {
        const cell = startCell.worksheet.getCell(row, col);
        cell.style = style;
      }
    }
  }

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
  const populateMultipleValuePlaceholders = (
    startRow,
    dataArray,
    depth = 0
  ) => {
    const isFirstLevel = depth === 0;
    console.log("Depth:", depth);

    dataArray.forEach((item, index) => {
      const newRow = worksheet.getRow(startRow + index);
      // console.log(item,calculateRowsForValues(item));

      for (const key in item) {
        if (
          key.startsWith("RepeatedValues_") ||
          key.startsWith("RepeatedValues")
        ) {
          if (isFirstLevel && index != 0) {
            startRow = startRow + calculateRowsForValues(item);
          }
          const nestedDataArray = item[key];
          const nestedPlaceholder = Object.keys(nestedDataArray[0])[0];
          // const column = pos.cellAddress.match(/[A-Z]+/)[0];
          let nestedStartRow = startRow + index;
          populateMultipleValuePlaceholders(
            nestedStartRow,
            nestedDataArray,
            depth + 1
          );
        }
      }

      for (const key in item) {
        const pos = placeholders[key];
        const currentRow = startRow + index;
        if (pos) {
          const column = pos.cellAddress.match(/[A-Z]+/)[0];
          const mergeRange = pos.mergeRange
            ? pos.mergeRange.match(/[A-Z]+/)[0]
            : null;

          if (pos.isMerged) {
            const mergeRangeStart = column + currentRow;
            const mergeRangeEnd = mergeRange + currentRow;
            worksheet.unMergeCells(mergeRangeStart, mergeRangeEnd);
            worksheet.mergeCells(`${mergeRangeStart}:${mergeRangeEnd}`);

            // Apply style to merged cells
            const startCell = worksheet.getCell(mergeRangeStart);
            const endCell = worksheet.getCell(mergeRangeEnd);
            applyStyleToRange(startCell, endCell, pos.style);
          }
          const cell = worksheet.getCell(column + currentRow);
          cell.value = item[key];
          cell.style = pos.style;
        }
      }
    });
    // Log to check if at the first level of recursion
    if (isFirstLevel) {
      console.log("At the first level of recursion");
    }
  };

  // First, insert all the empty rows
  const sheetData = jsonData[0].Sheet1;
  for (const dataKey in sheetData) {
    if (dataKey.startsWith("RepeatedValues_")) {
      const firstPlaceholder = Object.keys(sheetData[dataKey][0])[0];
      if (placeholders[firstPlaceholder]) {
        const startRow =
          parseInt(placeholders[firstPlaceholder].cellAddress.match(/\d+/)[0]) +
          1;
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
        const startRow =
          parseInt(placeholders[firstPlaceholder].cellAddress.match(/\d+/)[0]) +
          1;
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
