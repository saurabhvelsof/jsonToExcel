const ExcelJS = require('exceljs');
const fs = require('fs');

// File paths
const templateFilePath = './template_short.xlsx';
const jsonDataFilePath = './short_json_data.json';
const outputFilePath = './generated_excel.xlsx';

// Load the JSON data
const jsonData = JSON.parse(fs.readFileSync(jsonDataFilePath, 'utf8'));

// Create a new workbook
const workbook = new ExcelJS.Workbook();

// Load the template workbook
workbook.xlsx.readFile(templateFilePath).then(() => {
  const sheet = workbook.getWorksheet('Sheet1');
  const data = jsonData[0]["Sheet1"];

  // Function to replace placeholders in a row
  const replacePlaceholders = (row, data) => {
    row.eachCell((cell, colNumber) => {
      const cellValue = cell.value;
      if (typeof cellValue === 'string') {
        const placeholder = cellValue.match(/\[.*?\]/);
        if (placeholder) {
          const key = placeholder[0];
          if (data[key] !== undefined) {
            cell.value = data[key];
          }
        }
      }
    });
  };

  // Function to insert repeated values recursively
  const insertRepeatedValues = (sheet, rowNumber, repeatedData) => {
    repeatedData.forEach(dataRow => {
      sheet.insertRow(rowNumber, []);
      const row = sheet.getRow(rowNumber);
      replacePlaceholders(row, dataRow);
      row.commit();
      rowNumber++;

      // Check for nested repeated values
      Object.keys(dataRow).forEach(key => {
        if (key.startsWith('RepeatedValues')) {
          rowNumber = insertRepeatedValues(sheet, rowNumber, dataRow[key]);
        }
      });
    });
    return rowNumber;
  };

  // Populate single values
  sheet.eachRow((row, rowNumber) => {
    replacePlaceholders(row, data);
  });

  // Populate repeated values
  Object.keys(data).forEach(key => {
    if (key.startsWith('RepeatedValues_')) {
      let rowNumber = sheet.actualRowCount + 1;
      rowNumber = insertRepeatedValues(sheet, rowNumber, data[key]);
    }
  });

  // Write the workbook to a file
  workbook.xlsx.writeFile(outputFilePath).then(() => {
    console.log('Excel file generated successfully.');
  });
});
