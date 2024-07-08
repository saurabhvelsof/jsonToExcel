const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Load JSON data
const jsonFilePath = path.join(__dirname, 'output.json');
const jsonData = JSON.parse(fs.readFileSync(jsonFilePath, 'utf-8'));

// Load the Excel template
const templateFilePath = path.join(__dirname, 'FloodStateCR.xlsx');

const loadTemplateAndPopulateData = async () => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templateFilePath);
    const worksheet = workbook.getWorksheet(1); // Assuming data goes into the first sheet

    // Function to find placeholders and return their row and column indices
    const findPlaceholders = () => {
        const placeholders = {};
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                const cellValue = cell.value;
                if (typeof cellValue === 'string' && cellValue.startsWith('[') && cellValue.endsWith(']')) {
                    placeholders[cellValue] = { row: rowNumber, col: colNumber };
                }
            });
        });
        return placeholders;
    };

    const placeholders = findPlaceholders();

    // Replace placeholders with JSON data
    Object.keys(placeholders).forEach((key) => {
        const { row, col } = placeholders[key];
        if (jsonData[0].Sheet1[key] !== undefined) {
            worksheet.getRow(row).getCell(col).value = jsonData[0].Sheet1[key];
        }
    });

    // Helper function to populate nested data and merge cells
    const populateNestedData = (startRow, dataKey, mergeColumns) => {
        const data = jsonData[0].Sheet1[dataKey];
        if (data && data.length > 0) {
            let currentRow = startRow;
            data.forEach((item, index) => {
                const row = worksheet.insertRow(currentRow, Object.values(item));
                mergeColumns.forEach(col => {
                    const startCell = worksheet.getCell(`${col}${currentRow}`);
                    const endCell = worksheet.getCell(`${col}${currentRow + data.length - 1}`);
                    if (!worksheet.getCell(startCell.address).isMerged && !worksheet.getCell(endCell.address).isMerged) {
                        worksheet.mergeCells(startCell.address, endCell.address);
                    }
                });
                currentRow++;
            });
        }
    };

    // Define a mapping from data keys to their respective starting placeholders
    const nestedDataMapping = {
        'RepeatedValues_0': { placeholder: '[RepeatedValues_0_Start]', mergeColumns: ['A'] },
        'RepeatedValues_1': { placeholder: '[RepeatedValues_1_Start]', mergeColumns: ['A'] },
        'RepeatedValues_2': { placeholder: '[RepeatedValues_2_Start]', mergeColumns: ['A'] },
        'RepeatedValues_3': { placeholder: '[RepeatedValues_3_Start]', mergeColumns: ['A'] },
        'RepeatedValues_4': { placeholder: '[RepeatedValues_4_Start]', mergeColumns: ['A'] }
    };

    // Populate nested data sections dynamically
    Object.keys(nestedDataMapping).forEach(dataKey => {
        const { placeholder, mergeColumns } = nestedDataMapping[dataKey];
        if (placeholders[placeholder]) {
            const startRow = placeholders[placeholder].row + 1;
            populateNestedData(startRow, dataKey, mergeColumns);
        }
    });

    // Save the updated workbook
    const outputFilePath = path.join(__dirname, 'updated_report.xlsx');
    await workbook.xlsx.writeFile(outputFilePath);
    console.log('File saved.');
};

loadTemplateAndPopulateData().catch((error) => {
    console.error('Error processing template:', error);
});
