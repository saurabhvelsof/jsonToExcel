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

    // Function to find placeholders and their positions
    const findPlaceholders = () => {
        const placeholders = {};
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                const cellValue = cell.value;
                if (typeof cellValue === 'string' && cellValue.startsWith('[') && cellValue.endsWith(']')) {
                    if (!placeholders[cellValue]) {
                        placeholders[cellValue] = { row: rowNumber, col: colNumber };
                    }
                }
            });
        });
        return placeholders;
    };

    const placeholders = findPlaceholders();

    // Function to find the merge range for a given cell address
    const findMergeRange = (cellAddress) => {
        const cell = worksheet.getCell(cellAddress);
        const row = cell.row;
        const col = cell.col;
        for (const key in worksheet._merges) {
            const merge = worksheet._merges[key].model;
            if (row >= merge.top && row <= merge.bottom && col >= merge.left && col <= merge.right) {
                return {
                    startCol: merge.left,
                    endCol: merge.right,
                    startRow: merge.top,
                    endRow: merge.bottom,
                };
            }
        }
        return null;
    };

    // Replace placeholders with JSON data
    Object.keys(placeholders).forEach((key) => {
        const { row, col } = placeholders[key];
        if (jsonData[0].Sheet1[key] !== undefined) {
            worksheet.getRow(row).getCell(col).value = jsonData[0].Sheet1[key];
        }
    });

    // Helper function to populate nested data and handle merged cells
    const populateNestedData = (startRow, dataKey, placeholders) => {
        const data = jsonData[0].Sheet1[dataKey];
        if (data && data.length > 0) {
            data.forEach((item, index) => {
                worksheet.insertRow(startRow + index);
                Object.keys(item).forEach(placeholder => {
                    const pos = placeholders[placeholder];
                    if (pos) {
                        const cellAddress = `${String.fromCharCode(64 + pos.col)}${startRow + index}`;
                        worksheet.getCell(cellAddress).value = item[placeholder];

                        // Handle merged cells
                        const mergeInfo = findMergeRange(`${String.fromCharCode(64 + pos.col)}${pos.row}`);
                        if (mergeInfo) {
                            const startColChar = String.fromCharCode(64 + mergeInfo.startCol);
                            const endColChar = String.fromCharCode(64 + mergeInfo.endCol);
                            const newMergeRange = `${startColChar}${startRow + index}:${endColChar}${startRow + index}`;
                            worksheet.mergeCells(newMergeRange);
                        }
                    }
                });
            });
        }
    };

    // Infer nested data keys from JSON data
    Object.keys(jsonData[0].Sheet1).forEach(dataKey => {
        if (dataKey.startsWith('RepeatedValues_')) {
            const firstPlaceholder = Object.keys(jsonData[0].Sheet1[dataKey][0])[0];
            if (placeholders[firstPlaceholder]) {
                const startRow = placeholders[firstPlaceholder].row + 1;
                populateNestedData(startRow, dataKey, placeholders);
            }
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
