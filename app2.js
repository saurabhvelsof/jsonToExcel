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
                    placeholders[cellValue] = { row: rowNumber, col: colNumber };
                }
            });
        });
        return placeholders;
    };

    let placeholders = findPlaceholders();

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

    // Function to check if a range is already merged
    const isRangeAlreadyMerged = (worksheet, startCell, endCell) => {
        for (const key in worksheet._merges) {
            const merge = worksheet._merges[key].model;
            const mergeStartCell = `${String.fromCharCode(64 + merge.left)}${merge.top}`;
            const mergeEndCell = `${String.fromCharCode(64 + merge.right)}${merge.bottom}`;
            if ((startCell >= mergeStartCell && startCell <= mergeEndCell) ||
                (endCell >= mergeStartCell && endCell <= mergeEndCell)) {
                return true;
            }
        }
        return false;
    };

    // Function to update placeholders' positions after row insertion
    const updatePlaceholders = (insertedRow) => {
        Object.keys(placeholders).forEach(key => {
            if (placeholders[key].row >= insertedRow) {
                placeholders[key].row += 1;
            }
        });
    };

    // Function to copy styles from the placeholder cell to the new cell
    const copyCellStyles = (sourceCell, targetCell) => {
        targetCell.style = JSON.parse(JSON.stringify(sourceCell.style));
    };

    // Function to populate non-repeated data
    const populateNonRepeatedData = (dataKey, data) => {
        const { row, col } = placeholders[dataKey];
        worksheet.getRow(row).getCell(col).value = data[dataKey];
    };

    // Function to populate repeated (nested) data
    const populateRepeatedData = (startRow, dataArray) => {
        dataArray.forEach((item, index) => {
            const newRow = worksheet.insertRow(startRow + index);
            updatePlaceholders(startRow + index); // Update placeholder positions after insertion

            Object.keys(item).forEach(placeholder => {
                const pos = placeholders[placeholder];
                if (pos) {
                    const cellAddress = `${String.fromCharCode(64 + pos.col)}${startRow + index}`;
                    const newCell = worksheet.getCell(cellAddress);
                    const originalCell = worksheet.getCell(`${String.fromCharCode(64 + pos.col)}${pos.row}`);
                    newCell.value = item[placeholder];
                    copyCellStyles(originalCell, newCell);

                    // Handle merged cells
                    const mergeInfo = findMergeRange(`${String.fromCharCode(64 + pos.col)}${pos.row}`);
                    if (mergeInfo) {
                        const startColChar = String.fromCharCode(64 + mergeInfo.startCol);
                        const endColChar = String.fromCharCode(64 + mergeInfo.endCol);
                        const newMergeRange = `${startColChar}${startRow + index}:${endColChar}${startRow + index}`;

                        // Check if the new merge range is already merged
                        const startCell = `${startColChar}${startRow + index}`;
                        const endCell = `${endColChar}${startRow + index}`;
                        if (!isRangeAlreadyMerged(worksheet, startCell, endCell)) {
                            worksheet.mergeCells(newMergeRange);
                        }
                    }
                }
            });

            // Recursively handle nested repeated values
            Object.keys(item).forEach(key => {
                if (key.startsWith('RepeatedValues_')) {
                    const nestedDataArray = item[key];
                    populateRepeatedData(startRow + index, nestedDataArray);
                }
            });
        });
    };

    // Traverse the JSON data and call respective functions
    Object.keys(jsonData[0].Sheet1).forEach(dataKey => {
        if (dataKey.startsWith('RepeatedValues_')) {
            const firstPlaceholder = Object.keys(jsonData[0].Sheet1[dataKey][0])[0];
            if (placeholders[firstPlaceholder]) {
                const startRow = placeholders[firstPlaceholder].row + 1;
                populateRepeatedData(startRow, jsonData[0].Sheet1[dataKey]);
            }
        } else {
            populateNonRepeatedData(dataKey, jsonData[0].Sheet1);
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
