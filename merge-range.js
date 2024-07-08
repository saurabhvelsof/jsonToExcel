const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');

// Load the Excel template
const templateFilePath = path.join(__dirname, 'FloodStateCR.xlsx');

const findMergeRange = async (placeholder) => {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(templateFilePath);
    const worksheet = workbook.getWorksheet(1); // Assuming data goes into the first sheet

    // Function to find the merge range of a specific placeholder
    const getMergeRange = (placeholder) => {
        let mergeRange = null;
        worksheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                if (cell.value === placeholder) {
                    const cellAddress = cell.address;
                    worksheet._merges.forEach((mergedCellRange) => {
                        if (mergedCellRange.includes(cellAddress)) {
                            mergeRange = mergedCellRange;
                        }
                    });
                }
            });
        });
        return mergeRange;
    };

    const mergeRange = getMergeRange(placeholder);
    console.log(`Merge range for ${placeholder}: ${mergeRange}`);
};

findMergeRange('[RcAffectedDetails]').catch((error) => {
    console.error('Error finding merge range:', error);
});
