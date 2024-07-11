const ExcelJS = require('exceljs');

// Function to merge multiple Excel files
async function mergeExcelFiles(files, outputFile) {
    const startTime = Date.now();

    // Load the first Excel file
    const workbook1 = new ExcelJS.Workbook();
    await workbook1.xlsx.readFile(files[0]);
    const sheet1 = workbook1.getWorksheet(1);
    let highestRow1 = sheet1.rowCount;

    // Iterate through remaining files
    for (let i = 1; i < files.length; i++) {
        const workbook2 = new ExcelJS.Workbook();
        await workbook2.xlsx.readFile(files[i]);
        const sheet2 = workbook2.getWorksheet(1);

        // Iterate through rows in the current sheet and append to the first sheet
        sheet2.eachRow((row, rowNumber) => {
            const newRow = sheet1.getRow(highestRow1 + rowNumber);
            row.eachCell((cell, colNumber) => {
                newRow.getCell(colNumber).value = cell.value;
            });
            newRow.commit();
        });

        // Update the highest row count for the first sheet
        highestRow1 += sheet2.rowCount;
    }

    // Save merged file
    await workbook1.xlsx.writeFile(outputFile);

    const endTime = Date.now();
    const timeTaken = (endTime - startTime) / 1000;
    console.log(`Files merged successfully in ${timeTaken} seconds!`);
}

// Usage example
const files = [];
for (let i = 1; i <= 25; i++) {
    files.push(`./sectionsForMerging/section${i}.xlsx`);
}
// files.push('./sectionsForMerging/inf1.xlsx'); // Add additional file if needed
const outputFile = './mergeExcel/output.xlsx'; // Output file path

mergeExcelFiles(files, outputFile).catch(console.error);
