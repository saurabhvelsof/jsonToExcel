const ExcelJS = require('exceljs');
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('Sheet1');

// Load your existing template Excel file
workbook.xlsx.readFile('./new_template.xlsx').then(() => {
    const sheet = workbook.getWorksheet('Sheet1');

    // Define where you want to insert the new row (between Row 1 and Row 2)
    const rowIndexToInsert = 2; // Inserting after Row 1 (zero-indexed)

    // Insert a new row at the specified index
    const newRow = sheet.addRow([], rowIndexToInsert); // Add empty array for row values

    // Copy the merge from Row 1 (A1:B1) to the new row (A2:B2)
    const mergeRange = sheet.getCell('A1').model.master; // Get merge information from A1 (assuming A1:B1 is merged)

    if (mergeRange) {
        // Calculate merge range for the new row (A2:B2)
        const newMerge = {
            master: { ...mergeRange },
            ranges: [sheet.getCell(`A${rowIndexToInsert + 1}`).model],
        };

        // Apply merge to the new row
        sheet.model.merges.push(newMerge);
    }

    // Save the modified workbook
    return workbook.xlsx.writeFile('output.xlsx');
}).then(() => {
    console.log('Row inserted and merged successfully!');
}).catch((error) => {
    console.error('Error inserting row and merging cells:', error);
});
