const axios = require('axios');
const FormData = require('form-data');
const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');

// Ensure the apiCallsReports directory exists
const reportsDir = path.join(__dirname, 'apiCallsReports2');
if (!fs.existsSync(reportsDir)) {
  fs.mkdirSync(reportsDir);
}

// Dynamically generate file paths for templates
const templatePaths = Array.from({ length: 24 }, (_, i) => 
  path.join(__dirname, 'shortTemplates', `section${i + 1}.xlsx`)
);

// Single JSON path for all templates
const jsonPath = path.join(__dirname, 'shortJsonData', 'output.json');

// Function to create FormData
const createFormData = (templatePath, jsonPath) => {
  const data = new FormData();
  data.append('template', fs.createReadStream(templatePath));
  data.append('json', fs.createReadStream(jsonPath));
  return data;
};

// Function to make API call and measure time taken
const measureApiCallTime = async (url, templatePath, jsonPath, outputFilePath) => {
  const label = `Time taken for ${templatePath} and ${jsonPath}`;
  console.time(label);

  const data = createFormData(templatePath, jsonPath);
  const config = {
    method: 'post',
    maxBodyLength: Infinity,
    url: url,
    headers: { 
      ...data.getHeaders()
    },
    responseType: 'stream', // Important for handling binary data
    data: data
  };

  try {
    const response = await axios.request(config);
    const writer = fs.createWriteStream(outputFilePath);

    response.data.pipe(writer);

    return new Promise((resolve, reject) => {
      writer.on('finish', () => {
        console.timeEnd(label);
        resolve({ templatePath, jsonPath, outputFilePath, status: 'success' });
      });
      writer.on('error', (error) => {
        console.timeEnd(label);
        reject({ templatePath, jsonPath, outputFilePath, error: error.message });
      });
    });
  } catch (error) {
    console.timeEnd(label);
    return { templatePath, jsonPath, outputFilePath, error: error.message };
  }
};

// Function to make multiple API calls simultaneously
const makeApiCalls = async (url, templatePaths, jsonPath) => {
  console.time('Overall Time');

  const promises = templatePaths.map((templatePath, index) => {
    const outputFilePath = path.join(reportsDir, `output_${index + 1}.xlsx`);
    return measureApiCallTime(url, templatePath, jsonPath, outputFilePath);
  });

  const results = await Promise.all(promises);

  console.timeEnd('Overall Time');
  console.log('API Call Results:', results);

  // Check for any errors in the results
  const failedResults = results.filter(result => result.status !== 'success');
  if (failedResults.length > 0) {
    console.error('Some API calls failed:', failedResults);
    return;
  }

  // Merge all Excel files
  mergeExcelFiles(results.map(result => result.outputFilePath));
};

// Function to merge Excel files into a single sheet
const mergeExcelFiles = (filePaths) => {
  const outputWorkbook = XLSX.utils.book_new();
  const outputSheet = [];
  let currentRow = 0;

  filePaths.forEach((filePath, index) => {
    if (fs.existsSync(filePath)) {
      const workbook = XLSX.readFile(filePath);
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const range = XLSX.utils.decode_range(sheet['!ref']);
      for (let R = range.s.r; R <= range.e.r; ++R) {
        const row = [];
        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cell_address = { c: C, r: R };
          const cell_ref = XLSX.utils.encode_cell(cell_address);
          row.push(sheet[cell_ref] ? sheet[cell_ref].v : '');
        }
        outputSheet.push(row);
      }
    } else {
      console.error(`File not found: ${filePath}`);
    }
  });

  const worksheet = XLSX.utils.aoa_to_sheet(outputSheet);
  XLSX.utils.book_append_sheet(outputWorkbook, worksheet, 'MergedSheet');

  const outputFilePath = path.join(reportsDir, 'merged_output.xlsx');
  XLSX.writeFile(outputWorkbook, outputFilePath);
  console.log(`Merged file created at ${outputFilePath}`);
};

// URL for the API call
const apiUrl = 'http://139.59.40.47/api/Reporting/GenerateExcelMultiLevel';

makeApiCalls(apiUrl, templatePaths, jsonPath);
