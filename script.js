const csv = require('csvtojson');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

async function convertCSVtoXLSM(csvFileName, xlsmFileName) {
  try {
    const csvFolderPath = path.join(__dirname, 'csv'); // Path to CSV folder
    const xlsxFolderPath = path.join(__dirname, 'xlsx'); // Output path for xlsx folder

    const csvFilePath = path.join(csvFolderPath, csvFileName); // Path to CSV file inside 'csv' folder
    const xlsmFilePath = path.join(xlsxFolderPath, xlsmFileName); // Output path for xlsm file inside 'xlsx' folder

    // Check if the output folder exists, if not, create it
    if (!fs.existsSync(xlsxFolderPath)) {
      fs.mkdirSync(xlsxFolderPath);
    }

    // Convert CSV to JSON
    const jsonArray = await csv().fromFile(csvFilePath);

    // Create a new workbook
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(jsonArray);

    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // Write the workbook to a file as xlsm
    XLSX.writeFile(wb, xlsmFilePath, { bookType: 'xlsm', bookSST: true });
    console.log(`Conversion successful. File saved as ${xlsmFilePath}`);
  } catch (error) {
    console.error('An error occurred:', error);
  }
}


convertCSVtoXLSM('input.csv', 'output_file.xlsm');
