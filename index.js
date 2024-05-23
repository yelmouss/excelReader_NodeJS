const ExcelJS = require("exceljs");

const workbook = new ExcelJS.Workbook();
const startTime = performance.now();

workbook.xlsx
  .readFile("10000.xlsx")
  .then(() => {
    const worksheet = workbook.getWorksheet("Feuille 1");

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      console.log(`${row.values[1]} travaille sur la zone ${row.values[2]}`);
    });

    const endTime = performance.now();
    const timeTaken = endTime - startTime;
    console.log(`Time taken to read the file: ${timeTaken} milliseconds`);
  })
  .catch((error) => {
    console.error("Error reading Excel file:", error);
  });