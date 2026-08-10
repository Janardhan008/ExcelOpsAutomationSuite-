const ExcelJs = require('exceljs');

async function excelTest()
{
   const workbook = new ExcelJs.Workbook();
   await workbook.xlsx.readFile("C:\\Users\\janar\\Downloads\\download.xlsx");
   const worksheet = workbook.getWorksheet('Sheet1');
   worksheet.eachRow((row, rowNumber) => {
        // ... (folded - likely the same row/cell iteration & logging from before)
   })

   const cell = worksheet.getCell(3, 2);
   cell.value = "Iphone";
   await workbook.xlsx.writeFile("C:\\Users\\janar\\Downloads\\download.xlsx")
}

excelTest();