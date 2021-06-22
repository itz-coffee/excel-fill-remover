const Excel = require("exceljs");
const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

async function processWorkbook(filePath) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(filePath);

    workbook.eachSheet((sheet) => {
        sheet.eachRow((row) => {
            row.eachCell((cell) => {
                if (cell.fill.pattern == "solid") {
                    cell.fill = {
                        type: "pattern",
                        pattern: "none"
                    };
                }
            });
        });

        sheet.pageSetup.fitToPage = true;
        sheet.pageSetup.fitToWidth = 1;
        sheet.pageSetup.fitToHeight = 0;
    });

    await workbook.xlsx.writeFile(`${filePath}Processed.xlsx`);
    console.log(`${filePath} was processed!`);
}

readline.question("Drag and drop workbook and press [ENTER]: ", filePath => {
    processWorkbook(filePath);
    readline.close();
});
