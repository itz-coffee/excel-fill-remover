const Excel = require("exceljs");
const readline = require('readline').createInterface({
    input: process.stdin,
    output: process.stdout
});

async function processWorkbook(filePath) {
    const workbook = new Excel.Workbook();
    await workbook.xlsx.readFile(`./${filePath}.xlsx`);

    const worksheet = workbook.worksheets[0];

    worksheet.eachRow((row) => {
        row.eachCell((cell) => {
            if (cell.fill.pattern == "solid") {
                cell.fill = {
                    type: "pattern",
                    pattern: "none"
                };
            }
        });
    });

    await workbook.xlsx.writeFile(`./${filePath}Processed.xlsx`);
    console.log(`${filePath} was processed!`);
}

readline.question("Workbook name: ", filePath => {
    processWorkbook(filePath);
    readline.close();
});
