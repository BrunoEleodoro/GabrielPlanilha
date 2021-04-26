const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var moment = require('moment')

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.QUANTIDADE_RECURSOS + 1).value = "Quantidade de Recursos"
        var i = 2
        while (i <= worksheet.rowCount) {
            let time_worked = worksheet.getCell(config.STORE_WORKED_HOURS + i).value;
        
            worksheet.getCell(config.QUANTIDADE_RECURSOS + i).value = parseFloat(time_worked / 160).toFixed(6)
             
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
