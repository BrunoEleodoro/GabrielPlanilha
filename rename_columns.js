const utf8 = require('utf8');
// import *  as config from './load_columns';
const config = require('./load_columns');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        //setting the title of the column
        
        worksheet.getCell(config.HORARIO_INCIDENTE + 1).value = "Horario do Incidente"
        worksheet.getCell(config.SLA_TICKET + 1).value = "SLA do Ticket"
        worksheet.getCell(config.HORARIO_ACIONAMENTO + 1).value = "Horario Acionamento ISM "
        worksheet.getCell(config.ISM_SOLICITOU + 1).value = "ISM Solicitou Validacao?"

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
