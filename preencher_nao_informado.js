const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);

        var i = 2
        while (i <= worksheet.rowCount) {

            // let base_claim = worksheet.getCell(config.BASE_CALCULO_CLAIM + i).value
            worksheet.getCell(config.CLIENTS_COLUMN + i).value = worksheet.getCell(config.CLIENTS_COLUMN + i).value || "Nao Informado"
            worksheet.getCell(config.STORE_TYPE_COLUMN + i).value = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value || "Nao Informado"
            worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value || "Nao Informado"
            worksheet.getCell(config.PROBLEMA_REPORTADO + i).value = worksheet.getCell(config.PROBLEMA_REPORTADO + i).value || "Nao Informado"
            worksheet.getCell(config.CATEGORIA + i).value = worksheet.getCell(config.CATEGORIA + i).value || "Nao Informado"
            worksheet.getCell(config.SERVICE_LINE + i).value = worksheet.getCell(config.SERVICE_LINE + i).value || "Nao Informado"
            worksheet.getCell(config.HORARIO_INCIDENTE + i).value = worksheet.getCell(config.HORARIO_INCIDENTE + i).value || "Nao Informado"
            worksheet.getCell(config.SLA_TICKET + i).value = worksheet.getCell(config.SLA_TICKET + i).value || "Nao Informado"
            worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value || "Nao Informado"
            worksheet.getCell(config.ISM_SOLICITOU + i).value = worksheet.getCell(config.ISM_SOLICITOU + i).value || "Nao Informado"

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
