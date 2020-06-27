const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

function verifyNa(worksheet, key) {
    if (worksheet.getCell(key).value == null) {
        return "Nao Informado"
    }

    if (worksheet.getCell(key).value.toString().trim() == "") {
        return "Nao Informado"
    }

    return worksheet.getCell(key).value
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);

        var i = 2
        while (i <= worksheet.rowCount) {

            // let base_claim = worksheet.getCell(config.BASE_CALCULO_CLAIM + i).value
            worksheet.getCell(config.CLIENTS_COLUMN + i).value = verifyNa(worksheet, config.CLIENTS_COLUMN + i)
            worksheet.getCell(config.STORE_TYPE_COLUMN + i).value = verifyNa(worksheet, config.STORE_TYPE_COLUMN + i)
            worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = verifyNa(worksheet, config.STORE_SEVERITY_COLUNM + i)
            worksheet.getCell(config.PROBLEMA_REPORTADO + i).value = verifyNa(worksheet, config.PROBLEMA_REPORTADO + i)
            worksheet.getCell(config.CATEGORIA + i).value = verifyNa(worksheet, config.CATEGORIA + i)
            worksheet.getCell(config.SERVICE_LINE + i).value = verifyNa(worksheet, config.SERVICE_LINE + i)
            worksheet.getCell(config.HORARIO_INCIDENTE + i).value = verifyNa(worksheet, config.HORARIO_INCIDENTE + i)
            worksheet.getCell(config.SLA_TICKET + i).value = verifyNa(worksheet, config.SLA_TICKET + i)
            worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = verifyNa(worksheet, config.HORARIO_ACIONAMENTO + i)
            worksheet.getCell(config.ISM_SOLICITOU + i).value = verifyNa(worksheet, config.ISM_SOLICITOU + i)
            worksheet.getCell(config.TRIBE + i).value = verifyNa(worksheet, config.TRIBE + i)

            var severidade = worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value
            var type = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value

            if (severidade == "N/A") {
                if (type == "CH") {
                    worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = "N/A - CHANGE"
                } else if (type == "REPORT") {
                    worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = "N/A - REPORT"
                } else if (type == "SC") {
                    worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = "N/A - SEM CHAMADO"
                }
            }
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
