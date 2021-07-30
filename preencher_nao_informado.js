const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

function highlight(worksheet, key) {
    worksheet.getCell(key).fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: { argb: 'FFFFFF00' },
        bgColor: { argb: 'aa00ff' }
    };
}

function verifyClaim(worksheet, key) {
    if (worksheet.getCell(key).value == "NaN") {
        return parseFloat("0")
    }

    return worksheet.getCell(key).value
}


function verifyNa(worksheet, key) {
    if (worksheet.getCell(key).value == null) {
        highlight(worksheet, key)
        return "Nao Informado"
    }

    if (worksheet.getCell(key).value.toString().trim() == "") {
        highlight(worksheet, key)
        return "Nao Informado"
    }

    return worksheet.getCell(key).value
}

function verifyAnalise(worksheet, key) {

    if (worksheet.getCell(key).value == null) {
        // highlight(worksheet, key)
        return "Análise impossível de ser feita"
    }
    if (worksheet.getCell(key).value.toString().trim() == "") {
        // highlight(worksheet, key)
        return "Análise impossível de ser feita"
    }
    if (worksheet.getCell(key).value.toString().toLowerCase().trim() == "nan") {
        // highlight(worksheet, key)
        return "Análise impossível de ser feita"
    }
    /*
    if (key.includes(config.SLA_TICKET_VENCIDO)) {
        console.log(worksheet.getCell(key).value)
    }
    */
    return worksheet.getCell(key).value
}

function verifyTempoAtendimento(worksheet, key) {

    if (worksheet.getCell(key).value == null) {
        // highlight(worksheet, key)
        return "0"
    }
    if (worksheet.getCell(key).value.toString().trim() == "") {
        // highlight(worksheet, key)
        return "0"
    }
    if (worksheet.getCell(key).value.toString().toLowerCase().trim() == "nan") {
        // highlight(worksheet, key)
        return "0"
    }

    if (key.includes(config.SLA_TICKET_VENCIDO)) {
        console.log(worksheet.getCell(key).value)
    }
    return worksheet.getCell(key).value
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);

        var i = 2
        while (i <= worksheet.rowCount) {

            let base_claim = worksheet.getCell(config.BASE_CALCULO_CLAIM + i).value
            worksheet.getCell(config.BASE_CALCULO_CLAIM + i).value = verifyClaim(worksheet, config.BASE_CALCULO_CLAIM + i)
            worksheet.getCell(config.CLAIM + i).value = verifyClaim(worksheet, config.CLAIM + i)
            worksheet.getCell(config.CLIENTS_COLUMN + i).value = verifyNa(worksheet, config.CLIENTS_COLUMN + i)
            worksheet.getCell(config.STORE_TYPE_COLUMN + i).value = verifyNa(worksheet, config.STORE_TYPE_COLUMN + i)
            worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = verifyNa(worksheet, config.STORE_SEVERITY_COLUNM + i)
            worksheet.getCell(config.PROBLEMA_REPORTADO + i).value = verifyNa(worksheet, config.PROBLEMA_REPORTADO + i)
            worksheet.getCell(config.CATEGORIA + i).value = verifyNa(worksheet, config.CATEGORIA + i)
            worksheet.getCell(config.SERVICE_LINE + i).value = verifyNa(worksheet, config.SERVICE_LINE + i)
            worksheet.getCell(config.TRIBE + i).value = verifyNa(worksheet, config.TRIBE + i)

            worksheet.getCell(config.HORARIO_INCIDENTE + i).value = verifyNa(worksheet, config.HORARIO_INCIDENTE + i)
            worksheet.getCell(config.SLA_TICKET + i).value = verifyNa(worksheet, config.SLA_TICKET + i)
            worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = verifyNa(worksheet, config.HORARIO_ACIONAMENTO + i)
            worksheet.getCell(config.ISM_SOLICITOU + i).value = verifyNa(worksheet, config.ISM_SOLICITOU + i)

            worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = verifyAnalise(worksheet, config.SLA_TICKET_VENCIDO + i)
            worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = verifyTempoAtendimento(worksheet, config.TEMPO_ATENDIMENTO + i)
            worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).value = verifyAnalise(worksheet, config.ANALISE_PRAZO_ACIONAMENTO + i)
            if (worksheet.getCell(config.STORE_TYPE_COLUMN + i).value == "Nao Informado") {
                highlight(worksheet, config.STORE_TYPE_COLUMN + i);
            }

            var severidade = worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value
            var type = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value
            var label = "";

            if (severidade == "N/A") {
                worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).fill = undefined;
                worksheet.getCell(config.HORARIO_INCIDENTE + i).fill = undefined;
                worksheet.getCell(config.SLA_TICKET + i).fill = undefined;
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).fill = undefined;
                worksheet.getCell(config.HORARIO_ACIONAMENTO + i).fill = undefined;
                worksheet.getCell(config.ISM_SOLICITOU + i).fill = undefined;
                worksheet.getCell(config.TEMPO_ATENDIMENTO + i).fill = undefined;
                worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).fill = undefined;
                if (type == "CH") {
                    label = "N/A - CHANGE"
                } else if (type == "REPORT") {
                    label = "N/A - REPORT"
                } else if (type == "SC") {
                    label = "N/A - SEM CHAMADO"
                }
            }
            if (label != "") {
                worksheet.getCell(config.STORE_SEVERITY_COLUNM + i).value = label
                if (label == "N/A - SEM CHAMADO") {
                    worksheet.getCell(config.HORARIO_INCIDENTE + i).value = label
                    worksheet.getCell(config.SLA_TICKET + i).value = label
                    worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).value = label
                }
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = label
                // worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = label
                // worksheet.getCell(config.ISM_SOLICITOU + i).value = label
                worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = parseFloat("0")


            }
            if (severidade.toString().includes("N/A")) {
                let valor = "";
                if (type == "CH") {
                    valor = "N/A - CHANGE"
                } else if (type == "REPORT") {
                    valor = "N/A - REPORT"
                } else if (type == "SC") {
                    valor = "N/A - SEM CHAMADO"
                }
                worksheet.getCell(config.HORARIO_INCIDENTE + i).value = valor
                worksheet.getCell(config.SLA_TICKET + i).value = valor
                worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = valor
                worksheet.getCell(config.ISM_SOLICITOU + i).value = valor
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = valor
                worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = valor
                worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).value = valor
                worksheet.getCell(config.HORARIO_ENCERRAMENTO + i).value = valor
            }

            // Horário do Incident,
            // SLA do ticket
            // Horário Acionamento ISM
            // ISM solicitou Validação
            // Tempo de Atendimento ISM
            // Análise do Prazo de Acionamento ISM
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
