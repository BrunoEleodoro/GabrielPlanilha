const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// Se(Horário do Incident > SLA do Ticket,
// "Solicitado Prioridade com SLA Vencido",
// "Solicitado Prioridade Dentro do SLA"). 

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.SLA_TICKET_VENCIDO + 1).value = "SLA do Ticket Vencido?"
        var i = 2
        while (i <= worksheet.rowCount) {
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var sla_ticket = worksheet.getCell(config.SLA_TICKET + i).value

            var horario_acionamento_date = new Date(horario_acionamento);
            var sla_ticket_date = new Date(sla_ticket);

            console.log(`(${horario_acionamento_date} > ${sla_ticket_date}`, (horario_acionamento_date > sla_ticket_date))

            if (horario_acionamento_date > sla_ticket_date) {
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = "Solicitado Prioridade com SLA Vencido"
            } else {
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = "Solicitado Prioridade Dentro do SLA"
            }

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
