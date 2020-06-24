const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var moment = require('moment')

// Se(Horário do Incident > SLA do Ticket,
// "Solicitado Prioridade com SLA Vencido",
// "Solicitado Prioridade Dentro do SLA"). 
var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function calculateHours(startDate, endDate) {

    var duration = moment.duration(endDate.diff(startDate));
    var hours = duration.asHours();
    // hours = moment(hours * 3600 * 1000).format('HH:mm')
    return hours;
}

// return;

function parseDateToMoment(monthName, valor_celula) {
    var month = months.indexOf(monthName) + 1
    if (valor_celula == "" || valor_celula == null) {
        return moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    }
    // var index = valor_celula.indexOf("0" + month + "/")
    var index = valor_celula.split(" ")[0].indexOf(month)
    var data = ""

    if (index == -1) {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    } else if (index == 0) {
        data = moment(valor_celula, "MM/DD/YYYY HH:mm");
    } else if (index == 3) {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    } else if (index == 4) {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    } else {
        data = moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    }
    if (data.toString() == "Invalid date") {
        data = moment(valor_celula, "MM/DD/YYYY HH:mm");
    }

    return data
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.SLA_TICKET_VENCIDO + 1).value = "SLA do Ticket Vencido?"
        var i = 2
        while (i <= worksheet.rowCount) {
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var sla_ticket = worksheet.getCell(config.SLA_TICKET + i).value
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value

            var horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
            var sla_ticket_date = parseDateToMoment(monthName, sla_ticket);

            // console.log(`(${horario_acionamento_date} > ${sla_ticket_date}`, (horario_acionamento_date > sla_ticket_date))


            // if (horario_acionamento_date > sla_ticket_date) {
            // console.log('horario_acionamento_date', horario_acionamento_date, sla_ticket_date, horario_acionamento_date.isAfter(sla_ticket_date, 'hour'))

            if (horario_acionamento_date.toString() != "Invalid date" && horario_acionamento_date.toString() != "Invalid date") {
                if (horario_acionamento_date.isAfter(sla_ticket_date, 'seconds')) {
                    worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = "Solicitado Prioridade com SLA Vencido"
                } else {
                    worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = "Solicitado Prioridade Dentro do SLA"
                }
            } else {
                worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value = "Nan"
            }

            // if (i % 10 == 0) {
            //     break;
            // }
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
