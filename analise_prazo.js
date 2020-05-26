const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

var moment = require('moment')

var dateFormat1 = 'DD/MM/YYYY HH:mm';
var dateFormat2 = 'MM/DD/YYYY HH:mm';

function calculateHours(startDate, startDateFormat1, startDateFormat2, endDate, endDateFormat1, endDateFormat2) {
    var start_date_moment = moment(startDate, startDateFormat1);
    if (start_date_moment.toString() == "Invalid date") {
        start_date_moment = moment(startDate, startDateFormat2);
    }
    var end_date_moment = moment(endDate, endDateFormat1);
    if (end_date_moment.toString() == "Invalid date") {
        end_date_moment = moment(endDate, endDateFormat2);
    }
    var duration = moment.duration(end_date_moment.diff(start_date_moment));
    var hours = duration.asHours();
    // hours = moment(hours * 3600 * 1000).format('HH:mm')
    return hours;
}


// return;

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + 1).value = "Análise do Prazo de Acionamento ISM"
        var i = 2
        while (i <= worksheet.rowCount) {

            let horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            let sla_ticket = worksheet.getCell(config.SLA_TICKET + i).value
            let horario_incident = worksheet.getCell(config.HORARIO_INCIDENTE + i).value

            let sla_horario_acionamento = calculateHours(
                sla_ticket,
                "DD/MM/YYYY HH:mm",
                "MM/DD/YYYY HH:mm",
                horario_acionamento,
                "DD/MM/YYYY HH:mm",
                "MM/DD/YYYY HH:mm");

            let sla_horario_incident = calculateHours(
                sla_ticket,
                "DD/MM/YYYY HH:mm",
                "MM/DD/YYYY HH:mm",
                horario_incident,
                "MM/DD/YYYY HH:mm",
                "DD/MM/YYYY HH:mm");

            var horario_acionamento_incident = sla_horario_acionamento / sla_horario_incident

            var res = "";
            if (horario_acionamento_incident >= 0.90) {
                res = "Solicitado prioridade no momento zero"
            } else if (horario_acionamento_incident >= 0.50) {
                res = "Solicitado prioridade dentro do SLA acordado com ISM"
            } else if (horario_acionamento_incident >= 0.10) {
                res = "Solicitado prioridade próximo do SLA vencer"
            } else {
                res = "Solicitado prioridade com SLA vencido";
            }
            worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).value = res

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
