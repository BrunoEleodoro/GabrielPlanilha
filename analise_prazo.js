const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

var moment = require('moment')

var dateFormat1 = 'DD/MM/YYYY HH:mm';
var dateFormat2 = 'MM/DD/YYYY HH:mm';

// function calculateHours(startDate, startDateFormat1, startDateFormat2, endDate, endDateFormat1, endDateFormat2) {
//     var start_date_moment = moment(startDate, startDateFormat1);
//     if (start_date_moment.toString() == "Invalid date") {
//         start_date_moment = moment(startDate, startDateFormat2);
//     }
//     var end_date_moment = moment(endDate, endDateFormat1);
//     if (end_date_moment.toString() == "Invalid date") {
//         end_date_moment = moment(endDate, endDateFormat2);
//     }
//     var duration = moment.duration(end_date_moment.diff(start_date_moment));
//     var hours = duration.asHours();
//     // hours = moment(hours * 3600 * 1000).format('HH:mm')
//     return hours;
// }
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
    } else if (index == 1) {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    } else {
        data = moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    }
    if (data.toString() == "Invalid date") {
        data = moment(valor_celula, "MM/DD/YYYY HH:mm");
    }
    if (data.toString() == "Invalid date") {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    }

    return data
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + 1).value = "Análise do Prazo de Acionamento ISM"
        var i = 2
        while (i <= worksheet.rowCount) {

            let horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            let sla_ticket = worksheet.getCell(config.SLA_TICKET + i).value
            let horario_incident = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var sla_ticket_vencido = worksheet.getCell(config.SLA_TICKET_VENCIDO + i).value
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value

            var sla_ticket_date = parseDateToMoment(monthName, sla_ticket);
            var horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
            var horario_incident_date = parseDateToMoment(monthName, horario_incident)

            // console.log('sla_ticket_date', sla_ticket_date);
            // console.log('horario_acionamento_date', horario_acionamento_date);
            // console.log('horario_incident_date', horario_incident_date);

            var sla_horario_acionamento = calculateHours(sla_ticket_date, horario_acionamento_date)
            var sla_horario_incident = calculateHours(sla_ticket_date, horario_incident_date);

            //CALCULAR O HORARIO DE ACIONAMENTO
            // data = moment.duration(end_date_moment.diff(start_date_moment))
            // let sla_horario_acionamento = calculateHours(
            //     sla_ticket,
            //     "DD/MM/YYYY HH:mm",
            //     "MM/DD/YYYY HH:mm",
            //     horario_acionamento,
            //     "DD/MM/YYYY HH:mm",
            //     "MM/DD/YYYY HH:mm");

            // let sla_horario_incident = calculateHours(
            //     sla_ticket,
            //     "DD/MM/YYYY HH:mm",
            //     "MM/DD/YYYY HH:mm",
            //     horario_incident,
            //     "MM/DD/YYYY HH:mm",
            //     "DD/MM/YYYY HH:mm");
            if (sla_horario_acionamento < 0) {
                sla_horario_acionamento = sla_horario_acionamento * -1
            }
            if (sla_horario_incident < 0) {
                sla_horario_incident = sla_horario_incident * -1
            }
            var horario_acionamento_incident = sla_horario_acionamento / sla_horario_incident
            // console.log(sla_horario_acionamento, sla_horario_incident, horario_acionamento_incident)
            // console.table([
            //     { "sla_horario_acionamento": sla_horario_acionamento, "sla_horario_incident": sla_horario_incident, "horario_acionamento_incident": horario_acionamento_incident },
            // ])
            var res = "";
            // if (horario_acionamento_incident >= 0.90) {
            //     res = "10% do SLA"
            // } else if (horario_acionamento_incident > 0.20 && horario_acionamento_incident < 0.90) {
            //     res = "até 80% do SLA"
            // } else if (horario_acionamento_incident < 0.20) {
            //     res = "80% do SLA"
            // } else if (horario_acionamento_incident == "NaN") {
            //     res = "Nan"
            // } else {
            //     res = "Solicitado prioridade com SLA vencido";
            // }
            if (horario_acionamento_incident >= 0.90) {
                res = "10% do SLA"
            } else if (horario_acionamento_incident > 0.20) {
                res = "até 80% do SLA"
            } else if (horario_acionamento_incident <= 0.20) {
                res = "80% do SLA"
            } else if (horario_acionamento_incident.toString() == "NaN") {
                res = "NaN"
            }

            if (sla_ticket_vencido == "Solicitado Prioridade com SLA Vencido") {
                res = "Solicitado prioridade com SLA vencido"
            }
            // res = res + " / " + horario_acionamento_incident
            // console.log(i, res);
            worksheet.getCell(config.ANALISE_PRAZO_ACIONAMENTO + i).value = res
            // if (i == 3) {
            //     break;
            // }
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
