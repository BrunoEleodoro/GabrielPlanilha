const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

var moment = require('moment')

var dateFormat1 = 'DD/MM/YYYY HH:mm';
var dateFormat2 = 'MM/DD/YYYY HH:mm';

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
    if (data.toString() == "Invalid date") {
        data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    }

    return data
}

function decimalToHours(str) {

    var partes = str.toString().split('.')
    var horas = partes[0];
    var minutos = partes[1];
    if (minutos == undefined) {
        return "00:00";
    }
    if (parseFloat(partes[1]) > 60) {
        horas = parseFloat(horas) + 1
        minutos = (partes[1] - 60)
    } else if (parseFloat(partes[1]) == 6) {
        horas = parseFloat(horas) + 1
        minutos = 0
    }
    horas = horas.toString().padStart(2, '0')
    minutos = minutos.toString().padStart(2, '0')

    return horas + ":" + minutos
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.TEMPO_ATENDIMENTO + 1).value = "Tempo de Atendimento IBM"
        var tempo_atendimento_title = {};
        var i = 2
        while (i <= worksheet.rowCount) {
            var title = worksheet.getCell(config.STORE_TITLE_COLUMN + i).value
            var card_identifier = worksheet.getCell("AN" + i).value;
            var closed_at = worksheet.getCell(config.STORE_CLOSED_AT + i).value

            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var sla_do_ticket = worksheet.getCell(config.SLA_TICKET + i).value
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var quantidade_tickets_per_user = worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + i).value

            var key = horario_incidente + "" + sla_do_ticket + "" + horario_acionamento;

            var horario_encerramento = worksheet.getCell(config.HORARIO_ENCERRAMENTO + i).value
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value

            var horario_encerramento_date = parseDateToMoment(monthName, horario_encerramento);
            var horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
            var closed_at_date = parseDateToMoment(monthName, closed_at);

            var keys = ["#0aln3i", "#09ty4g", "#09vzqq", "#09u3t1", "#09ueus", "#09uj0e", "#09ukfj", "#09ulap", "#09uoi1", "#09urs5", "#09uu9y", "#09uy2g", "#09uy73", "#09vbnd", "#09vbr0", "#09vcu1", "#09vv8o", "#09vvvh"]
            if (keys.includes(card_identifier)) {
                console.log(card_identifier, horario_encerramento_date, horario_acionamento_date, closed_at_date);
            }

            var hours = calculateHours(
                horario_encerramento_date,
                horario_acionamento_date);

            if (hours < 0) {
                hours = hours * -1
            }
            hours = parseFloat(hours);
            //tempo_atendimento_title[key] = tempo_atendimento_title[key] + hours;
            //worksheet.getCell(config.TEMPO_ATENDIMENTO + i).numFmt = 'hh:mm';
            if (quantidade_tickets_per_user == 1) {
                worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = hours.toFixed(2)
            }
            //worksheet.getCell(config.TEMPO_ATENDIMENTO + i).numFmt = 'h:mm:ss';
            //worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = { formula: "MOD(MROUND(\"" + decimalToHours(hours.toFixed(2)) + "\",\"0:30\"),1)" }            

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
