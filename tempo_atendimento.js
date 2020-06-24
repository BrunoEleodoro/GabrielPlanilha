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
        worksheet.getCell(config.TEMPO_ATENDIMENTO + 1).value = "Tempo de Atendimento ISM"
        var i = 2
        while (i <= worksheet.rowCount) {
            var closed_at = worksheet.getCell(config.STORE_CLOSED_AT + i).value
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value

            var horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
            var closed_at_date = parseDateToMoment(monthName, closed_at);

            var hours = calculateHours(
                horario_acionamento_date,
                closed_at_date);

            if (hours < 0) {
                hours = hours * -1
            }
            // worksheet.getCell(config.TEMPO_ATENDIMENTO + i).numFmt = 'hh:mm';
            worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = hours.toFixed(2)

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
