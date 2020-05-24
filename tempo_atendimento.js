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

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.TEMPO_ATENDIMENTO + 1).value = "Tempo de Atendimeto ISM"
        var i = 2
        while (i <= worksheet.rowCount) {
            var closed_at = worksheet.getCell(config.STORE_CLOSED_AT + i).value
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value

            var hours = calculateHours(
                horario_acionamento,
                "DD/MM/YY HH:mm",
                "MM/DD/YY HH:mm",

                closed_at,
                "MM/DD/YYYY HH:mm",
                "DD/MM/YYYY HH:mm");

            if (hours < 0) {
                hours = hours * -1
            }

            worksheet.getCell(config.TEMPO_ATENDIMENTO + i).value = hours

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
