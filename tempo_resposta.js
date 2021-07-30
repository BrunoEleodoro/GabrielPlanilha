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

function parseDateToMoment(monthName, valor_celula) {
    let data = "";
    valor_celula = valor_celula.replace("/21 ", "/2021 ")
    var month = months.indexOf(monthName) + 1
    if (valor_celula == "" || valor_celula == null) {
        return moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    }
    // var index = valor_celula.indexOf("0" + month + "/")
    var index = valor_celula.split(" ")[0].indexOf(month)

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


function parseTimeToMoment(monthName, valor_celula) {
    let data = "";
    valor_celula = valor_celula.replace("/21 ", "/2021 ")
    if (valor_celula.split(" ").length > 0) {
        data = moment(valor_celula.split(" ")[1], "HH:mm")
    }
    // var month = months.indexOf(monthName) + 1
    // if (valor_celula == "" || valor_celula == null) {
    //     return moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    // }
    // // var index = valor_celula.indexOf("0" + month + "/")
    // var index = valor_celula.split(" ")[0].indexOf(month)
    // var data = ""

    // if (index == -1) {
    //     data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    // } else if (index == 0) {
    //     data = moment(valor_celula, "MM/DD/YYYY HH:mm");
    // } else if (index == 3) {
    //     data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    // } else if (index == 4) {
    //     data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    // } else {
    //     data = moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    // }
    // if (data.toString() == "Invalid date") {
    //     data = moment(valor_celula, "MM/DD/YYYY HH:mm");
    // }
    // if (data.toString() == "Invalid date") {
    //     data = moment(valor_celula, "DD/MM/YYYY HH:mm");
    // }

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

function convertToHHMM(value) {
    var decimalTimeString = value;
    var decimalTime = parseFloat(decimalTimeString);
    decimalTime = decimalTime * 60 * 60;
    var hours = Math.floor((decimalTime / (60 * 60)));
    decimalTime = decimalTime - (hours * 60 * 60);
    var minutes = Math.floor((decimalTime / 60));
    decimalTime = decimalTime - (minutes * 60);
    var seconds = Math.round(decimalTime);
    if (hours < 10) {
        hours = "0" + hours;
    }
    if (minutes < 10) {
        minutes = "0" + minutes;
    }
    if (seconds < 10) {
        seconds = "0" + seconds;
    }
    return hours + ":" + minutes;
}

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.TEMPO_RESPOSTA + 1).value = "Tempo de Resposta ISM"
        var i = 2
        var horario_incidente_map = {}
        while (i <= worksheet.rowCount) {
            var title = worksheet.getCell(config.TITLE_COLUMN + i).value
            var created_at = worksheet.getCell(config.CREATED_AT + i).value
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var sla_do_ticket = worksheet.getCell(config.SLA_TICKET + i).value
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value;
            var card_identifier = worksheet.getCell("AN" + i).value;

            if (horario_incidente_map[horario_incidente] == null) {
                horario_incidente_map[horario_incidente] = []
            }
            horario_incidente_map[horario_incidente].push(i)

            var key = horario_incidente + "" + sla_do_ticket + "" + horario_acionamento;
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value
            var type = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value

            var horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
            var created_at = parseDateToMoment(monthName, created_at);

            var newHorarioAcionamentoDate = created_at.format('DD/MM/YYYY').toString();
            var horarioAcionamentoFinal = newHorarioAcionamentoDate + " " + horario_acionamento.split(" ")[1]

            //var horarioAcionamentoFinal = newHorarioAcionamentoDate + " " + (horario_acionamento.split(" ")[1] == undefined ? horario_acionamento.split(" ")[0] : horario_acionamento.split(" ")[1])
            worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value = horarioAcionamentoFinal

            var keys = ["#0al7mx", "#0aliij", "#0aym5v"]
            if (keys.includes(card_identifier)) {
                console.log({
                    'cod': '1',
                    'card_identifier': card_identifier,
                    'horario_acionamento': worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value,
                    'created_at': worksheet.getCell(config.CREATED_AT + i).value
                })
            }
            if (type == "CH" || type == "REPORT" || type == "SC") {
                worksheet.getCell(config.TEMPO_RESPOSTA + i).value = null;
            }
            i++;
        }
        i = 2
        while (i <= worksheet.rowCount) {
            var title = worksheet.getCell(config.TITLE_COLUMN + i).value
            var created_at = worksheet.getCell(config.CREATED_AT + i).value
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var horario_acionamento = worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value
            var card_identifier = worksheet.getCell("AN" + i).value;
            var monthName = worksheet.getCell(config.STORE_MONTH + i).value
            var type = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value

            var horario_acionamento_date = parseTimeToMoment(monthName, horario_acionamento);
            var created_at_date = parseTimeToMoment(monthName, created_at);

            var hours = calculateHours(
                horario_acionamento_date,
                created_at_date);

            hours = hours < 0 ? hours * -1 : hours
            if (hours >= 4) {
                horario_acionamento_date = parseDateToMoment(monthName, horario_acionamento);
                created_at_date = parseDateToMoment(monthName, created_at);
                hours = calculateHours(
                    horario_acionamento_date,
                    created_at_date);
                console.log('hours', hours)
            }


            var keys = ["#0al7mx", "#0aliij", "#0aym5v"]
            if (keys.includes(card_identifier)) {
                console.log({
                    'cod': '2',
                    'hours': hours,
                    'card_identifier': card_identifier,
                    'horario_acionamento': worksheet.getCell(config.HORARIO_ACIONAMENTO + i).value,
                    'created_at': worksheet.getCell(config.CREATED_AT + i).value
                })
            }

            // var keys = ["#0afl0u", "#0apm0z", "#0apx6s", "#0aqcf4"]
            // if (keys.includes(card_identifier)) {
            //     console.log(card_identifier, horario_acionamento_date, created_at, hours)
            // }

            // worksheet.getCell(config.TEMPO_ATENDIMENTO + i).numFmt = 'hh:mm';
            worksheet.getCell(config.TEMPO_RESPOSTA + i).value = hours.toFixed(2)


            if (type == "CH" || type == "REPORT" || type == "SC") {
                worksheet.getCell(config.TEMPO_RESPOSTA + i).value = null;
            }
            i++;
        }

        var i = 0;
        var incidentes = Object.keys(horario_incidente_map);
        var limite = incidentes.length;
        while (i < limite) {
            var horario_incidente_indexes = horario_incidente_map[incidentes[i]]
            var k = 0;
            var lowest_created_at_index = -1;
            var lowest_created_at = null;
            while (k < horario_incidente_indexes.length) {
                var index = horario_incidente_indexes[k]
                var created_at = worksheet.getCell(config.CREATED_AT + index).value
                var card_identifier = worksheet.getCell("AN" + index).value
                var tempo_resposta = worksheet.getCell(config.TEMPO_RESPOSTA + index).value
                var data = moment(created_at, "DD/MM/YYYY HH:mm");
                if (data.toString() == "Invalid date") {
                    data = moment(created_at, "DD/MM/YYYY HH:mm");
                }
                if (data.toString() == "Invalid date") {
                    data = moment(created_at, "MM/DD/YYYY HH:mm");
                }

                if (lowest_created_at == null) {
                    lowest_created_at = data;
                    lowest_created_at_index = index;
                } else if (lowest_created_at.isAfter(data)) {
                    lowest_created_at = data;
                    lowest_created_at_index = index;
                }
                var keys = ["#0ar1so", "#0ar3bz", "#0ar5p8"]
                if (keys.includes(card_identifier)) {
                    // console.log("tempo_resposta aqui", card_identifier, index, tempo_resposta)
                }
                k++;
            }
            //IN89517080
            k = 0;
            while (k < horario_incidente_indexes.length) {
                var index = horario_incidente_indexes[k]
                var card_identifier = worksheet.getCell("AN" + index).value
                var horario_encerramento = worksheet.getCell(config.HORARIO_ENCERRAMENTO + index).value
                if (horario_encerramento != null) {
                    var hours = worksheet.getCell(config.TEMPO_RESPOSTA + lowest_created_at_index).value
                    var keys = ["#0al7mx", "#0aliij", "#0aym5v"]
                    if (keys.includes(card_identifier)) {
                        console.log('tempo _resposta', card_identifier, hours)
                    }
                    if (hours >= 24) {
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).value = "24:00";

                    } else if (hours <= 0.04) {
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).value = "00:05";

                    } else if (hours.toString().includes("NaN")) {
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).value = "00:00";

                    } else if (hours.formula != null) {
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).numFmt = 'hh:mm';
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).value = hours

                    } else {
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).numFmt = 'hh:mm';
                        worksheet.getCell(config.TEMPO_RESPOSTA + index).value = { formula: "MOD(MROUND(\"" + convertToHHMM(hours) + "\",\"0:05\"),1)" }
                    }
                } else {
                    worksheet.getCell(config.TEMPO_RESPOSTA + index).value = null
                }

                k++;
            }

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
