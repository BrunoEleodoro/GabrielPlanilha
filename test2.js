// var fs = require('fs');

// var contents = fs.readFileSync("a.csv", { encoding: 'utf8' });

// console.log(contents.toString().split("\n")[32].toString())
var moment = require('moment');

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

var horario_acionamento = "4/29/20 18:14:00"
var sla_ticket = "4/29/20 20:13:31"
var horario_incident = "4/29/20 18:13:31"

var horario_acionamento_date = parseDateToMoment("May", horario_acionamento);
var sla_ticket_date = parseDateToMoment("May", sla_ticket);
var horario_incident_date = parseDateToMoment("May", horario_incident);

var sla_horario_acionamento = calculateHours(sla_ticket_date, horario_acionamento_date)
var sla_horario_incident = calculateHours(sla_ticket_date, horario_incident_date);

console.log(horario_acionamento_date)
console.log(sla_ticket_date)
console.log(horario_incident_date)

if (sla_horario_acionamento < 0) {
    sla_horario_acionamento = sla_horario_acionamento * -1
}
if (sla_horario_incident < 0) {
    sla_horario_incident = sla_horario_incident * -1
}
var horario_acionamento_incident = sla_horario_acionamento / sla_horario_incident
console.log('horario_acionamento_incident', horario_acionamento_incident)

// console.log(horario_acionamento_date.isAfter(sla_ticket_date, 'seconds'))


// console.log(contents.toString().split("\n")[941].toString().split(",").length)
// console.log(contents.toString().split("\n")[942].toString().split(",").length)
// console.log(contents.toString().split("\n")[943].toString().split(",").length)

