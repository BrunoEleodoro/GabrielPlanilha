// var fs = require('fs');

// var contents = fs.readFileSync("a.csv", { encoding: 'utf8' });

// console.log(contents.toString().split("\n")[32].toString())
var moment = require('moment');

var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function parseDateToMoment(monthName, valor_celula) {
    var month = months.indexOf(monthName) + 1
    if (valor_celula == "" || valor_celula == null) {
        return moment("00/00/0000 00:00", "MM/DD/YYYY HH:mm");
    }
    // var index = valor_celula.indexOf("0" + month + "/")
    var index = valor_celula.toString().split(" ")[0].indexOf(month)
    var data = ""
    console.log(index)

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

    return data
}

var data = parseDateToMoment("June", "16/06/2020 00:19");
console.log(data)