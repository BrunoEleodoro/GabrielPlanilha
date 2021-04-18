// var fs = require('fs');

// var contents = fs.readFileSync("a.csv", { encoding: 'utf8' });

// console.log(contents.toString().split("\n")[32].toString())
var moment = require('moment');

function calculateHours(startDate, endDate) {

    var duration = moment.duration(endDate.diff(startDate));
    var hours = duration.asHours();
    // hours = moment(hours * 3600 * 1000).format('HH:mm')
    return hours;
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

// console.log(convertToHHMM(calculateHours(

//     moment("28/02/21 22:12:00", "DD/MM/YYYY HH:mm"),
//     moment("02/28/2021 22:14", "MM/DD/YYYY HH:mm")
// )));

// console.log(convertToHHMM(0.04))
console.log( 0.03333333333333333 <= 0.04)