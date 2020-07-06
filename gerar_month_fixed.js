require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];
var path = require('path')
const config = require('./load_columns');
const moment = require('moment');
// CONTROLLERS
const LABELS_COLUMN = process.env.LABELS_COLUMN
const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET
const CREATED_BY_COLUMN = process.env.CREATED_BY_COLUMN
const TITLE_COLUMN = process.env.TITLE_COLUMN
const PRIMARY_LABELS_COLUMN = process.env.PRIMARY_LABELS_COLUMN
const CLOSED_AT = process.env.CLOSED_AT
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN
const CREATED_AT = process.env.CREATED_AT

const STORE_CREATED_BY_COLUMN = process.env.STORE_CREATED_BY_COLUMN
const STORE_SHIFT = process.env.STORE_SHIFT
const STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN
const STORE_TITLE_COLUMN = process.env.STORE_TITLE_COLUMN
const STORE_PRIMARY_LABELS_COLUMN = process.env.STORE_PRIMARY_LABELS_COLUMN
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM
const STORE_CLOSED_AT = process.env.STORE_CLOSED_AT
const STORE_WEEK_DAY = process.env.STORE_WEEK_DAY
const STORE_MONTH = process.env.STORE_MONTH

const SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS
const SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES
const SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME
const SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1
const SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2
const SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3
const SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4
const SOURCE_COLUMNS_LIST = process.env.SOURCE_COLUMNS_LIST
const DESTINATION_COLUMNS_LIST = process.env.DESTINATION_COLUMNS_LIST


var months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

function changeDayAndMonthPosition(date, separator) {
    var pieces = date.split(separator)
    var newDate = pieces[1] + separator + pieces[0] + separator + pieces[2]
    return newDate
}



function parseDateToMoment(monthName, valor_celula) {
    var month = months.indexOf(monthName) + 1
    if (valor_celula == "" || valor_celula == null) {
        return moment("00/00/0000 00:00", "DD/MM/YYYY HH:mm");
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

    return data
}



// READ WORKBOOK
workbook.xlsx.readFile(path.join(__dirname, SOURCE_FILE))
    .then(function () {

        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        const d = new Date();

        //setting the title of the column
        worksheet.getCell(STORE_MONTH + 1).value = "month"

        while (i <= worksheet.rowCount) {

            var valor_celula = worksheet.getCell(CREATED_AT + i).value
            if (valor_celula != null) {
                var currentMonth = months[d.getMonth()]
                var data = moment(valor_celula, "DD/MM/YYYY HH:mm");
                if (data.toString() == "Invalid date") {
                    data = moment(valor_celula, "MM/DD/YYYY HH:mm");
                }

                var card_identifier = worksheet.getCell("AN" + i).value
                if (card_identifier == "07m4ea") {
                    var data2 = parseDateToMoment(currentMonth, valor_celula)

                    console.log('monthName', valor_celula, months[data2.month()])
                }

                worksheet.getCell(STORE_MONTH + i).value = months[data.month()]
            }
            // if (i % 5 == 0) {
            //     break;
            // }
            i++;
        }


        console.log('finalizado!');
        return workbook.xlsx.writeFile(path.join(__dirname, OUTPUT_FILE));
    });