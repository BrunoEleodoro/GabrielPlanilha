require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

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

const STORE_WORKED_HOURS = process.env.STORE_WORKED_HOURS

function secondsToTime(secs) {
    var pad = "000"
    var hours = Math.floor(secs / (60 * 60));

    var divisor_for_minutes = secs % (60 * 60);
    var minutes = Math.floor(divisor_for_minutes / 60);

    var divisor_for_seconds = divisor_for_minutes % 60;
    var seconds = Math.ceil(divisor_for_seconds);

    var obj = {
        "h": hours.toString().padStart(2, '0'),
        "m": minutes.toString().padStart(2, '0'),
        "s": seconds.toString().padStart(2, '0')
    };
    return obj;
}

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;

        //setting the title of the column
        worksheet.getCell(STORE_WORKED_HOURS + 1).value = "time worked"

        while (i <= worksheet.rowCount) {
            var closed_at_date = worksheet.getCell(STORE_CLOSED_AT + i).value
            var created_at_date = worksheet.getCell(CREATED_AT + i).value
            if (closed_at_date != null && created_at_date != null) {

                var startDate = new Date(created_at_date)
                var endDate = new Date(closed_at_date)

                var seconds = (endDate.getTime() - startDate.getTime()) / 1000;

                var t = new Date(1970, 0, 1);
                t.setSeconds(seconds);

                var finalTime = secondsToTime(seconds);
                console.log(finalTime.h);
                if (parseFloat(finalTime.h) > 8) {
                    finalTime = "08" + ":" + "00" + ":" + "00"
                } else {
                    finalTime = finalTime.h + ":" + finalTime.m + ":" + finalTime.s
                }
                worksheet.getCell(STORE_WORKED_HOURS + i).value = new Date()
                worksheet.getCell(STORE_WORKED_HOURS + i).value = finalTime
                worksheet.getCell(STORE_WORKED_HOURS + i).numFmt = 'hh:mm:ss';

            }
            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })

// console.log(new Date('01/01/2019 10:11').getHours())
// console.log(new Date(2019, 01, 01, 10, 11, 00, 00).getMinutes())