require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

// CONTROLLERS
const STORE_PRIMARY_LABELS_COLUMN = process.env.STORE_PRIMARY_LABELS_COLUMN
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const DESCRIPTION_COLUMN = process.env.DESCRIPTION_COLUMN
const STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN
const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET
const STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM
const TRIBE = process.env.TRIBE
const HORARIO_PICO = process.env.HORARIO_PICO
const CREATED_AT = process.env.CREATED_AT

const TOTAL_WAITING_TIME = process.env.TOTAL_WAITING_TIME
const OPERATIONAL_LEAD_TIME = process.env.OPERATIONAL_LEAD_TIME
const STORE_QUANTIDADE_TICKETS = process.env.STORE_QUANTIDADE_TICKETS
const STORE_DAY = process.env.STORE_DAY
const STORE_YEAR = process.env.STORE_YEAR
const QUANTIDADE_TICKETS_PER_USER = process.env.QUANTIDADE_TICKETS_PER_USER

//Summary for severities
const SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS
const SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES

// Summary for severities of clients
const SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME
const SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1
const SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2
const SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3
const SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        worksheet.getCell(HORARIO_PICO + 1).value = "Horario de Pico"
        while (i <= worksheet.rowCount) {
            var created_at_value = worksheet.getCell(CREATED_AT + i).value
            if (created_at_value != null && created_at_value.includes("/")) {
                var partes = created_at_value.split(" ");
                var time = partes[1];
                if (time != undefined) {
                    // var pieces = parseFloat(time.split(":").join("."))
                    // console.log(i, pieces)
                    // if (pieces >= 23.50 || pieces < 0.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "00:00"
                    // } else if (pieces >= 0.10 && pieces < 0.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "00:00"
                    // } else if (pieces >= 0.30 && pieces < 0.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "00:20"
                    // } else if (pieces >= 0.50 && pieces < 1.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "00:40"
                    // } else if (pieces >= 1.10 && pieces < 1.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "01:00"
                    // } else if (pieces >= 1.30 && pieces < 1.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "01:20"
                    // } else if (pieces >= 1.50 && pieces < 2.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "01:40"
                    // } else if (pieces >= 2.10 && pieces < 2.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "02:00"
                    // } else if (pieces >= 2.30 && pieces < 2.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "02:20"
                    // } else if (pieces >= 2.50 && pieces < 3.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "02:40"
                    // } else if (pieces >= 3.10 && pieces < 3.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "03:00"
                    // } else if (pieces >= 3.30 && pieces < 3.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "03:20"
                    // } else if (pieces >= 3.50 && pieces < 4.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "03:40"
                    // } else if (pieces >= 4.10 && pieces < 4.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "04:00"
                    // } else if (pieces >= 4.30 && pieces < 4.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "04:20"
                    // } else if (pieces >= 4.50 && pieces < 5.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "04:40"
                    // } else if (pieces >= 5.10 && pieces < 5.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "05:00"
                    // } else if (pieces >= 5.30 && pieces < 5.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "05:20"
                    // } else if (pieces >= 5.50 && pieces < 6.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "05:40"
                    // } else if (pieces >= 6.10 && pieces < 6.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "06:00"
                    // } else if (pieces >= 6.30 && pieces < 6.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "06:20"
                    // } else if (pieces >= 6.50 && pieces < 7.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "06:40"
                    // } else if (pieces >= 7.10 && pieces < 7.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "07:00"
                    // } else if (pieces >= 7.30 && pieces < 7.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "07:20"
                    // } else if (pieces >= 7.50 && pieces < 8.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "07:40"
                    // } else if (pieces >= 8.10 && pieces < 8.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "08:00"
                    // } else if (pieces >= 8.30 && pieces < 8.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "08:20"
                    // } else if (pieces >= 8.50 && pieces < 9.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "08:40"
                    // } else if (pieces >= 9.10 && pieces < 9.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "09:00"
                    // } else if (pieces >= 9.30 && pieces < 9.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "09:20"
                    // } else if (pieces >= 9.50 && pieces < 10.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "09:40"
                    // } else if (pieces >= 10.10 && pieces < 10.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "10:00"
                    // } else if (pieces >= 10.30 && pieces < 10.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "10:20"
                    // } else if (pieces >= 10.50 && pieces < 11.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "10:40"
                    // } else if (pieces >= 11.10 && pieces < 11.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "11:00"
                    // } else if (pieces >= 11.30 && pieces < 11.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "11:20"
                    // } else if (pieces >= 11.50 && pieces < 12.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "11:40"
                    // } else if (pieces >= 12.10 && pieces < 12.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "12:00"
                    // } else if (pieces >= 12.30 && pieces < 12.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "12:20"
                    // } else if (pieces >= 12.50 && pieces < 13.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "12:40"
                    // } else if (pieces >= 13.10 && pieces < 13.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "13:00"
                    // } else if (pieces >= 13.30 && pieces < 13.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "13:20"
                    // } else if (pieces >= 13.50 && pieces < 14.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "13:40"
                    // } else if (pieces >= 14.10 && pieces < 14.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "14:00"
                    // } else if (pieces >= 14.30 && pieces < 14.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "14:20"
                    // } else if (pieces >= 14.50 && pieces < 15.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "14:40"
                    // } else if (pieces >= 15.10 && pieces < 15.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "15:00"
                    // } else if (pieces >= 15.30 && pieces < 15.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "15:20"
                    // } else if (pieces >= 15.50 && pieces < 16.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "15:40"
                    // } else if (pieces >= 16.10 && pieces < 16.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "16:00"
                    // } else if (pieces >= 16.30 && pieces < 16.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "16:20"
                    // } else if (pieces >= 16.50 && pieces < 17.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "16:40"
                    // } else if (pieces >= 17.10 && pieces < 17.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "17:00"
                    // } else if (pieces >= 17.30 && pieces < 17.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "17:20"
                    // } else if (pieces >= 17.50 && pieces < 18.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "17:40"
                    // } else if (pieces >= 18.10 && pieces < 18.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "18:00"
                    // } else if (pieces >= 18.30 && pieces < 18.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "18:20"
                    // } else if (pieces >= 18.50 && pieces < 19.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "18:40"
                    // } else if (pieces >= 19.10 && pieces < 19.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "19:00"
                    // } else if (pieces >= 19.30 && pieces < 19.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "19:20"
                    // } else if (pieces >= 19.50 && pieces < 20.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "19:40"
                    // } else if (pieces >= 20.10 && pieces < 20.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "20:00"
                    // } else if (pieces >= 20.30 && pieces < 20.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "20:20"
                    // } else if (pieces >= 20.50 && pieces < 21.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "20:40"
                    // } else if (pieces >= 21.10 && pieces < 21.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "21:00"
                    // } else if (pieces >= 21.30 && pieces < 21.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "21:20"
                    // } else if (pieces >= 21.50 && pieces < 22.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "21:40"
                    // } else if (pieces >= 22.10 && pieces < 22.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "22:00"
                    // } else if (pieces >= 22.30 && pieces < 22.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "22:20"
                    // } else if (pieces >= 22.50 && pieces < 23.10) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "22:40"
                    // } else if (pieces >= 23.10 && pieces < 23.30) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "23:00"
                    // } else if (pieces >= 23.30 && pieces < 23.50) {
                    //     worksheet.getCell(HORARIO_PICO + i).value = "23:20"
                    // }
                    worksheet.getCell(HORARIO_PICO + i).numFmt = 'h:mm:ss';
                    worksheet.getCell(HORARIO_PICO + i).value = { formula: "MOD(MROUND(\"" + time + "\",\"0:15\"),1)" }

                }

                //  > 00,00 e < 07,59 então adicionar "3" na coluna shift
                //  > 08,00 e < 15,59 então adicionar "1" na coluna shift
                //  > 16,00 e < 23,59 então adicionar "2" na coluna shift
            }

            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
