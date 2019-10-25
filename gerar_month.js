require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];
var path = require('path')

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


// READ WORKBOOK
workbook.xlsx.readFile(path.join(__dirname, SOURCE_FILE))
    .then(function () {

        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;

        //setting the title of the column
        worksheet.getCell(STORE_MONTH + 1).value = "month"
        var valor_anterior = ""
        var mes_anterior = ""
        var mes_certo = ""

        while (i <= worksheet.rowCount) {

            var valor_celula = worksheet.getCell(STORE_CLOSED_AT + i).value

            if (valor_celula != null) {
                var pieces = valor_celula.split(" ")
                var date = pieces[0].trim();
                if (parseFloat(date.split("/")[0]) > 12) {
                    worksheet.getCell(STORE_CLOSED_AT + i).value = changeDayAndMonthPosition(date, "/") + " " + pieces[1]
                }
                // worksheet.getCell(STORE_MONTH + i).value = dayName
            }

            i++;
        }

        var first_parts = []
        var first_values = []
        i = 2;
        while (i <= worksheet.rowCount) {
            var valor_celula = worksheet.getCell(STORE_CLOSED_AT + i).value
            if (valor_celula != null) {
                var pieces = valor_celula.split(" ")
                var first = pieces[0].trim().split('/')[0];
                var index = first_parts.indexOf(first)
                if (index == -1) {
                    first_parts.push(first);
                    first_values.push(1)
                } else {
                    first_values[index] = parseFloat(first_values[index]) + 1
                }
                // worksheet.getCell(STORE_MONTH + i).value = dayName
            }
            i++;
            // if (i > 30) { break }
        }
        // console.log(JSON.stringify(first_parts), JSON.stringify(first_values))

        var maior_valor = 0;
        var index = 0;
        i = 0;
        while (i <= first_parts.length) {
            var valor = first_values[i]
            if (maior_valor < valor) {
                index = i
                maior_valor = valor
            }
            i++;
            // if (i > 30) { break }
        }

        var primeiro_maior_valor = maior_valor;
        var primeiro_maior_valor_index = index;

        // first_parts.splice(index)
        // first_values.splice(index)

        var maior_valor = 0;
        var index = 0;
        i = 0;
        while (i <= first_parts.length) {
            var valor = first_values[i]
            if (maior_valor < valor && valor != primeiro_maior_valor) {
                index = i
                maior_valor = valor
            }
            i++;
        }

        var segundo_maior_valor = maior_valor;
        var segundo_maior_valor_index = index;

        var meses_permitidos = []
        meses_permitidos.push(first_parts[primeiro_maior_valor_index])
        meses_permitidos.push(first_parts[segundo_maior_valor_index])



        i = 2;
        while (i <= worksheet.rowCount) {
            var valor_celula = worksheet.getCell(STORE_CLOSED_AT + i).value
            if (valor_celula != null) {
                var pieces = valor_celula.split(" ")
                var date = pieces[0].trim();

                var index1 = meses_permitidos.indexOf(date.split("/")[0])
                var index2 = meses_permitidos.indexOf(date.split("/")[1])

                if (index1 == -1 && index2 != -1) {
                    worksheet.getCell(STORE_CLOSED_AT + i).value = changeDayAndMonthPosition(date, "/")
                }

                var date = new Date(worksheet.getCell(STORE_CLOSED_AT + i).value.split(" ")[0].trim())
                var dayName = months[date.getMonth()];
                worksheet.getCell(STORE_MONTH + i).value = dayName
            }
            i++;
            // if (i > 30) { break }
        }
        // while (i <= worksheet.rowCount) {
        //     var valor_celula = worksheet.getCell(STORE_CLOSED_AT + i).value
        //     if (valor_celula != null) {
        //         var pieces = valor_celula.split(" ")
        //         var date = new Date(pieces[0])
        //         var dayName = months[date.getMonth()];

        //         if (valor_anterior != "" && valor_anterior != dayName) {
        //             // console.log(mes_anterior, date.getMonth())
        //             if (mes_anterior != date.getMonth() && date.getMonth() != mes_certo) {
        //                 // dayName = valor_anterior
        //                 dayName = months[parseFloat(pieces[0].split("/")[1]) - 1]
        //                 var newDateUS = pieces[0].split("/")[1] + "/" + pieces[0].split("/")[0] + "/" + pieces[0].split("/")[2]
        //                 var date = new Date(newDateUS)
        //                 worksheet.getCell(STORE_CLOSED_AT + i).value = newDateUS + " " + pieces[1]
        //                 dayName = months[date.getMonth()];
        //             }

        //         }

        //         if (mes_anterior == "") {
        //             mes_anterior = date.getMonth() - 1
        //             if (mes_anterior == 0) {
        //                 mes_anterior = 11
        //             }
        //         }
        //         if(mes_certo == "") {
        //             mes_certo = date.getMonth()
        //         }
        //         // if (dayName == null) {

        //         // }
        //         worksheet.getCell(STORE_MONTH + i).value = dayName
        //         valor_anterior = dayName
        //     }

        //     i++;
        // }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(path.join(__dirname, OUTPUT_FILE));
    });