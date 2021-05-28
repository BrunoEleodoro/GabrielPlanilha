require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');
const config = require('./load_columns');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();


workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        var i = 2;

        //setting the title of the column
        worksheet.getCell(config.TIME_WORKED_ISM + 1).value = "time worked ism"

        var time_workeds = {}
        var time_workeds_amount = {}
        while (i <= worksheet.rowCount) {
            var time_worked = worksheet.getCell(config.STORE_WORKED_HOURS + i).value
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var quantidade_tickets_per_user = worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + i).value;

            var key = horario_incidente; 

            if (time_workeds[key] == null) {
                time_workeds[key] = 0;
                time_workeds_amount[key] = 0;
            }
            time_workeds[key] = time_workeds[key] + parseFloat(time_worked.toString())
            //time_workeds[key].push(time_worked)
            if(quantidade_tickets_per_user == 1) {
                time_workeds_amount[key]++;
            }
            i++;
        }

        i = 2;
        while (i <= worksheet.rowCount) {
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value
            var quantidade_tickets_per_user = worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + i).value;
            if(quantidade_tickets_per_user == 1) {
                worksheet.getCell(config.TIME_WORKED_ISM + i).value = parseFloat(time_workeds[horario_incidente] / time_workeds_amount[horario_incidente]).toFixed(2)
            }
            i++;
        }

        console.log('finalizado!');

        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })

// console.log(new Date('01/01/2019 10:11').getHours())
// console.log(new Date(2019, 01, 01, 10, 11, 00, 00).getMinutes())
