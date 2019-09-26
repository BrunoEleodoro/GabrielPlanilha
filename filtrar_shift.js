require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

// CONTROLLERS
const CREATED_BY_COLUMN = process.env.CREATED_BY_COLUMN
const TITLE_COLUMN = process.env.TITLE_COLUMN
const PRIMARY_LABELS_COLUMN = process.env.PRIMARY_LABELS_COLUMN
const CLOSED_AT = process.env.CLOSED_AT
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN

const STORE_CREATED_BY_COLUMN = process.env.CREATED_BY_COLUMN
const STORE_TITLE_COLUMN = process.env.TITLE_COLUMN
const STORE_PRIMARY_LABELS_COLUMN = process.env.PRIMARY_LABELS_COLUMN
const STORE_CLOSED_AT = process.env.CLOSED_AT
const STORE_CLIENTS_COLUMN = process.env.CLIENTS_COLUMN

const STORE_SHIFT = process.env.STORE_SHIFT
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM
const STORE_WEEK_DAY = process.env.STORE_WEEK_DAY
const STORE_MONTH = process.env.STORE_MONTH

const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET


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

        //setting the title of the column
        worksheet.getCell(STORE_SHIFT + 1).value = "shift"

        while (i <= worksheet.rowCount) {
            var valor_celula_p = worksheet.getCell(CREATED_BY_COLUMN + i).value
            if (valor_celula_p != null) {
                var shift = null

                if (valor_celula_p.includes("Gabriela Ferreira Dias Dos Santos")) {
                    shift = parseFloat("1")
                } else if (valor_celula_p.includes("Lucas Gaspar Hoffelder")) {
                    shift = parseFloat("1")
                } else if (valor_celula_p.includes("Otavio De Almeida Sambo")) {
                    shift = parseFloat("1")
                } else if (valor_celula_p.includes("Jacqueline Cristina Da Silva")) {
                    shift = parseFloat("2")
                } else if (valor_celula_p.includes("Gabriel Siqueira")) {
                    shift = parseFloat("2")
                } else if (valor_celula_p.includes("Renan Diego Mafeis")) {
                    shift = parseFloat("2")
                } else if (valor_celula_p.includes("Diego Dayvison Alves De Araujo Ferreira")) {
                    shift = parseFloat("3")
                } else if (valor_celula_p.includes("Matheus Reis Villela")) {
                    shift = parseFloat("3")
                } else if (valor_celula_p.includes("Lalisa Viola Faria Santos")) {
                    shift = parseFloat("3")
                } else if (valor_celula_p.includes("Henrique Possari")) {
                    shift = parseFloat("2")
                }

                worksheet.getCell(STORE_SHIFT + i).value = shift
            }
            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
