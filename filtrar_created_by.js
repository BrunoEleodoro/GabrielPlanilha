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

const listCreatedBy = []
// READ WORKBOOK
async function main() {
    var workbook = new Excel.Workbook();
    await workbook.csv.readFile(SOURCE_FILE)
        .then(function () {
            var worksheet = workbook.worksheets[0];
            var i = 2;

            while (i <= worksheet.rowCount) {
                var valor_celula_p = worksheet.getCell(CREATED_BY_COLUMN + i).value

                if (valor_celula_p != null) {
                    var created_by = null

                    if (valor_celula_p.toLowerCase().includes("gabriela ferreira dias dos santos")) {
                        created_by = "gabriela ferreira dias dos santos"
                    } else if (valor_celula_p.toLowerCase().includes("lucas gaspar hoffelder")) {
                        created_by = "lucas gaspar hoffelder"
                    } else if (valor_celula_p.toLowerCase().includes("otavio de almeida sambo")) {
                        created_by = "otavio de almeida sambo"
                    } else if (valor_celula_p.toLowerCase().includes("catia harume yamamoto")) {
                        created_by = "catia harume yamamoto"
                    } else if (valor_celula_p.toLowerCase().includes("jacqueline cristina da silva")) {
                        created_by = "jacqueline cristina da silva"
                    } else if (valor_celula_p.toLowerCase().includes("gabriel siqueira")) {
                        created_by = "gabriel siqueira"
                    } else if (valor_celula_p.toLowerCase().includes("renan diego mafeis")) {
                        created_by = "renan diego mafeis"
                    } else if (valor_celula_p.toLowerCase().includes("diego dayvison alves de araujo ferreira")) {
                        created_by = "diego dayvison alves de araujo ferreira"
                    } else if (valor_celula_p.toLowerCase().includes("matheus reis villela")) {
                        created_by = "matheus reis villela"
                    } else if (valor_celula_p.toLowerCase().includes("lalisa viola faria santos")) {
                        created_by = "lalisa viola faria santos"
                    }
                    listCreatedBy.push(created_by);
                }
                i++;
            }
            return 'finalizad'
        })
    var workbook = new Excel.Workbook();
    workbook.xlsx.writeFile(OUTPUT_FILE)
        .then(function () {
            var worksheet = workbook.addWorksheet("Dados");
            var i = 0;
            while (i < listCreatedBy.length) {
                worksheet.getCell(STORE_CREATED_BY_COLUMN + i).value = listCreatedBy[i].toUpperCase()
                i++;
            }
            console.log('finalizado!');
            return workbook.xlsx.writeFile(OUTPUT_FILE);
        });

}

main();