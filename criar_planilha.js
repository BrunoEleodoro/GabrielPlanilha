require('dotenv').config({ path: 'config_criar_planilha' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

var fs = require('fs')
var path = require('path')

// CONTROLLERS
const LABELS_COLUMN = process.env.LABELS_COLUMN
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const DESCRIPTION_COLUMN = process.env.DESCRIPTION_COLUMN
const STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN
const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET
const STORE_CREATED_BY_COLUMN = process.env.STORE_CREATED_BY_COLUMN
const STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM

//Summary for severities
const SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS
const SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES

// Summary for severities of clients
const SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME
const SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1
const SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2
const SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3
const SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4

// Create new spreadsheet
const SOURCE_COLUMNS_LIST = process.env.SOURCE_COLUMNS_LIST.split(",")
const DESTINATION_COLUMNS_LIST = process.env.DESTINATION_COLUMNS_LIST.split(",")


const listCreatedBy = []
const matrice = []

function convert(input) {
    if (input == null) {
        return input
    }
    var iconv = require('iconv-lite');
    var output = iconv.decode(input, "UTF-8");
    // output = iconv.decode(output, "UTF-8");
    return output;
}

// READ WORKBOOK
async function main() {
    var linhas_novas = []
    var titles = []
    var workbook = new Excel.Workbook();
    if (fs.existsSync(path.join(__dirname, SOURCE_FILE))) {
        await workbook.csv.readFile(path.join(__dirname, SOURCE_FILE))
            .then(function () {
                var worksheet = workbook.worksheets[0];
                var i = 2;
                console.log('line count', worksheet.rowCount)
                var limite = (worksheet.rowCount + 1);
                while (i <= limite) {
                    var linha_antiga = []

                    var k = 0;
                    while (k < SOURCE_COLUMNS_LIST.length) {
                        //getting the titles
                        if (i == 2) {
                            titles.push(worksheet.getRow(1).getCell(SOURCE_COLUMNS_LIST[k]).value)
                        }
                        var value = worksheet.getRow(i).getCell(SOURCE_COLUMNS_LIST[k]).value
                        if (value != null) {
                            linha_antiga.push(convert(value.toString()))
                        } else {
                            linha_antiga.push(value)
                        }

                        k++;
                    }

                    linhas_novas.push(linha_antiga);
                    i++;
                }
                return 'finalizado'
            })

        var workbook = new Excel.Workbook();
        workbook.xlsx.writeFile(OUTPUT_FILE)
            .then(function () {
                var worksheet = workbook.addWorksheet("Dados");
                var i = 2;
                while (i <= linhas_novas.length) {

                    var k = 0;
                    while (k < DESTINATION_COLUMNS_LIST.length) {
                        //writing the titles
                        if (i == 2) {
                            worksheet.getRow(1).getCell(DESTINATION_COLUMNS_LIST[k]).value = titles[k]
                        }

                        var value = linhas_novas[i - 2][k]
                        worksheet.getRow(i).getCell(DESTINATION_COLUMNS_LIST[k]).value = value
                        k++;
                    }
                    i++;
                }
                console.log('finalizado!');
                return workbook.xlsx.writeFile(OUTPUT_FILE);
            });

    } else {
        console.log('file not found', path.join(__dirname, SOURCE_FILE))
        console.log('try this one', SOURCE_FILE, fs.existsSync(SOURCE_FILE))
    }

}

main();
