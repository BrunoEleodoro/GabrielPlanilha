require('dotenv').config({ path: 'config_criar_planilha' })
const utf8 = require('utf8');

var Excel = require('exceljs');
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

let criar_planilha = {}
criar_planilha["created by"] = "A"
criar_planilha["title"] = "C"
criar_planilha["primary labels"] = "E"
criar_planilha["created at"] = "H"
criar_planilha["card assignees"] = "K"
criar_planilha["closed at"] = "S"
criar_planilha["total waiting time (minutes)"] = "T"
criar_planilha["operational lead time (minutes)"] = "U"
criar_planilha["custom_field_1"] = "AD"
criar_planilha["custom_field_2"] = "AE"
criar_planilha["custom_field_3"] = "AF"
criar_planilha["custom_field_4"] = "AG"
criar_planilha["custom_field_5"] = "AH"
criar_planilha["card identifier"] = "AN"


async function main() {
    var XLSX = require('xlsx');
    var workbook = XLSX.readFile(path.join(__dirname, SOURCE_FILE));
    var sheet_name_list = workbook.SheetNames;
    var fs = require('fs');

    const csv = require('csvtojson')
    csv()
        .fromFile(path.join(__dirname, SOURCE_FILE))
        .then((jsonObj) => {
            fs.writeFileSync('planilha.json', JSON.stringify(jsonObj))
        })

    var data = JSON.parse(fs.readFileSync('planilha.json'))
    console.log(data.length)
    var workbook = new Excel.Workbook();
    workbook.xlsx.writeFile(OUTPUT_FILE)
        .then(function () {
            var worksheet = workbook.addWorksheet("Dados");
            var i = 0;
            while (i < data.length) {
                if (i == 0) {
                    let k = 0;
                    let keys = Object.keys(criar_planilha);
                    while (k < keys.length) {
                        worksheet.getRow(i + 1).getCell(criar_planilha[keys[k]]).value = keys[k]
                        k++;
                    }
                } else {
                    let k = 0;
                    let keys = Object.keys(criar_planilha);
                    while (k < keys.length) {
                        worksheet.getRow(i + 1).getCell(criar_planilha[keys[k]]).value = data[i][keys[k]]
                        k++;
                    }
                }
                // worksheet.getRow(i).getCell(DESTINATION_COLUMNS_LIST[k]).value = value

                i++;
            }

            return workbook.xlsx.writeFile(OUTPUT_FILE);
        });

}

main();