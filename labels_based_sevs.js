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
const CARD_ASSIGNEES = process.env.CARD_ASSIGNEES

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

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;

        //setting the title of the column
        // worksheet.getCell(STORE_MONTH + 1).value = "month"

        while (i <= worksheet.rowCount) {
            var valor_celula = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
            var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
            var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
            if (valor_celula != null && assignee != null) {
                var pieces = valor_celula.split(",");
                var k = 0;
                var res = "";
                while (k < pieces.length) {
                    res += pieces[k].trim() + "_sev" + sev + "_" + type + ", "
                    k++;
                }
                res = res.substr(0, res.length - 2);

                var pieces_assignee = []
                if (!assignee.includes(",")) {
                    pieces_assignee.push(assignee)
                } else {
                    pieces_assignee = assignee.split(",");
                }
                
                var res2 = ""
                var pieces = res.split(",");
                var j = 0;
                while (j < pieces_assignee.length) {
                    var email = pieces_assignee[j].split("@")[0]
                    var k = 0;
                    while (k < pieces.length) {
                        res2 += pieces[k].trim() + "_" + email + ", "
                        k++;
                    }
                    j++;
                }

                res2 = res2.substr(0, res2.length - 2);

                worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = res+","+res2
            }

            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
