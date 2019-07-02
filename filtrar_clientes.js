require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];


var fs = require('fs'),
    path = require('path'),
    filePath = path.join(__dirname, 'blacklist.json');

fs.readFile(filePath, { encoding: 'utf-8' }, function (err, data) {
    if (!err) {
        not_allowed = JSON.parse(data);
    } else {
        console.log(err);
    }
});

function toUTF8(body) {
    // convert from iso-8859-1 to utf-8
    var ic = new iconv.Iconv('iso-8859-1', 'utf-8');
    var buf = ic.convert(body);
    return buf.toString('utf-8');
}
// CONTROLLERS
const PRIMARY_LABELS_COLUMN = process.env.PRIMARY_LABELS_COLUMN
const DESCRIPTION_COLUMN = process.env.DESCRIPTION_COLUMN
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN
const TITLE_COLUMN = process.env.TITLE_COLUMN
const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET
const STORE_SEVERITY_COLUNM = process.env.STORE_SEVERITY_COLUNM
const STORE_PRIMARY_LABELS_COLUMN = process.env.STORE_PRIMARY_LABELS_COLUMN

//Summary for severities
const SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS
const SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES

// Summary for severities of clients
const SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME
const SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1
const SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2
const SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3
const SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4

const clientes_possiveis = [
    "CARREFOUR",
    "FLEURY",
    "GERDAU",
    "BRF",
    "DPSP",
    "COPERSUCAR",
    "TIGRE",
    "DROGARIA ONOFRE",
    "ORBITALL",
    "LASA",
    "LEROY MERLIN",
    "RECORD",
    "SAINT-GOBAIN",
    "INTERMEDICA",
    "BCG",
    "GPA",
    "ADP",
    "VIA VAREJO",
    "LIVELO",
    "HONDA",
    "GALGO",
    "CIELO",
    "BRMALLS",
    "ETERNIT",
    "ARTERIS",
    "CONSTRUDECOR",
    "LEAO",
    "CMOC",
    "SENIOR",
    "AREZZO",
    "G2",
    "GENERALI",
    "WPP",
    "CEBRACE",
    "MULTIPLUS",
    "CONFESOL",
    "MANGELS",
    "ZETRASOFT",
    "FAST SHOP",
    "LEROY",
    "FIRST DATA",
    "BOA VISTA",
    "SPRINGER",
    "RIOCARD",
    "ESSILOR",
    "PROFARMA",
    "STELO",
    "FASTSHOP",
    "ALPARGATAS",
    "BANCO PINE",
    "SANTA HELENA",
    "REDECARD",
    "PESA",
    "MULTI VAREJO",
    "ORBITAL",
    "ASSAI",
    "CRDC",
    "MERCEDES BENZ"
]

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        var total = 0;
        //setting the title of the column
        worksheet.getCell(STORE_CLIENT_COLUMN + 1).value = "client"
        while (i <= worksheet.rowCount) {
            var valor_celula_p = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var k = 0;
            if (valor_celula_p != null) {
                var possivel_cliente = "";
                while (k < clientes_possiveis.length) {
                    if (valor_celula_p.toLowerCase().includes(clientes_possiveis[k].toLowerCase())) {
                        possivel_cliente = clientes_possiveis[k].toUpperCase()
                        break
                    }
                    k++;
                }

                if (possivel_cliente == "") {
                    var valor_client = worksheet.getCell(TITLE_COLUMN + i).value
                    valor_client = valor_client.split("-")[0]
                    // removing the brackets
                    valor_client = valor_client.replace("[", "").replace("]", "")
                    // removing the blank space from the beggining and from the end
                    valor_client = valor_client.trim()
                    possivel_cliente = valor_client
                    if (possivel_cliente.trim().toUpperCase() == "MBB" || possivel_cliente.trim().toUpperCase() == "MER") {
                        possivel_cliente = "MERCEDES BENZ"
                    }
                    if (possivel_cliente.trim().toUpperCase() == "RRC") {
                        possivel_cliente = "RECORD"
                    }
                    if (possivel_cliente.trim().toUpperCase() == "CAR") {
                        possivel_cliente = "CARREFOUR"
                    }
                    
                    worksheet.getCell(STORE_CLIENT_COLUMN + i).value = possivel_cliente.toUpperCase()
                } else {
                    worksheet.getCell(STORE_CLIENT_COLUMN + i).value = possivel_cliente.toUpperCase()
                }
            } else {
                var valor_client = worksheet.getCell(TITLE_COLUMN + i).value
                valor_client = valor_client.split("-")[0]
                // removing the brackets
                valor_client = valor_client.replace("[", "").replace("]", "")
                // removing the blank space from the beggining and from the end
                valor_client = valor_client.trim()
                possivel_cliente = valor_client
                if (possivel_cliente.trim().toUpperCase() == "MBB" || possivel_cliente.trim().toUpperCase() == "MER") {
                    possivel_cliente = "MERCEDES BENZ"
                }
                if (possivel_cliente.trim().toUpperCase() == "RRC") {
                    possivel_cliente = "RECORD"
                }
                if (possivel_cliente.trim().toUpperCase() == "CAR") {
                    possivel_cliente = "CARREFOUR"
                }
                worksheet.getCell(STORE_CLIENT_COLUMN + i).value = possivel_cliente.toUpperCase()
            }

            // if (true) {
            //     if (worksheet.getCell("D" + i).value == worksheet.getCell(STORE_CLIENT_COLUMN + i).value) {
            //         worksheet.getCell("M" + i).fill = {
            //             type: 'pattern',
            //             pattern: 'solid',
            //             fgColor: { argb: 'FFFFFF00' },
            //             bgColor: { argb: 'aa00ff' }
            //         };
            //     }

            // }
            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
