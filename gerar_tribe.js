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
var pessoas_clientes = []
var pessoas = ["Marcia Y Couto",
    "Edmar Lauria Marques",
    "Geraldo Vicente Ferreira",
    "Fabiano Gonçalves",
    "Paulinho Rossetti",
    "Multiplos Clientes",
    "Lívio Teixeira",
    "IBM Infra",
]
pessoas_clientes["Marcia Y Couto"] = [
    "ADP",
    "AREZZO",
    "ARTERIS",
    "BANCO PINE",
    "BCG",
    "BURGER KING",
    "CEBRACE",
    "CONSTRUDECOR",
    "CRESOL",
    "FLEURY",
    "GRUPO SIMOES",
    "LASA",
    "MULTIPLUS",
    "PROXXI",
    "RIOGALEAO",
    "SAINT GOBAIN",
    "SENIOR",
    "TICKET",
    "TOYOTA",
    "ZETRASOFT"
]
pessoas_clientes["Lívio Teixeira"] = [
    "CAIXA ECONOMICA"
]
pessoas_clientes["Multiplos Clientes"] = [
    "MULTIPLOS CLIENTES"
]
pessoas_clientes["IBM Infra"] = [
    "IRM"
]
pessoas_clientes["Edmar Lauria Marques"] = [
    "FIDELITY NATIONAL",
    "BANRISUL",
    "BOA VISTA",
    "BRF",
    "FIAT",
    "GERDAU",
    "HONDA",
    "MERCEDES",
    "PESA",
    "WPP",
    "RIOCARD",
    "UNILEVER",
    "GERDAU SIDERPERU",
    "GERDAU GLN",
    "GERDAU REMOTAS",
    "GERDAU MEXICO",
    "GERDAU PERU",
]
pessoas_clientes["Fabiano Gonçalves"] = [
    "ALPARGATAS",
    "APOLLO",
    "BRMALLS",
    "CARREFOUR",
    "CMOC",
    "COPERSUCAR",
    "CRDC",
    "DROGARIA SP",
    "ESSILOR",
    "ETERNIT",
    "ANBIMA",
    "GENERALI",
    "LEAO",
    "LEROY MERLIN",
    "MANGELS",
    "GPA",
    "RECORD",
    "REDECARD",
    "SANTA HELENA",
    "SPRINGER",
    "TIGRE",
    "VIA VAREJO",
    "GALGO"
]
pessoas_clientes["Geraldo Vicente Ferreira"] = [
    "SANTANDER",
]
pessoas_clientes["Paulinho Rossetti"] = [
    "BRADESCO",
    "BRADESCO MSA",
    "CIELO",
    "ELO",
    "LIVELO",
    "STELO"
]
// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        worksheet.getCell(TRIBE + 1).value = "Tribo"
        while (i <= worksheet.rowCount) {
            var k = 0;
            while (k < pessoas.length) {
                var value = worksheet.getCell(CLIENTS_COLUMN + i).value
                value = value.trim()
                if (i == 6) {
                    console.log(pessoa, value)
                    console.log(pessoas_clientes[pessoa].indexOf(value))
                }
                var pessoa = pessoas[k]
                if (pessoas_clientes[pessoa].indexOf(value) >= 0) {
                    worksheet.getCell(TRIBE + i).value = pessoa
                    break;
                }
                k++;
            }
            //     worksheet.getCell(TOTAL_WAITING_TIME + i).value = parseFloat(worksheet.getCell(TOTAL_WAITING_TIME + i).value)
            //     worksheet.getCell(OPERATIONAL_LEAD_TIME + i).value = parseFloat(worksheet.getCell(OPERATIONAL_LEAD_TIME + i).value)
            //     worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value = parseFloat(worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value)
            //     worksheet.getCell(STORE_DAY + i).value = parseFloat(worksheet.getCell(STORE_DAY + i).value)
            //     worksheet.getCell(STORE_YEAR + i).value = parseFloat(worksheet.getCell(STORE_YEAR + i).value)
            //     worksheet.getCell(QUANTIDADE_TICKETS_PER_USER + i).value = parseFloat(worksheet.getCell(QUANTIDADE_TICKETS_PER_USER + i).value)
            i++;
        }
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })