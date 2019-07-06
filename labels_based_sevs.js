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

const AREAS_ENVOLVIDAS = process.env.AREAS_ENVOLVIDAS
const SERVICE_LINE = process.env.SERVICE_LINE
const PROBLEMA_REPORTADO = process.env.PROBLEMA_REPORTADO
const ANALISE_ACIONAMENTO = process.env.ANALISE_ACIONAMENTO
const LABELS_ALEATORIAS = process.env.LABELS_ALEATORIAS
const ACAO_ISM = process.env.ACAO_ISM

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

var areas_envolvidas = [
    "acionamento técnico ibm",
    "acionamento técnico",
    "acionamento tecnico",
    "acionamento t",
    "acionamento cliente",
    "acionamento sam",
    "sam",
    "acionamento sme",
    "sme",
    "acionamento dpe",
    "dpe",
    "acionamento dm",
    "dm"
]

var service_line = [
    "sql support",
    "sql issue",
    "sap support",
    "sap issue",
    "middleware support",
    "middleware issue",
    "iam support",
    "iam issue",
    "network support",
    "network issue",
    "firewall support",
    "firewall issue",
    "notes/exchange support",
    "notes issue",
    "backup support",
    "backup issue",
    "intel issue",
    "intel support",
    "oracle issue",
    "oracle support",
    "devops/bigdata issue",
    "unix issue",
    "unix support"
]

var problema_reportado = [
    "job issue",
    "script issue",
    "printer issue",
    "restore request",
    "application issue",
    "server reboot",
    "reboot server",
    "validação de backup"
]

var analise_do_acionamento = [
    "acionamento ism indevido",
    "indevido",
    "severidade indevida",
    "sem chamado",
    "sla indevida"
]

var acao_ism = [
    "monitoracao/report",
    "acompanhar",
    "priorizar"
]

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);

        worksheet.getCell(AREAS_ENVOLVIDAS + 1).value = "Areas envolvidas"
        worksheet.getCell(SERVICE_LINE + 1).value = "Service line"
        worksheet.getCell(PROBLEMA_REPORTADO + 1).value = "Problema reportado"
        worksheet.getCell(ANALISE_ACIONAMENTO + 1).value = "Análise do Acionamento"
        worksheet.getCell(LABELS_ALEATORIAS + 1).value = "Labels aleatorias"
        worksheet.getCell(ACAO_ISM + 1).value = "Ação ISM"

        i = 2
        while (i <= worksheet.rowCount) {
            var valor_celula = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
            var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
            var cliente = worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            if (valor_celula != null) {
                var pieces = valor_celula.split(",");
                var k = 0;
                var res = "";
                worksheet.getCell(AREAS_ENVOLVIDAS + i).value = " "
                worksheet.getCell(SERVICE_LINE + i).value = " "
                worksheet.getCell(PROBLEMA_REPORTADO + i).value = " "
                worksheet.getCell(ANALISE_ACIONAMENTO + i).value = " "
                worksheet.getCell(LABELS_ALEATORIAS + i).value = " "
                worksheet.getCell(ACAO_ISM + i).value = " "
                while (k < pieces.length) {
                    // if (pieces[k].trim().toLowerCase().includes("acionamento")) {
                    //     res += pieces[k].trim() + ", "
                    // }
                    var found = false;
                    var label = pieces[k].trim().toLowerCase()

                    if (areas_envolvidas.indexOf(label) >= 0) {
                        worksheet.getCell(AREAS_ENVOLVIDAS + i).value = worksheet.getCell(AREAS_ENVOLVIDAS + i).value + label + ", "
                        found = true;
                    }
                    if (service_line.indexOf(label) >= 0) {
                        worksheet.getCell(SERVICE_LINE + i).value = worksheet.getCell(SERVICE_LINE + i).value + label + ", "
                        found = true;
                    }
                    if (problema_reportado.indexOf(label) >= 0) {
                        worksheet.getCell(PROBLEMA_REPORTADO + i).value = worksheet.getCell(PROBLEMA_REPORTADO + i).value + label + ", "
                        found = true;
                    }
                    if (analise_do_acionamento.indexOf(label) >= 0) {
                        worksheet.getCell(ANALISE_ACIONAMENTO + i).value = worksheet.getCell(ANALISE_ACIONAMENTO + i).value + label + ", "
                        found = true;
                    }
                    if (acao_ism.indexOf(label) >= 0) {
                        worksheet.getCell(ACAO_ISM + i).value = worksheet.getCell(ACAO_ISM + i).value + label + ", "
                        found = true;
                    }

                    if (!found) {
                        if (!label.trim().toLowerCase().includes(cliente.toLowerCase().trim()) &&
                            !label.trim().toLowerCase().includes("sev") &&
                            !label.trim().toLowerCase().includes("incident") &&
                            !label.trim().toLowerCase().includes("service request") &&
                            !label.trim().toLowerCase().includes("change") &&
                            !label.trim().toLowerCase().includes("sem chamado")) {
                            worksheet.getCell(LABELS_ALEATORIAS + i).value = worksheet.getCell(LABELS_ALEATORIAS + i).value + label + ", "
                        }
                    }
                    k++;
                }
                worksheet.getCell(AREAS_ENVOLVIDAS + i).value = worksheet.getCell(AREAS_ENVOLVIDAS + i).value.substr(0, worksheet.getCell(AREAS_ENVOLVIDAS + i).value.length - 2).trim()
                worksheet.getCell(SERVICE_LINE + i).value = worksheet.getCell(SERVICE_LINE + i).value.substr(0, worksheet.getCell(SERVICE_LINE + i).value.length - 2).trim()
                worksheet.getCell(PROBLEMA_REPORTADO + i).value = worksheet.getCell(PROBLEMA_REPORTADO + i).value.substr(0, worksheet.getCell(PROBLEMA_REPORTADO + i).value.length - 2).trim()
                worksheet.getCell(ANALISE_ACIONAMENTO + i).value = worksheet.getCell(ANALISE_ACIONAMENTO + i).value.substr(0, worksheet.getCell(ANALISE_ACIONAMENTO + i).value.length - 2).trim()
                worksheet.getCell(LABELS_ALEATORIAS + i).value = worksheet.getCell(LABELS_ALEATORIAS + i).value.substr(0, worksheet.getCell(LABELS_ALEATORIAS + i).value.length - 2).trim()
                worksheet.getCell(ACAO_ISM + i).value = worksheet.getCell(ACAO_ISM + i).value.substr(0, worksheet.getCell(ACAO_ISM + i).value.length - 2).trim()
                // worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = res+","+res2
                // worksheet.getCell(AREAS_ENVOLVIDAS + i).value = res
            }
            i++;
        }


        // var i = 2;

        // //setting the title of the column
        // // worksheet.getCell(STORE_MONTH + 1).value = "month"

        // while (i <= worksheet.rowCount) {
        //     var valor_celula = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
        //     var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
        //     var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
        //     var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     if (valor_celula != null && assignee != null) {
        //         var pieces = valor_celula.split(",");
        //         var k = 0;
        //         var res = "";
        //         while (k < pieces.length) {
        //             res += pieces[k].trim() + "_sev" + sev + "_" + type + ", "
        //             k++;
        //         }
        //         res = res.substr(0, res.length - 2);
        //         // worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = res+","+res2
        //         worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = res
        //     }

        //     i++;
        // }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
