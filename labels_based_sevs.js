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

const CATEGORIA = process.env.CATEGORIA
const SERVICE_LINE = process.env.SERVICE_LINE
const PROBLEMA_REPORTADO = process.env.PROBLEMA_REPORTADO
const ANALISE_ACIONAMENTO = process.env.ANALISE_ACIONAMENTO
const LABELS_ALEATORIAS = process.env.LABELS_ALEATORIAS
const ACAO_ISM = process.env.ACAO_ISM
const CANAL_ACIONAMENTO = process.env.CANAL_ACIONAMENTO
const SOLICITACOES = process.env.SOLICITACOES
const QUEM_VOCE_ACIONOU = process.env.QUEM_VOCE_ACIONOU
const QUEM_TE_ACIONOU = process.env.QUEM_TE_ACIONOU
const LABELS_CHAMADOS_INDEVIDOS = process.env.LABELS_CHAMADOS_INDEVIDOS

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

var categorias = ["application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "application",
    "backup",
    "backup",
    "backup",
    "backup",
    "capacity",
    "capacity",
    "capacity",
    "capacity",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "database",
    "disk space",
    "disk space",
    "disk space",
    "e-mail",
    "infrastructury",
    "infrastructury",
    "infrastructury",
    "network",
    "network",
    "security",
    "security",
    "security",
    "security",
    "security",
    "security",
    "security",
    "security",
    "servers",
    "servers",
    "servers",
    "servers",
    "servers",
    "servers",
    "tools",
    "tools",
    "tools",
    "tools",
    "tools",
    "user request",
    "user request",
    "user request",
    "user request",
    "user request",
    "user request",]

var service_line = ["adabas support",
    "at&t support",
    "automation support",
    "backup support",
    "cloud support",
    "cms support",
    "db2 support",
    "devops/bigdata support",
    "exchange support",
    "mss support",
    "gcc support",
    "iam support",
    "intel support",
    "mainframe support",
    "middleware support",
    "network support",
    "oracle support",
    "maximo support",
    "unix support",
    "peoplesoft support",
    "sql support",
    "production support",
    "san disk support",
    "sap support",
    "tws support",
]

var problema_reportado = [
    "application issue",
    "ca application issue",
    "ecommerce issue",
    "f5 application issue",
    "ftp issue",
    "interface issue",
    "intranet prd app",
    "odi application",
    "peoplesoft app issue",
    "peoplesoft trace request",
    "rdf issue",
    "replica de ficha",
    "roadnet issue",
    "runbook application issue",
    "soa application issue",
    "stop/start service",
    "tasi issue",
    "tibico application",
    "diferimento brf",
    "invoice issue",
    "lentidao no sap",
    "sap issue",
    "sap transport",
    "archieve issue",
    "backup request",
    "restore follow up",
    "restore request",
    "add disk approval",
    "filesystem full issue",
    "filesystem mount issue",
    "performance issue",
    "banco de loja issue",
    "database creation",
    "database down",
    "database export request",
    "database issue",
    "database locked id",
    "execucao script",
    "job backup rerun",
    "job issue",
    "session kill request",
    "tablespace issue",
    "disk full issue",
    "high cpu workload issue",
    "increased disk space",
    "email/exchange issue",
    "parada eletrica",
    "power outage issue",
    "site cliente fora",
    "link issue",
    "vpn issue",
    "antivirus issue",
    "certified issue",
    "firewall issue",
    "firewall rule request",
    "password reset",
    "shared id locked",
    "uat approval request",
    "user access issue",
    "printer issue",
    "server down issue",
    "server hung",
    "server issue",
    "server reboot",
    "snapshot request",
    "citrix issue",
    "mainframe issue",
    "maximo issue",
    "monitoring issue",
    "softlayer issue",
    "change open request",
    "file creation",
    "file transfer",
    "file user access",
    "status request",
    "validacao ambiente",
]

var analise_do_acionamento = [
    "acionamento ism indevido",
    "indevido",
    "severidade indevido",
    "sem chamado",
    "dentro do sla",
    "sla breach",
    "sla indevido",
]

var acao_ism = [
    "monitoracao/report",
    "acompanhar",
    "priorizar",
]

var canal_acionamento = [
    "acionamento via email",
    "acionamento via sametime",
    "acionamento via slack",
    "acionamento via telefone",
]

var solicitacoes = [
    "backup request",
    "execucao job backup",
    "execucao script",
    "stop/start service",
    "restore request",
    "server reboot",
    "snapshot",
    "solicitacao status",
    "sap transport",
    "acompanhar restore",
    "criacao de pasta",
    "password reset",
    "validacao ambiente",
]

var quem_voce_acionou = [
    "acionamento cliente",
    "acionamento dpe",
    "acionamento duty manager",
    "acionamento gp",
    "acionamento sam",
    "acionamento service desk",
    "acionamento sil",
    "acionamento sme",
    "acionamento tec br",
    "acionamento tec in",
]

var quem_te_acionou = [
    "acionado por cliente",
    "acionado por dpe",
    "acionado por duty manager",
    "acionado por gp",
    "acionado por sam",
    "acionado por service desk",
    "acionado por sil",
    "acionado por sme",
    "acionado por tec br",
    "acionado por tec in",
    "acionado por gcc support",
    "acionado por producao",
    "acionamento indevido vvo",
    "acionamento indevido gpa",
]

var chamados_indevidos = [
    "change follow up",
    "change not reported",
    "change late task",
    "extentend change windows",
    "change checkpoint",
    "change approval",
    "change fallback",
    "change failed",
    "change close task sam",
    "change open request",
    "emergency change",
]

function convert(input) {
    if (input == null) {
        return input
    }
    var iconv = require('iconv-lite');
    var output = iconv.decode(input, "ISO-8859-1");
    // output = iconv.decode(output, "UTF-8");
    return output;
}

function checkIfCategoriaIsPresent(label) {
    let i = 0;
    var found = "";
    while (i < categorias.length) {
        var categoria = categorias[i];
        if (label.toLowerCase().includes(categoria.toLowerCase())) {
            found = categoria.toLowerCase()
            break;
        }
        i++;
    }
    return found;
}

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);

        worksheet.getCell(CATEGORIA + 1).value = "Categoria"
        worksheet.getCell(SERVICE_LINE + 1).value = "Service line"
        worksheet.getCell(PROBLEMA_REPORTADO + 1).value = "Problema reportado"
        // worksheet.getCell(ANALISE_ACIONAMENTO + 1).value = "Análise do Acionamento"
        worksheet.getCell(LABELS_ALEATORIAS + 1).value = "Labels aleatorias"
        // worksheet.getCell(ACAO_ISM + 1).value = "Ação ISM"
        worksheet.getCell(CANAL_ACIONAMENTO + 1).value = "Canal de Acionamento"
        // worksheet.getCell(SOLICITACOES + 1).value = "Solicitações"
        worksheet.getCell(QUEM_VOCE_ACIONOU + 1).value = "Quem você acionou"
        worksheet.getCell(QUEM_TE_ACIONOU + 1).value = "Quem te acionou"
        worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + 1).value = "Labels Relacionado a Change"

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
                worksheet.getCell(CATEGORIA + i).value = " "
                worksheet.getCell(SERVICE_LINE + i).value = " "
                worksheet.getCell(PROBLEMA_REPORTADO + i).value = " "
                // worksheet.getCell(ANALISE_ACIONAMENTO + i).value = " "
                worksheet.getCell(LABELS_ALEATORIAS + i).value = " "
                // worksheet.getCell(ACAO_ISM + i).value = " "
                worksheet.getCell(CANAL_ACIONAMENTO + i).value = " "
                // worksheet.getCell(SOLICITACOES + i).value = " "
                worksheet.getCell(QUEM_VOCE_ACIONOU + i).value = " "
                worksheet.getCell(QUEM_TE_ACIONOU + i).value = " "
                worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value = " "
                while (k < pieces.length) {
                    // if (pieces[k].trim().toLowerCase().includes("acionamento")) {
                    //     res += pieces[k].trim() + ", "
                    // }
                    var found = false;
                    var label = pieces[k].trim().toLowerCase()

                    if (label.includes("acionamento t")) {
                        label = convert(label)
                        label = label.replace("ý", "é");
                    }
                    var checkCategoria = checkIfCategoriaIsPresent(label);
                    if (checkCategoria != "") {

                        if (worksheet.getCell(CATEGORIA + i).value.length == 1) {
                            worksheet.getCell(CATEGORIA + i).value = worksheet.getCell(CATEGORIA + i).value + checkCategoria + ", "
                            found = true;
                        }
                    }
                    if (service_line.indexOf(label) >= 0) {
                        worksheet.getCell(SERVICE_LINE + i).value = worksheet.getCell(SERVICE_LINE + i).value + label + ", "
                        found = true;
                    }
                    if (problema_reportado.indexOf(label) >= 0) {
                        worksheet.getCell(PROBLEMA_REPORTADO + i).value = worksheet.getCell(PROBLEMA_REPORTADO + i).value + label + ", "
                        found = true;
                    }
                    // if (analise_do_acionamento.indexOf(label) >= 0) {
                    //     worksheet.getCell(ANALISE_ACIONAMENTO + i).value = worksheet.getCell(ANALISE_ACIONAMENTO + i).value + label + ", "
                    //     found = true;
                    // }
                    // if (acao_ism.indexOf(label) >= 0) {
                    //     worksheet.getCell(ACAO_ISM + i).value = worksheet.getCell(ACAO_ISM + i).value + label + ", "
                    //     found = true;
                    // }
                    if (canal_acionamento.indexOf(label) >= 0) {
                        worksheet.getCell(CANAL_ACIONAMENTO + i).value = worksheet.getCell(CANAL_ACIONAMENTO + i).value + label + ", "
                        found = true;
                    }

                    // if (solicitacoes.indexOf(label) >= 0) {
                    //     worksheet.getCell(SOLICITACOES + i).value = worksheet.getCell(SOLICITACOES + i).value + label + ", "
                    //     found = true;
                    // }
                    if (quem_voce_acionou.indexOf(label) >= 0) {
                        worksheet.getCell(QUEM_VOCE_ACIONOU + i).value = worksheet.getCell(QUEM_VOCE_ACIONOU + i).value + label + ", "
                        found = true;
                    }
                    if (quem_te_acionou.indexOf(label) >= 0) {
                        worksheet.getCell(QUEM_TE_ACIONOU + i).value = worksheet.getCell(QUEM_TE_ACIONOU + i).value + label + ", "
                        found = true;
                    }
                    if (chamados_indevidos.indexOf(label) >= 0) {
                        worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value = worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value + label + ", "
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
                worksheet.getCell(CATEGORIA + i).value = worksheet.getCell(CATEGORIA + i).value.substr(0, worksheet.getCell(CATEGORIA + i).value.length - 2).trim()
                worksheet.getCell(SERVICE_LINE + i).value = worksheet.getCell(SERVICE_LINE + i).value.substr(0, worksheet.getCell(SERVICE_LINE + i).value.length - 2).trim()
                worksheet.getCell(PROBLEMA_REPORTADO + i).value = worksheet.getCell(PROBLEMA_REPORTADO + i).value.substr(0, worksheet.getCell(PROBLEMA_REPORTADO + i).value.length - 2).trim()
                // worksheet.getCell(ANALISE_ACIONAMENTO + i).value = worksheet.getCell(ANALISE_ACIONAMENTO + i).value.substr(0, worksheet.getCell(ANALISE_ACIONAMENTO + i).value.length - 2).trim()
                worksheet.getCell(LABELS_ALEATORIAS + i).value = worksheet.getCell(LABELS_ALEATORIAS + i).value.substr(0, worksheet.getCell(LABELS_ALEATORIAS + i).value.length - 2).trim()
                // worksheet.getCell(ACAO_ISM + i).value = worksheet.getCell(ACAO_ISM + i).value.substr(0, worksheet.getCell(ACAO_ISM + i).value.length - 2).trim()
                worksheet.getCell(CANAL_ACIONAMENTO + i).value = worksheet.getCell(CANAL_ACIONAMENTO + i).value.substr(0, worksheet.getCell(CANAL_ACIONAMENTO + i).value.length - 2).trim()
                // worksheet.getCell(SOLICITACOES + i).value = worksheet.getCell(SOLICITACOES + i).value.substr(0, worksheet.getCell(SOLICITACOES + i).value.length - 2).trim()
                worksheet.getCell(QUEM_VOCE_ACIONOU + i).value = worksheet.getCell(QUEM_VOCE_ACIONOU + i).value.substr(0, worksheet.getCell(QUEM_VOCE_ACIONOU + i).value.length - 2).trim()
                worksheet.getCell(QUEM_TE_ACIONOU + i).value = worksheet.getCell(QUEM_TE_ACIONOU + i).value.substr(0, worksheet.getCell(QUEM_TE_ACIONOU + i).value.length - 2).trim()
                worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value = worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value.substr(0, worksheet.getCell(LABELS_CHAMADOS_INDEVIDOS + i).value.length - 2).trim()
                // worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = res+","+res2
                // worksheet.getCell(CATEGORIA + i).value = res

                if (type == "CH") {
                    worksheet.getCell(CATEGORIA + i).value = "N/A - CHANGE"
                    worksheet.getCell(SERVICE_LINE + i).value = "N/A - CHANGE"
                    worksheet.getCell(PROBLEMA_REPORTADO + i).value = "N/A - CHANGE"
                } else if (type == "REPORT") {
                    worksheet.getCell(CATEGORIA + i).value = "N/A - REPORT"
                    worksheet.getCell(SERVICE_LINE + i).value = "N/A - REPORT"
                    worksheet.getCell(PROBLEMA_REPORTADO + i).value = "N/A - REPORT"
                }
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
