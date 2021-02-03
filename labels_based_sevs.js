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

var categoria_problema = []
categoria_problema["application issue"] = "application"
categoria_problema["network issue"] = "network"
categoria_problema["ca application issue"] = "application"
categoria_problema["ecommerce issue"] = "application"
categoria_problema["f5 application issue"] = "application"
categoria_problema["ftp issue"] = "application"
// categoria_problema["interface issue"] = "application"
categoria_problema["intranet prd app"] = "application"
categoria_problema["odi application"] = "application"
categoria_problema["peoplesoft app issue"] = "application"
categoria_problema["peoplesoft trace request"] = "application"
categoria_problema["rdf issue"] = "application"
categoria_problema["replica de ficha"] = "application"
categoria_problema["roadnet issue"] = "application"
categoria_problema["runbook application issue"] = "application"
categoria_problema["soa application issue"] = "application"
categoria_problema["stop/start service"] = "application"
categoria_problema["tasi issue"] = "application"
categoria_problema["tibico application"] = "application"
categoria_problema["diferimento brf"] = "application"
categoria_problema["invoice issue"] = "application"
categoria_problema["lentidao no sap"] = "application"
categoria_problema["sap issue"] = "application"
categoria_problema["sap transport"] = "application"
categoria_problema["archieve issue"] = "backup"
categoria_problema["backup request"] = "user request"
categoria_problema["restore follow up"] = "backup"
categoria_problema["restore request"] = "user request"
categoria_problema["add disk approval"] = "capacity"
// categoria_problema["filesystem full issue"] = "capacity"
categoria_problema["filesystem mount issue"] = "capacity"
categoria_problema["performance issue"] = "capacity"
categoria_problema["banco de loja issue"] = "database"
categoria_problema["database creation request"] = "database"
categoria_problema["database down"] = "database"
categoria_problema["database export request"] = "user request"
categoria_problema["database issue"] = "database"
categoria_problema["database locked id"] = "database"
categoria_problema["execucao script"] = "database"
categoria_problema["job backup rerun"] = "user request"
categoria_problema["job issue"] = "application"
categoria_problema["session kill request"] = "database"
categoria_problema["tablespace issue"] = "database"
// categoria_problema["disk full issue"] = "disk space"
categoria_problema["high cpu workload issue"] = "capacity"
categoria_problema["increased disk space"] = "capacity"
categoria_problema["email/exchange issue"] = "e-mail"
categoria_problema["parada eletrica"] = "infrastructury"
categoria_problema["power outage issue"] = "infrastructury"
categoria_problema["site cliente fora"] = "application"
categoria_problema["link issue"] = "network"
categoria_problema["vpn issue"] = "network"
categoria_problema["antivirus issue"] = "security"
categoria_problema["certified issue"] = "security"
categoria_problema["firewall issue"] = "security"
categoria_problema["firewall rule request"] = "user request"
categoria_problema["password reset"] = "user request"
categoria_problema["shared id locked"] = "security"
categoria_problema["uat approval request"] = "user request"
categoria_problema["user access issue"] = "security"
categoria_problema["printer issue"] = "servers"
categoria_problema["server down issue"] = "servers"
categoria_problema["server hung"] = "servers"
categoria_problema["server issue"] = "servers"
categoria_problema["server reboot"] = "servers"
categoria_problema["snapshot request"] = "user request"
categoria_problema["citrix issue"] = "application"
categoria_problema["mainframe issue"] = "application"
categoria_problema["maximo issue"] = "tools"
categoria_problema["monitoring issue"] = "tools"
categoria_problema["softlayer issue"] = "network"
categoria_problema["change open request"] = "user request"
categoria_problema["file creation request"] = "user request"
categoria_problema["file creation request"] = "user request"
categoria_problema["file user access"] = "user request"
categoria_problema["status request"] = "user request"
categoria_problema["validacao ambiente"] = "user request"
categoria_problema["backup issue"] = "backup issue"
categoria_problema["server unliked"] = "database"
categoria_problema["voip issue"] = "network"
categoria_problema["dns issue"] = "network"
categoria_problema["server unreachable"] = "servers"
categoria_problema["hardware issue"] = "servers"
categoria_problema["server access issue"] = "servers"
categoria_problema["user profile creation"] = "user request"
categoria_problema["shared folder access"] = "user request"
categoria_problema["info request"] = "user request"
categoria_problema["shared folder creation"] = "user request"
categoria_problema["cancelamento de job"] = "user request"
categoria_problema["vmware creation request"] = "user request"
categoria_problema["security issue"] = "security"

categoria_problema["control m issue"] = "issue"
categoria_problema["space issue"] = "issue"
categoria_problema["softlayer issue"] = "issue"
categoria_problema["firewall rule creation"] = "user request"
categoria_problema["vmware creation"] = "vmware creation"
categoria_problema["file transfer request"] = "user request"
categoria_problema["monitoracao/report"] = "monitoracao/report"

categoria_problema["sharepoint issue"] = "application"
categoria_problema["crtl m issue"] = "application"
categoria_problema["space issue"] = "capacity"
categoria_problema["oracle issue"] = "database"
categoria_problema["server unreachable"] = "servers"
categoria_problema["hardware issue"] = "servers"
categoria_problema["server access issue"] = "servers"
categoria_problema["user profile creation"] = "user request"
categoria_problema["info request"] = "user request"
categoria_problema["shared folder creation"] = "user request"
categoria_problema["validacao backup"] = "user request"
categoria_problema["cancelamento de job"] = "user request"
categoria_problema["server memory issue"] = "capacity"
categoria_problema["sql issue"] = "database"
categoria_problema["server unlinked"] = "application"
categoria_problema["odi application issue"] = "application"

var service_line = [
    "adabas support",
    "gbs support",
    "san disk support",
    "suporte local",
    "at&t support",
    "automation support",
    "storage baas",
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
    "storage san support",
    "adabas support",
    "service now support",
    "siebel support",
    "as400 support",
    "basis support",
    "tws support",
    "tss support",
    "notes support",
    "imi support"
]

var problema_reportado = [
    "server memory issue",
    "sql issue",
    "server unlinked",
    "odi application issue",
    "application issue",
    "ca application issue",
    "network issue",
    "ecommerce issue",
    "f5 application issue",
    "ftp issue",
    "security issue",
    "monitoracao/report",
    // "interface issue",
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
    "control m issue",
    "space issue",
    "oracle issue",
    "softlayer issue",
    "voip issue",
    "dns issue",
    "firewall rule creation",
    "user profile creation",
    "server unreachable",
    "hardware issue",
    "server access issue",
    "info request",
    "shared folder creation",
    "cancelamento de job",
    "vmware creation",
    "server unliked",
    // "filesystem full issue",
    "filesystem mount issue",
    "performance issue",
    "banco de loja issue",
    "database creation request",
    "database down",
    "database export request",
    "database issue",
    "database locked id",
    "execucao script",
    "job backup rerun",
    "job issue",
    "session kill request",
    "tablespace issue",
    // "disk full issue",
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
    "file creation request",
    "file transfer request",
    "file user access",
    "status request",
    "validacao ambiente",
    "sharepoint issue",
    "crtl m issue",
    "backup issue",
    "space issue",


    "oracle issue",
    "server unliked",
    "voip issue",
    "dns issue",
    "server unreachable",
    "hardware issue",
    "server access issue",
    "user profile creation",
    "shared folder access",
    "info request",
    "shared folder creation",
    "validacao backup",
    "cancelamento de job",
    "vmware creation request",
]

var acao_ism = [
    "acompanhar",
    "priorizar",
]

var canal_acionamento = [
    "acionamento via email",
    "acionamento via crit",
    "acionamento via slack",
    "acionamento via telefone",
    "acionamento indevido",
    "acionamento de hypercare"
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
    "acionamento dpm",
    "acionamento service desk",
    "acionamento sil",
    "acionamento sme",
    "acionamento tec br",
    "acionamento tec in",
    "acionamento imi",
    "acionamento squad leader",
    "acionamento gbs",
    "acionamento tec local"

]

var quem_te_acionou = [
    "acionado por cliente",
    "acionado por dpe",
    "acionado por duty manager",
    "acionado por gp",
    "acionado por dpm",
    "acionado por service desk",
    "acionado por sil",
    "acionado por sme",
    "acionado por tec br",
    "acionado por tec in",
    "acionado por gcc support",
    "acionado por producao",
    "acionado por dpm",
    "acionado por gbs",
    "acionado por squad leader",
    "acionado por imi",
    "acionado por gcc br"

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
    "change close task dpm",
    "change open request",
    "emergency change",
    "change status"
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
                    label = label.replace("  ", " ")

                    if (label.includes("acionamento t")) {
                        label = convert(label)
                        label = label.replace("ý", "é");
                    }
                    // var checkCategoria = checkIfCategoriaIsPresent(label);
                    // if (checkCategoria != "") {

                    //     if (worksheet.getCell(CATEGORIA + i).value.length == 1) {
                    //         worksheet.getCell(CATEGORIA + i).value = worksheet.getCell(CATEGORIA + i).value + checkCategoria + ", "
                    //         found = true;
                    //     }
                    // }

                    if (service_line.indexOf(label) >= 0) {
                        worksheet.getCell(SERVICE_LINE + i).value = worksheet.getCell(SERVICE_LINE + i).value + label + ", "
                        found = true;
                    }
                    if (problema_reportado.indexOf(label) >= 0) {
                        worksheet.getCell(PROBLEMA_REPORTADO + i).value = worksheet.getCell(PROBLEMA_REPORTADO + i).value + label + ", "
                        // worksheet.getCell(CATEGORIA + i).value = categoria_problema[label] != null ? categoria_problema[label].toString() : ""
                        // console.log('oi beatriz', label)
                        worksheet.getCell(CATEGORIA + i).value = categoria_problema[label].toString()

                        // console.log('problema_reportado', label, categoria_problema[label])
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
                    if (i == 2) {
                        console.log(label)
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
                // worksheet.getCell(CATEGORIA + i).value = worksheet.getCell(CATEGORIA + i).value.substr(0, worksheet.getCell(CATEGORIA + i).value.length - 2).trim()
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
