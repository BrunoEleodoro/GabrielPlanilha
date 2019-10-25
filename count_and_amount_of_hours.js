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
const CREATED_AT = process.env.CREATED_AT
const CARD_ASSIGNEES = process.env.CARD_ASSIGNEES
const STORE_WORKED_HOURS = process.env.STORE_WORKED_HOURS

const AREAS_ENVOLVIDAS = process.env.AREAS_ENVOLVIDAS
const SERVICE_LINE = process.env.SERVICE_LINE
const PROBLEMA_REPORTADO = process.env.PROBLEMA_REPORTADO
const ANALISE_ACIONAMENTO = process.env.ANALISE_ACIONAMENTO
const LABELS_ALEATORIAS = process.env.LABELS_ALEATORIAS
const ACAO_ISM = process.env.ACAO_ISM
const MEIO_COMUNICACAO = process.env.MEIO_COMUNICACAO
const SOLICITACOES = process.env.SOLICITACOES
const QUEM_VOCE_ACIONOU = process.env.QUEM_VOCE_ACIONOU
const QUEM_TE_ACIONOU = process.env.QUEM_TE_ACIONOU
const LABELS_RELACIONADO_A_CHANGE = process.env.LABELS_RELACIONADO_A_CHANGE

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

const STORE_QUANTIDADE_TICKETS = process.env.STORE_QUANTIDADE_TICKETS
const QUANTIDADE_TICKETS_PER_USER = process.env.QUANTIDADE_TICKETS_PER_USER

const SEV_SUMMARY_LABELS = process.env.SEV_SUMMARY_LABELS
const SEV_SUMMARY_VALUES = process.env.SEV_SUMMARY_VALUES
const SEV_SUMMARY_CLIENT_NAME = process.env.SEV_SUMMARY_CLIENT_NAME
const SEV_SUMMARY_CLIENT_SEV1 = process.env.SEV_SUMMARY_CLIENT_SEV1
const SEV_SUMMARY_CLIENT_SEV2 = process.env.SEV_SUMMARY_CLIENT_SEV2
const SEV_SUMMARY_CLIENT_SEV3 = process.env.SEV_SUMMARY_CLIENT_SEV3
const SEV_SUMMARY_CLIENT_SEV4 = process.env.SEV_SUMMARY_CLIENT_SEV4
const SOURCE_COLUMNS_LIST = process.env.SOURCE_COLUMNS_LIST
const DESTINATION_COLUMNS_LIST = process.env.DESTINATION_COLUMNS_LIST

// Sev1, sev2, sev3, sev4, sem severidade
// numero de chamados que a pessoa trabalhou, 
// CH, SR,

var people = []
people["gabriela ferreira dias dos santos"] = "gfsantos@br.ibm.com"
people["lucas gaspar hoffelder"] = "lgaspar@br.ibm.com"
people["otavio de almeida sambo"] = "osambo@br.ibm.com"
people["catia harume yamamoto"] = "catiay@br.ibm.com"
people["jacqueline cristina da silva"] = "jacquecs@br.ibm.com"
people["gabriel siqueira"] = "gsiq@br.ibm.com"
people["renan diego mafeis"] = "renandm@br.ibm.com"
people["diego dayvison alves de araujo ferreira"] = "diegoaf@br.ibm.com"
people["matheus reis villela"] = "mrvilela@br.ibm.com"
people["lalisa viola faria santos"] = "lalisavi@br.ibm.com"
people["mariana rangel vieira valim"] = "valimvrm@br.ibm.com"

function replaceEmailToName(email) {
    email = email.toLowerCase()
    var i = 0;
    var keys = Object.keys(people);
    while (i < keys.length) {
        if (people[keys[i]] == email) {
            return keys[i]
        }
        i++;
    }
}

var months_name = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var todos_emails = []
        var todos_emails_lista = []
        var i = 2;
        var todas_as_labels = []

        var integrantes_ism = []
        var clientes = []
        var types = []
        var sevs = []
        var days = []
        var months_years = []
        var years = []
        var titles = []
        worksheet.getCell(STORE_QUANTIDADE_TICKETS + 1).value = "Quantidade de tickets"
        worksheet.getCell(QUANTIDADE_TICKETS_PER_USER + 1).value = "Quantidade de tickets per user"
        while (i <= worksheet.rowCount) {

            // var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
            var client = worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            // var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
            // var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
            // var data = worksheet.getCell(CREATED_AT + i).value
            var title = worksheet.getCell(STORE_TITLE_COLUMN + i).value
            var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
            var time_worked = worksheet.getCell(STORE_WORKED_HOURS + i).value

            worksheet.getCell(QUANTIDADE_TICKETS_PER_USER + i).value = '1'

            if (title != null) {
                if (title.includes("(copy)")) {
                    worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value = '0'
                } else {
                    worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value = '1'
                }
            }
            // var index = clientes.indexOf(client);
            // if (client != null && index == -1) {
            //     clientes.push(client)
            //     titles[client] = []
            //     titles[client].push({ title: title, assignee: assignee, worked_hours: time_worked })
            // } else if (client != null && index >= 0) {
            //     titles[client].push({ title: title, assignee: assignee, worked_hours: time_worked })
            // }

            // if (assignee != null && integrantes_ism.indexOf(assignee) == -1) {
            //     integrantes_ism.push(assignee)
            // }
            i++;

        }
        // var titles_originals = JSON.parse(JSON.stringify(titles));
        // // console.log(titles)

        // var repetidos = []

        // var i = 0;
        // while (i < titles.length) {
        //     var title = titles[i];
        //     var res = check(titles_originals, title)
        //     if (res.length > 0) {
        //         repetidos[title] = res
        //     }
        //     // if (titles.indexOf(title) != null) {
        //     //     if (repetidos[title] == null) {
        //     //         repetidos[title] = 1;
        //     //     } else {
        //     //         repetidos[title] = repetidos[title] + 1
        //     //     }
        //     //     delete titles[titles.indexOf(title)]
        //     //     // repetidos[title] = 1;
        //     // }
        //     i++;
        // }

        // var i = 2;
        // while (i <= worksheet.rowCount) {
        //     var client = worksheet.getCell(STORE_CLIENT_COLUMN + i).value
        //     var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
        //     var title = worksheet.getCell(STORE_TITLE_COLUMN + i).value
        //     var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     var time_worked = worksheet.getCell(STORE_WORKED_HOURS + i).value

        //     i++;
        // }

        // var i = 0;
        // var keys = Object.keys(repetidos)
        // while (i < keys.length) {
        //     var k = 0;
        //     while (k < repetidos[keys[i]].length) {
        //         var index_title_repetido = repetidos[keys[i]][k];
        //         var client = worksheet.getCell(STORE_CLIENT_COLUMN + (index_title_repetido + 2)).value
        //         var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + (index_title_repetido + 2)).value
        //         var title = worksheet.getCell(STORE_TITLE_COLUMN + (index_title_repetido + 2)).value
        //         var assignee = worksheet.getCell(CARD_ASSIGNEES + (index_title_repetido + 2)).value
        //         var time_worked = worksheet.getCell(STORE_WORKED_HOURS + (index_title_repetido + 2)).value

        //         if (k > 2) {
        //             worksheet.getCell(STORE_QUANTIDADE_TICKETS + (index_title_repetido + 2)).value = '0'
        //         } else {
        //             if (client == "CARREFOUR" && labels.includes("issue")) {
        //                 worksheet.getCell(STORE_QUANTIDADE_TICKETS + (index_title_repetido + 2)).value = '1'
        //             } else {
        //                 worksheet.getCell(STORE_QUANTIDADE_TICKETS + (index_title_repetido + 2)).value = '0'
        //             }
        //         }

        //         k++;
        //     }
        //     i++;
        // }

        var i = 2;
        while (i <= worksheet.rowCount) {
            var client = worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var title = worksheet.getCell(STORE_TITLE_COLUMN + i).value
            var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
            var time_worked = worksheet.getCell(STORE_WORKED_HOURS + i).value

            if (worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value == null || worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value == "") {
                worksheet.getCell(STORE_QUANTIDADE_TICKETS + i).value = '1'
            }
            i++;
        }
        // console.log(repetidos)

        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })

function check(titles_param, title) {
    let titles = JSON.parse(JSON.stringify(titles_param))
    var res = [];
    var i = 0

    while (titles.indexOf(title) != -1) {
        // console.log(title, titles.indexOf(title))
        // console.log(i)
        if (i > 0) {
            res.push(titles.indexOf(title))
        }
        delete titles[titles.indexOf(title)]
        i++;
    }
    return res
    // var i = 0;
    // while (i < titles.length) {
    //     var title = titles[i];

    //     if (titles.indexOf(title) != null) {
    //         if (repetidos[title] == null) {
    //             repetidos[title] = 1;
    //         } else {
    //             repetidos[title] = repetidos[title] + 1
    //         }
    //         delete titles[titles.indexOf(title)]
    //         // repetidos[title] = 1;
    //     }
    //     i++;
    // }
}
