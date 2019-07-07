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


        while (i <= worksheet.rowCount) {

            var assignee = worksheet.getCell(CARD_ASSIGNEES + i).value
            var client = worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
            var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value

            if (assignee != null && assignee.includes(",")) {
                var k = 0;
                var pieces = assignee.split(",")
                while (k < pieces.length) {
                    integrantes_ism.push(replaceEmailToName(pieces[k].trim()));
                    clientes.push(client);
                    types.push(type);
                    sevs.push(sev);
                    k++;
                }
            } else if (assignee != null) {
                integrantes_ism.push(replaceEmailToName(assignee.trim()));
                clientes.push(client);
                types.push(type);
                sevs.push(sev);
            }

            // if (labels != null && type != null && sev != null) {
            //     type = type.toString().trim().toLowerCase()
            //     sev = sev.toString().trim().toLowerCase()

            //     labels = labels.split(",");
            //     var k = 0;
            //     while (k < labels.length) {
            //         if (todas_as_labels.indexOf(labels[k].toLowerCase().trim()) == -1) {
            //             todas_as_labels.push(labels[k].toLowerCase().trim())
            //         }
            //         k++;
            //     }
            // }
            i++;

        }
        console.log(integrantes_ism.length, clientes.length, types.length, sevs.length)

        // var i = 2;
        // while (i <= worksheet.rowCount) {
        //     var assignees = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     if (assignees != null) {
        //         if (assignees.includes(",")) {
        //             var k = 0;
        //             while (k < assignees.split(',').length) {
        //                 if (todos_emails_lista.indexOf(assignees.split(',')[k].trim()) == -1) {
        //                     todos_emails_lista.push(assignees.split(',')[k].trim())
        //                 }
        //                 k++;
        //             }
        //         }
        //     } else {
        //         todos_emails_lista.push(assignees.trim())
        //     }
        //     i++;
        // }


        //get the users
        // var i = 2;
        // while (i <= worksheet.rowCount) {
        //     var assignees = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     if (assignees != null) {
        //         if (assignees.includes(",")) {
        //             var k = 0;
        //             while (k < assignees.split(',').length) {
        //                 if (todos_emails.indexOf(assignees.split(',')[k].trim()) == -1) {
        //                     todos_emails.push(assignees.split(',')[k].trim())
        //                 }
        //                 k++;
        //             }
        //         }
        //     } else {
        //         if (assignees != null && todos_emails.indexOf(assignees.trim()) == -1) {
        //             todos_emails.push(assignees.trim())
        //         }
        //     }
        //     i++;
        // }

        // var todos_os_assignee = []
        // var relacao_assignee_labels = []
        // i = 0
        // while (i <= todos_emails.length) {
        //     // var assignees = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     var email = todos_emails[i]
        //     var k = 0;
        //     relacao_assignee_labels[email] = []
        //     while (k < todas_as_labels.length) {
        //         var label = todas_as_labels[k];
        //         relacao_assignee_labels[email][label] = 0
        //         k++;
        //     }
        //     i++;
        // }


        // i = 2;
        // while (i <= worksheet.rowCount) {
        //     var assignees = worksheet.getCell(CARD_ASSIGNEES + i).value
        //     var primary_labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
        //     var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
        //     var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
        //     if (type != null && sev != null) {
        //         type = type.toString().trim().toLowerCase()
        //         sev = sev.toString().trim().toLowerCase()
        //     }

        //     if (primary_labels != null && assignees != null && assignees.includes(",")) {
        //         var list_assignees = assignees.split(",")
        //         labels = primary_labels.split(",");
        //         var j = 0;
        //         while (j < list_assignees.length) {
        //             var assignee = list_assignees[j].trim();
        //             var k = 0;
        //             while (k < labels.length) {
        //                 var label = labels[k].toLowerCase().trim()
        //                 // console.log(relacao_assignee_labels[assignee][label], assignee, label)
        //                 relacao_assignee_labels[assignee][label] = parseFloat(relacao_assignee_labels[assignee][label]) + 1
        //                 k++;
        //             }


        //             j++;
        //         }
        //     } else {
        //         if (primary_labels != null && assignees != null) {
        //             labels = primary_labels.split(",");
        //             var k = 0;
        //             while (k < labels.length) {
        //                 var label = labels[k].toLowerCase().trim()
        //                 relacao_assignee_labels[assignees.trim()][label] = parseFloat(relacao_assignee_labels[assignees.trim()][label]) + 1
        //                 // console.log(relacao_assignee_labels[assignees][label], assignees, label)
        //                 k++;
        //             }
        //         }
        //     }

        //     i++;
        // }

        var column_starts_at = 23
        // console.log(todos_os_clientes[2], relacao_clientes_labels[todos_os_clientes[2]])
        var i = 0;
        while (i < integrantes_ism.length) {
            // worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            worksheet.getRow(i + 93).getCell(column_starts_at).value = integrantes_ism[i]
            i++;
        }

        

        worksheet.getRow(92).getCell(23).value = "Integrante ISM"
        worksheet.getRow(92).getCell(24).value = "Cliente"
        worksheet.getRow(92).getCell(25).value = "Type"
        worksheet.getRow(92).getCell(26).value = "Severidade"
        worksheet.getRow(92).getCell(27).value = "Media"

        var i = 0;
        while (i < clientes.length) {
            // worksheet.getCell(STORE_CLIENT_COLUMN + i).value
            worksheet.getRow(i + 93).getCell(24).value = clientes[i]
            worksheet.getRow(i + 93).getCell(25).value = types[i]
            worksheet.getRow(i + 93).getCell(26).value = "sev"+sevs[i]
            worksheet.getRow(i + 93).getCell(27).value = parseFloat(1)
            i++;
        }
        // var i = 0;
        // while (i < todos_emails.length) {
        //     // worksheet.getCell(STORE_CLIENT_COLUMN + i).value
        //     var type = worksheet.getCell(STORE_TYPE_COLUMN + (i + 2)).value
        //     var sev = worksheet.getCell(STORE_SEVERITY_COLUNM + (i + 2)).value
        //     if (type != null && sev != null) {
        //         type = type.toString().trim().toLowerCase()
        //         sev = sev.toString().trim().toLowerCase()
        //     }
        //     var client = todos_emails[i].trim()
        //     var k = 0;
        //     while (k < todas_as_labels.length) {
        //         var label = todas_as_labels[k].toLowerCase().trim()
        //         worksheet.getRow(i + 93).getCell(24 + k).value = parseFloat(relacao_assignee_labels[client][label])
        //         // console.log(relacao_assignee_labels[client][label], client, label)
        //         k++;
        //     }

        //     i++;
        // }

        // console.log(relacao_assignee_labels['gsiq@br.ibm.com'])
        // console.log(JSON.stringify(relacao_assignee_labels['gsiq@br.ibm.com']))
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })


    /*
       var assignees = worksheet.getCell(CARD_ASSIGNEES + i).value
            var client = worksheet.getCell(CLIENTS_COLUMN + i).value
            var type = worksheet.getCell(STORE_TYPE_COLUMN + i).value
            var severidade = worksheet.getCell(STORE_SEVERITY_COLUNM + i).value
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
    */