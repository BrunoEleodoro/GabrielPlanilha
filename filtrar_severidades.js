require('dotenv').config({ path: 'config' })

var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// CONTROLLERS
const STORE_PRIMARY_LABELS_COLUMN = process.env.STORE_PRIMARY_LABELS_COLUMN
const DESCRIPTION_COLUMN = process.env.DESCRIPTION_COLUMN
const STORE_TYPE_COLUMN = process.env.STORE_TYPE_COLUMN
const STORE_CLIENT_COLUMN = process.env.STORE_CLIENT_COLUMN
const CLIENTS_COLUMN = process.env.CLIENTS_COLUMN
const SOURCE_FILE = process.env.SOURCE_FILE
const OUTPUT_FILE = process.env.OUTPUT_FILE
const WORKSHEET = process.env.WORKSHEET
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


workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        var total = 0;
        var clientes = []
        var clientes_sevs = []
        var non_sevs_indexes = []
        var itens_without_sev = []
        var total_sev1 = 0
        var total_sev2 = 0
        var total_sev3 = 0
        var total_sev4 = 0
        var last_index = 0

        // the first step is search for all the spreadsheet values for unknow severities
        while (i <= worksheet.rowCount) {
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            if (labels != null) {
                if (
                    !labels.toLowerCase().trim().includes("sev1") &&
                    !labels.toLowerCase().trim().includes("sev2") &&
                    !labels.toLowerCase().trim().includes("sev3") &&
                    !labels.toLowerCase().trim().includes("sev4")
                ) {
                    // if the severity is NOT 1,2,3 or 4, them generate a random number for that 
                    var x = Math.floor((Math.random() * 4) + 1)
                    worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value + ",gerado sev" + parseFloat(x);

                    // painting the cell when the severity is unknow
                    worksheet.getCell(STORE_SEVERITY_COLUNM + i).fill = {
                        type: 'pattern',
                        pattern: 'solid',
                        fgColor: { argb: 'FFFFFF00' },
                        bgColor: { argb: 'aa00ff' }
                    };
                }
            }
            i++;
        }

        // the second step is define the severity based on the labels
        i = 2;
        while (i <= worksheet.rowCount) {
            // getting the client 
            var cliente = worksheet.getCell(CLIENTS_COLUMN + i).value
            // create an array of clients
            var index_cliente = clientes.indexOf(cliente)
            var cliente_obj = {}
            if (index_cliente == -1) {
                clientes.push(cliente);
                clientes_sevs.push({
                    'sev1': 0,
                    'sev2': 0,
                    'sev3': 0,
                    'sev4': 0,
                    'non_sev': 0
                })
                // cliente_obj = clientes_sevs[0];
                index_cliente = clientes_sevs.length - 1
            }

            // Here comes the magic
            var labels = worksheet.getCell(STORE_PRIMARY_LABELS_COLUMN + i).value
            var severidade = "n";
            // checking if the labels is not empty
            if (labels != null) {
                // if the label cell contains "sev1", so the severity will be "1"
                if (labels.toLowerCase().trim().includes("sev1")) {
                    severidade = "1"
                    total_sev1 = total_sev1 + 1
                    clientes_sevs[index_cliente].sev1 = clientes_sevs[index_cliente].sev1 + 1
                }
                // if the label cell contains "sev2", so the severity will be "2"
                else if (labels.toLowerCase().trim().includes("sev2")) {
                    severidade = "2"
                    total_sev2 = total_sev2 + 1
                    clientes_sevs[index_cliente].sev2 = clientes_sevs[index_cliente].sev2 + 1
                }
                // if the label cell contains "sev3", so the severity will be "3"
                else if (labels.toLowerCase().trim().includes("sev3")) {
                    severidade = "3"
                    total_sev3 = total_sev3 + 1
                    clientes_sevs[index_cliente].sev3 = clientes_sevs[index_cliente].sev3 + 1
                }
                // if the label cell contains "sev4", so the severity will be "4"
                else if (labels.toLowerCase().trim().includes("sev4")) {
                    severidade = "4"
                    total_sev4 = total_sev4 + 1
                    clientes_sevs[index_cliente].sev4 = clientes_sevs[index_cliente].sev4 + 1
                }
            }

            if (severidade != "n") {
                //saving the severity value in the colunm
                worksheet.getCell(STORE_SEVERITY_COLUNM + i).value = parseFloat(severidade)
                total = total + 1
                last_index = i
            }

            i++;
        }

        // Creating the summary for each severity
        worksheet.getCell(SEV_SUMMARY_LABELS + 3).value = "sev1"
        worksheet.getCell(SEV_SUMMARY_VALUES + 3).value = total_sev1

        worksheet.getCell(SEV_SUMMARY_LABELS + 4).value = "sev2"
        worksheet.getCell(SEV_SUMMARY_VALUES + 4).value = total_sev2

        worksheet.getCell(SEV_SUMMARY_LABELS + 5).value = "sev3"
        worksheet.getCell(SEV_SUMMARY_VALUES + 5).value = total_sev3

        worksheet.getCell(SEV_SUMMARY_LABELS + 6).value = "sev4"
        worksheet.getCell(SEV_SUMMARY_VALUES + 6).value = total_sev4

        worksheet.getCell(SEV_SUMMARY_LABELS + 7).value = "no sev"
        worksheet.getCell(SEV_SUMMARY_VALUES + 7).value = worksheet.rowCount - total

        console.log('sev1', total_sev1)
        console.log('sev2', total_sev2)
        console.log('sev3', total_sev3)
        console.log('sev4', total_sev4)
        console.log(total, worksheet.rowCount)


        // Creating the summary for each client
        worksheet.getCell(SEV_SUMMARY_CLIENT_NAME + 2).value = "Client name"
        worksheet.getCell(SEV_SUMMARY_CLIENT_SEV1 + 2).value = "SEV1"
        worksheet.getCell(SEV_SUMMARY_CLIENT_SEV2 + 2).value = "SEV2"
        worksheet.getCell(SEV_SUMMARY_CLIENT_SEV3 + 2).value = "SEV3"
        worksheet.getCell(SEV_SUMMARY_CLIENT_SEV4 + 2).value = "SEV4"
        var i = 0;
        while (i < clientes.length) {
            var pos = i + 3;
            worksheet.getCell(SEV_SUMMARY_CLIENT_NAME + (pos)).value = clientes[i]

            worksheet.getCell(SEV_SUMMARY_CLIENT_SEV1 + (pos)).value = clientes_sevs[i].sev1
            worksheet.getCell(SEV_SUMMARY_CLIENT_SEV2 + (pos)).value = clientes_sevs[i].sev2
            worksheet.getCell(SEV_SUMMARY_CLIENT_SEV3 + (pos)).value = clientes_sevs[i].sev3
            worksheet.getCell(SEV_SUMMARY_CLIENT_SEV4 + (pos)).value = clientes_sevs[i].sev4
            // worksheet.getCell('AP' + (pos)).value = clientes_sevs[i].non_sev
            i++;
        }

        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
