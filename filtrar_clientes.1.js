require('dotenv').config()

var Excel = require('exceljs');
var workbook = new Excel.Workbook();

// CONTROLLERS
const LABELS_COLUMN = process.env.LABELS_COLUMN
const DESCRIPTION_COLUMN = process.env.DESCRIPTION_COLUMN
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

// READ WORKBOOK
workbook.xlsx.readFile(SOURCE_FILE)
    .then(function () {
        //getting worksheet
        var worksheet = workbook.getWorksheet(WORKSHEET);
        var i = 2;
        var total = 0;
        while (i <= worksheet.rowCount) {
            
            //getting value from cell K
            var valor_celula_k = worksheet.getCell(DESCRIPTION_COLUMN+i).value
            
            //getting value from cell P
            var valor_celula_p = worksheet.getCell(LABELS_COLUMN+i).value
            
            //checking if the cell value is not empty
            if (valor_celula_k != null) {

                // getting the all before the `-` dash
                valor_celula_k = valor_celula_k.split("-")[0]
                
                // removing the brackets
                valor_celula_k = valor_celula_k.replace("[", "").replace("]", "")

                // removing the blank space from the beggining and from the end
                valor_celula_k = valor_celula_k.trim()

                //checking if the cell value is not empty
                if (valor_celula_p != null) {

                    // breaking in pieces the value coming from cell P, the description
                    var partes_p = valor_celula_p.split(",");
                    var k = 0;
                    var achei = false;

                    //for each piece of the description, check if the client name is present there.
                    while (k < partes_p.length) {

                        // if the client name exists in some part of the description
                        if (partes_p[k].toLowerCase().includes(valor_celula_k.toLowerCase())) {
                            achei = true;
                            // set the cell value to the name of the client.
                            worksheet.getCell(STORE_CLIENT_COLUMN+i).value = partes_p[k].toUpperCase().trim()
                            break
                        }
                        k++;
                    }

                    // If I didn't find, the client name will be all before the `-` slash
                    if (!achei) {
                        worksheet.getCell(STORE_CLIENT_COLUMN+i).value = valor_celula_k
                        worksheet.getRow(i).commit()
                    }
                    // if (worksheet.getCell(STORE_CLIENT_COLUMN+i).value == worksheet.getCell("L"+i).value) {
                    //     total = total+1;
                    //     worksheet.getCell("AG"+i).fill = {
                    //         type: 'pattern',
                    //         pattern: 'darkTrellis',
                    //         fgColor: { argb: 'FFFFFF00' },
                    //         bgColor: { argb: 'FF0000FF' }
                    //     };
                    // }

                }

            }

            i++;
        }
        // console.log(total, worksheet.rowCount)
        console.log('finalizado!');
        return workbook.xlsx.writeFile(OUTPUT_FILE);
    })
