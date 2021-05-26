require('dotenv').config({ path: 'config' })
const utf8 = require('utf8');
const config = require('./load_columns');
var moment = require('moment')

var Excel = require('exceljs');
var workbook = new Excel.Workbook();
var not_allowed = [];

// Sev1, sev2, sev3, sev4, sem severidade
// numero de chamados que a pessoa trabalhou, 
// CH, SR,

// READ WORKBOOK
workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);

        worksheet.getCell(config.STORE_QUANTIDADE_TICKETS + 1).value = "Quantidade de tickets"
        worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + 1).value = "Quantidade de tickets per user"

        var horario_incidente_map = {}
        var i = 2;
        while (i <= worksheet.rowCount) {
            var title = worksheet.getCell(config.STORE_TITLE_COLUMN + i).value
            var created_at = worksheet.getCell(config.CREATED_AT + i).value
            var card_identifier = worksheet.getCell("AN" + i).value;
            var horario_incidente = worksheet.getCell(config.HORARIO_INCIDENTE + i).value;

            worksheet.getCell(config.STORE_QUANTIDADE_TICKETS + i).value = parseFloat("1")
            worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + i).value = parseFloat("1")

            if (horario_incidente_map[horario_incidente] == null) {
                horario_incidente_map[horario_incidente] = []
            }
            horario_incidente_map[horario_incidente].push(i)

            i++;

        }

        var i = 0;
        var incidentes = Object.keys(horario_incidente_map);
        var limite = incidentes.length;
        while (i < limite) {
            var horario_incidente_indexes = horario_incidente_map[incidentes[i]]

            //primeira verificacao
            //Verificacao: se o horario de encerramento estiver 
            //presente em somente um dos registros, entao este e o card original
            var k = 0;
            var quantidade_horario_encerramento = 0;
            while (k < horario_incidente_indexes.length) {
                var index = horario_incidente_indexes[k]
                var horario_encerramento = worksheet.getCell(config.HORARIO_ENCERRAMENTO + index).value;
                if (horario_encerramento != null) {
                    quantidade_horario_encerramento++;
                }
                if (quantidade_horario_encerramento > 1) {
                    break
                }
                k++;
            }

            //caso o hora de encerramento esteja 
            //presente em mais de uma celula, 
            //colocar 1 nas celulas que contem hora de encerramento.
            if (quantidade_horario_encerramento > 1) {
                var k = 0;
                while (k < horario_incidente_indexes.length) {
                    var index = horario_incidente_indexes[k]
                    var horario_encerramento = worksheet.getCell(config.HORARIO_ENCERRAMENTO + index).value;
                    var ism_solicitou_validacao = worksheet.getCell(config.ISM_SOLICITOU + index).value;
                    if (horario_encerramento != null) {

                        //Caso a coluna AG tenha o texto 
                        //“Nao, problema continua sendo tratado”, 
                        //automaticamente colocar 0 na coluna quantidade de 
                        //tickets per user.
                        if (ism_solicitou_validacao == "Nao, problema continua sendo tratado.") {
                            worksheet.getCell(config.QUANTIDADE_TICKETS_PER_USER + index).value = "0"
                        }
                    }
                    k++;
                }
            }

            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })