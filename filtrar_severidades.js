var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('Metricas Maio.xlsx')
    .then(function () {
        var worksheet = workbook.getWorksheet("Dados");
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
        while (i <= worksheet.rowCount) {
            var cliente = worksheet.getCell('L' + i).value
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
            } else {

            }
            var valor_celula_af = worksheet.getRow(i).getCell(16).value
            var severidade = "n";
            if (valor_celula_af != null) {
                if (valor_celula_af.toLowerCase().trim().includes("sev1")) {
                    severidade = "1"
                    total_sev1 = total_sev1 + 1
                    clientes_sevs[index_cliente].sev1 = clientes_sevs[index_cliente].sev1 + 1
                } else if (valor_celula_af.toLowerCase().trim().includes("sev2")) {
                    severidade = "2"
                    total_sev2 = total_sev2 + 1
                    clientes_sevs[index_cliente].sev2 = clientes_sevs[index_cliente].sev2 + 1
                } else if (valor_celula_af.toLowerCase().trim().includes("sev3")) {
                    severidade = "3"
                    total_sev3 = total_sev3 + 1
                    clientes_sevs[index_cliente].sev3 = clientes_sevs[index_cliente].sev3 + 1
                } else if (valor_celula_af.toLowerCase().trim().includes("sev4")) {
                    severidade = "4"
                    total_sev4 = total_sev4 + 1
                    clientes_sevs[index_cliente].sev4 = clientes_sevs[index_cliente].sev4 + 1
                }
            }

            if (severidade != "n") {
                worksheet.getRow(i).getCell(34).value = parseFloat(severidade)
                total = total + 1
                last_index = i
            } else {
                var x = Math.floor((Math.random() * 4) + 1)
                worksheet.getRow(i).getCell(16).value = worksheet.getRow(i).getCell(16).value + ",gerado sev"+parseFloat(x);
                // clientes_sevs[index_cliente].non_sev = clientes_sevs[index_cliente].non_sev + 1
                worksheet.getRow(i).getCell(34).fill = {
                    type: 'pattern',
                    pattern: 'solid',
                    fgColor: { argb: 'FFFFFF00' },
                    bgColor: { argb: 'aa00ff' }
                };
                i = last_index;
            }


            i++;
        }
        worksheet.getRow(3).getCell(35).value = "sev1"
        worksheet.getRow(3).getCell(36).value = total_sev1

        worksheet.getRow(4).getCell(35).value = "sev2"
        worksheet.getRow(4).getCell(36).value = total_sev2

        worksheet.getRow(5).getCell(35).value = "sev3"
        worksheet.getRow(5).getCell(36).value = total_sev3

        worksheet.getRow(6).getCell(35).value = "sev4"
        worksheet.getRow(6).getCell(36).value = total_sev4

        worksheet.getRow(7).getCell(35).value = "no sev"
        worksheet.getRow(7).getCell(36).value = worksheet.rowCount - total

        console.log('sev1', total_sev1)
        console.log('sev2', total_sev2)
        console.log('sev3', total_sev3)
        console.log('sev4', total_sev4)
        console.log(total, worksheet.rowCount)

        var i = 0;
        while (i < clientes.length) {
            var pos = i + 2;
            worksheet.getCell('AK' + (pos)).value = clientes[i]

            worksheet.getCell('AL' + (pos)).value = clientes_sevs[i].sev1
            worksheet.getCell('AM' + (pos)).value = clientes_sevs[i].sev2
            worksheet.getCell('AN' + (pos)).value = clientes_sevs[i].sev3
            worksheet.getCell('AO' + (pos)).value = clientes_sevs[i].sev4
            // worksheet.getCell('AP' + (pos)).value = clientes_sevs[i].non_sev
            i++;
        }





        return workbook.xlsx.writeFile('new.xlsx');
    })
