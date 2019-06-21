var Excel = require('exceljs');
var workbook = new Excel.Workbook();

workbook.xlsx.readFile('Metricas Maio.xlsx')
    .then(function () {
        var worksheet = workbook.getWorksheet("Dados");
        var i = 0;
        var total = 0;
        while (i < worksheet.rowCount) {
            var valor_celula_k = worksheet.getRow(i).getCell(11).value
            var valor_celula_p = worksheet.getRow(i).getCell(16).value
            if (valor_celula_k != null) {

                valor_celula_k = valor_celula_k.split("-")[0]
                valor_celula_k = valor_celula_k.replace("[", "").replace("]", "")
                valor_celula_k = valor_celula_k.trim()

                if (valor_celula_p != null) {
                    var partes_p = valor_celula_p.split(",");
                    var k = 0;
                    var achei = false;
                    while (k < partes_p.length) {
                        if (partes_p[k].toLowerCase().includes(valor_celula_k.toLowerCase())) {
                            achei = true;
                            worksheet.getRow(i).getCell(33).value = partes_p[k].toUpperCase().trim()
                            break
                        }
                        k++;
                    }

                    if (!achei) {
                        worksheet.getRow(i).getCell(33).value = valor_celula_k
                        worksheet.getRow(i).commit()
                    }
                    if (worksheet.getRow(i).getCell(33).value == worksheet.getRow(i).getCell(12).value) {
                        total = total+1;
                        worksheet.getRow(i).getCell(34).fill = {
                            type: 'pattern',
                            pattern: 'darkTrellis',
                            fgColor: { argb: 'FFFFFF00' },
                            bgColor: { argb: 'FF0000FF' }
                        };
                    }

                }

            }

            i++;
        }
        console.log(total, worksheet.rowCount)
        return workbook.xlsx.writeFile('new.xlsx');
    })
