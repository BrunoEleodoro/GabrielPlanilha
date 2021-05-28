const utf8 = require('utf8');
const config = require('./load_columns');
var Excel = require('exceljs');
var workbook = new Excel.Workbook();

var moment = require('moment')
var problema_complexidade = {};
problema_complexidade["datacenter access request"] = "baixa"
problema_complexidade["backup issue"] = "baixa"
problema_complexidade["sql issue"] = "media"
problema_complexidade["printer creation request"] = "baixa"
problema_complexidade["server memory issue"] = "baixa"
problema_complexidade["odi application issue"] = "media"
problema_complexidade["application issue"] = "media"
problema_complexidade["network issue"] = "baixa"
problema_complexidade["ecommerce issue"] = "alta"
problema_complexidade["f5 application issue"] = "media"
problema_complexidade["ftp issue"] = "baixa"
problema_complexidade["security issue"] = "alta"
problema_complexidade["monitoracao/report"] = "baixa"
problema_complexidade["intranet prd app"] = "media"
problema_complexidade["peoplesoft app issue"] = "media"
problema_complexidade["peoplesoft trace request"] = "baixa"
problema_complexidade["rdf issue"] = "media"
problema_complexidade["replica de ficha"] = "alta"
problema_complexidade["roadnet issue"] = "baixa"
problema_complexidade["runbook application issue"] = "media"
problema_complexidade["application issue"] = "media"
problema_complexidade["stop/start service"] = "baixa"
problema_complexidade["tasi issue"] = "media"
problema_complexidade["tibico application"] = "media"
problema_complexidade["diferimento brf"] = "baixa"
problema_complexidade["invoice issue"] = "alta"
problema_complexidade["lentidao no sap"] = "alta"
problema_complexidade["sap issue"] = "alta"
problema_complexidade["sap transport"] = "media"
problema_complexidade["archieve issue"] = "media"
problema_complexidade["backup request"] = "baixa"
problema_complexidade["restore follow up"] = "baixa"
problema_complexidade["restore request"] = "baixa"
problema_complexidade["add disk approval"] = "baixa"
problema_complexidade["control m issue"] = "media"
problema_complexidade["firewall rule creation"] = "baixa"
problema_complexidade["filesystem mount issue"] = "alta"
problema_complexidade["performance issue"] = "baixa"
problema_complexidade["banco de loja issue"] = "baixa"
problema_complexidade["database creation request"] = "baixa"
problema_complexidade["database down"] = "alta"
problema_complexidade["database export request"] = "media"
problema_complexidade["database issue"] = "media"
problema_complexidade["database locked id"] = "baixa"
problema_complexidade["execucao script"] = "baixa"
problema_complexidade["job backup rerun"] = "baixa"
problema_complexidade["job issue"] = "baixa"
problema_complexidade["session kill request"] = "baixa"
problema_complexidade["tablespace issue"] = "media"
problema_complexidade["high cpu workload issue"] = "baixa"
problema_complexidade["increased disk space"] = "media"
problema_complexidade["email/exchange issue"] = "media"
problema_complexidade["power outage issue"] = "baixa"
problema_complexidade["link issue"] = "media"
problema_complexidade["vpn issue"] = "alta"
problema_complexidade["antivirus issue"] = "media"
problema_complexidade["certified issue"] = "alta"
problema_complexidade["firewall issue"] = "media"
problema_complexidade["password reset"] = "baixa"
problema_complexidade["uat approval request"] = "baixa"
problema_complexidade["user access issue"] = "media"
problema_complexidade["printer issue"] = "media"
problema_complexidade["server down issue"] = "alta"
problema_complexidade["server hung"] = "baixa"
problema_complexidade["server issue"] = "media"
problema_complexidade["server reboot"] = "baixa"
problema_complexidade["snapshot request"] = "baixa"
problema_complexidade["citrix issue"] = "media"
problema_complexidade["mainframe issue"] = "alta"
problema_complexidade["maximo issue"] = "media"
problema_complexidade["monitoring issue"] = "baixa"
problema_complexidade["softlayer issue"] = "alta"
problema_complexidade["change open request"] = "baixa"
problema_complexidade["file creation request"] = "baixa"
problema_complexidade["file transfer request"] = "media"
problema_complexidade["file user access"] = "media"
problema_complexidade["status request"] = "baixa"
problema_complexidade["validacao ambiente"] = "baixa"
problema_complexidade["sharepoint issue"] = "alta"
problema_complexidade["space issue"] = "baixa"
problema_complexidade["oracle issue"] = "media"
problema_complexidade["server unlinked"] = "baixa"
problema_complexidade["voip issue"] = "media"
problema_complexidade["dns issue"] = "media"
problema_complexidade["server unreachable"] = "baixa"
problema_complexidade["hardware issue"] = "alta"
problema_complexidade["server access issue"] = "media"
problema_complexidade["user profile creation"] = "baixa"
problema_complexidade["shared folder access"] = "baixa"
problema_complexidade["info request"] = "baixa"
problema_complexidade["shared folder creation"] = "baixa"
problema_complexidade["validacao backup"] = "baixa"
problema_complexidade["cancelamento de job"] = "baixa"
problema_complexidade["vmware creation request"] = "media"
problema_complexidade["logs request"] = "media"

workbook.xlsx.readFile(config.SOURCE_FILE)
    .then(function () {
        var worksheet = workbook.getWorksheet(config.WORKSHEET);
        worksheet.getCell(config.COMPLEXIDADE + 1).value = "Complexidade do Chamado"
        var i = 2
        while (i <= worksheet.rowCount) {
            let problema_reportado = worksheet.getCell(config.PROBLEMA_REPORTADO + i).value;
            let type = worksheet.getCell(config.STORE_TYPE_COLUMN + i).value;
            worksheet.getCell(config.COMPLEXIDADE + i).value = problema_complexidade[problema_reportado]
            if (type.toUpperCase() == "CH") {
                worksheet.getCell(config.COMPLEXIDADE + i).value = "media";
            } else if (type.toUpperCase() == "REPORT") {
                worksheet.getCell(config.COMPLEXIDADE + i).value = "baixa";
            } else if (type.toUpperCase() == "SC") {
                worksheet.getCell(config.COMPLEXIDADE + i).value = "baixa";
            }
            i++;
        }

        console.log('finalizado!');
        return workbook.xlsx.writeFile(config.OUTPUT_FILE);
    })
