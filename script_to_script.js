function convertDoubleToTime(value) {
    var partes = value.split('.');
    return partes[0].padStart(2, '0') + ':' + partes[1].padStart(2, '0')
}

function fixTime(time) {
    var partes = time.toFixed(2).toString().split('.');
    if (partes[1] >= 60) {
        return parseFloat((parseFloat(partes[0]) + 1) + "." + (parseFloat(partes[1]) - 60))
    } else {
        return time;
    }
}

var i = 0;
var hora_inicial = 0.10;
var hora_final = 0.30;
var horariopico = 0;

while (i < 70) {

    console.log('if (pieces >= ' + hora_inicial.toFixed(2) + ' && pieces < ' + hora_final.toFixed(2) + ') {')
    console.log('   worksheet.getCell(HORARIO_PICO + i).value = "' + convertDoubleToTime(horariopico.toFixed(2)) + '"')
    console.log('}')

    hora_inicial = hora_inicial + 0.20
    hora_final = hora_final + 0.20
    horariopico = horariopico + 0.20

    // console.log('before', hora_inicial)
    // console.log('before', hora_final)
    // console.log('before', horariopico)

    hora_inicial = fixTime(hora_inicial);
    hora_final = fixTime(hora_final);
    horariopico = fixTime(horariopico);

    // console.log('after', hora_inicial)
    // console.log('after', hora_final)
    // console.log('after', horariopico)

    i++;
}



