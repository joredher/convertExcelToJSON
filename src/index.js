let XLSX = require("xlsx")
let fs = require("fs")
const eliminarDiacriticosEs = (texto) => {
    return texto
        .normalize('NFD')
        .replace(/([^n\u0300-\u036f]|n(?!\u0303(?![\u0300-\u036f])))[\u0300-\u036f]+/gi,"$1")
        .normalize()
}

const ExcelAJSON = () => {
  const excel = XLSX.readFile(
    "C:/laragon/www/convertExcelToJSON/datos.xlsx"
  );
  var nombreHoja = excel.SheetNames; // regresa un array
  let datos = XLSX.utils.sheet_to_json(excel.Sheets[nombreHoja[0]]);

  const jDatos = [];

  datos.forEach((item) => {
    if (!jDatos.find(x => x.departamento === item.DEPARTAMENTO)) {
        jDatos.push({"departamento": item.DEPARTAMENTO, "centrosPoblados": {}})
    }

    let keyMunicipio = item.MUNICIPIO.toLowerCase()
    keyMunicipio = eliminarDiacriticosEs(keyMunicipio)
    keyMunicipio = keyMunicipio.replace(/\s(.)/g, function($1) { return $1.toUpperCase(); }).replace(/\s/g, '').replace(/^(.)/, function($1) { return $1.toLowerCase(); });
    let indexKey = keyMunicipio.replace('ñ','ni')

    jDatos.forEach(value => {
        if (value.departamento === item.DEPARTAMENTO) {
            if (value.centrosPoblados[indexKey]) {
                value.centrosPoblados[indexKey].push(item['CENTRO POBLADO'])
            } else {
                value.centrosPoblados[indexKey] = [item['CENTRO POBLADO']] 
            }
        }
    })
  })

  fs.writeFile("centrosPoblados.json", JSON.stringify(jDatos), 'utf-8', (err) => {
    if (err) throw err
    console.log('El archivo se ha guardado.')
  })

};
ExcelAJSON();