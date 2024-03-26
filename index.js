var XLSX = require("xlsx");
var workbook = XLSX.readFile("excel/dados.xlsx");

let planilha = workbook.Sheets[workbook.SheetNames[0]];


 for (let index = 2; index < 5; index++) {
    const id = planilha[`A${index}`].v;
    const name = planilha[`B${index}`].v;

   console.log({id:id, name:name})
 }