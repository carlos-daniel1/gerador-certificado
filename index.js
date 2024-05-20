const { PDFDocument, StandardFonts, rgb} = require("pdf-lib");
const { writeFileSync } = require("fs");

var XLSX = require("xlsx");
var workbook = XLSX.readFile("excel/alunos.xlsx");

let planilha = workbook.Sheets[workbook.SheetNames[0]];

for (let index = 2; index < 3; index++) {
  let RA = planilha[`A${index}`].v;
  let nome = planilha[`B${index}`].v;
  let curso = planilha[`C${index}`].v
  let duracao = planilha[`D${index}`].v
  let cidade = planilha[`E${index}`].v
  let data = planilha[`F${index}`].v

  let text = `Certificamos que o aluno: ${nome}, RA: ${RA} concluiu o curso de\n`+
  `${curso}, com carga horária de ${duracao}, na cidade de ${cidade}, data ${data}\n`
  
    //createPDF(name)

    console.log(text);
 }

async function createPDF(user) {
  const PDFdoc = await PDFDocument.create();
  const page = PDFdoc.addPage();
  
  const fontSize = 25
  const { width, height } = page.getSize()

  const timesRomanFont = await PDFdoc.embedFont(StandardFonts.TimesRoman)

  page.drawText(`Hello my name is ${user}`, {
  x: 50,
  y: height - 4 * fontSize,
  size: fontSize,
  font: timesRomanFont,
  color: rgb(0, 0.53, 0.71),
})

  writeFileSync(user + ".pdf", await PDFdoc.save());
}
