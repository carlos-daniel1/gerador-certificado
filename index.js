const { PDFDocument, StandardFonts, rgb } = require("pdf-lib");
const { writeFileSync, readFileSync } = require("fs");

var XLSX = require("xlsx");
var workbook = XLSX.readFile("excel/alunos.xlsx");

let planilha = workbook.Sheets[workbook.SheetNames[0]];


const numeroCertificados = 3

// pegar dados planilha
for (let index = 2; index < (numeroCertificados + 2); index++) {
  let RA = planilha[`A${index}`].v;
  let nome = planilha[`B${index}`].v;
  let curso = planilha[`C${index}`].v
  let duracao = planilha[`D${index}`].v
  let cidade = planilha[`E${index}`].v
  let data = planilha[`F${index}`].v

  let text = `Certificamos que o aluno: ${nome}, com registro acadêmico: ${RA} \nconcluiu o curso de ` +
    `${curso}, com carga horária de ${duracao}, \nna cidade de ${cidade}, data ${data}\n`

  createPDF(nome, text)

}


// modificar pdf, inserir dados aluno
async function createPDF(nome, text) {

  const document = await PDFDocument.load(readFileSync("./model_pdf/modelo-certificado.pdf"));

  const page = document.getPage(0);

  const fontSize =  17
  const { width, height } = page.getSize()

  const romanFont = await document.embedFont(StandardFonts.TimesRoman)
  const boldFont = await document.embedFont(StandardFonts.TimesRomanBold)

  page.drawText(text, {
    x: 50,
    y: height - 14 * fontSize,
    size: fontSize,
    font: romanFont,
    color: rgb(0, 0, 0),
  }
)

  writeFileSync(nome + ".pdf", await document.save());
}

