import { createCanvas } from "canvas";
import JsBarcode from "jsbarcode";
import fs from "fs";
import ExcelJS from "exceljs";
import * as xlsx from "xlsx";
import { PDFDocument, rgb } from "pdf-lib";
import readline from "readline";
import ora from "ora";

// Cria interface de leitura no terminal
const rl = readline.createInterface({
  input: process.stdin,
  output: process.stdout
});

// Função para perguntar algo no terminal
function perguntar(pergunta) {
  return new Promise(resolve => rl.question(pergunta, resolve));
}

// Ler o arquivo Excel existente
const pastaExcel = fs.readFileSync("./planilhas/BasePatrimonios.xlsx");
const workbookXlsx = xlsx.read(pastaExcel, { type: "buffer" });
const sheetName = workbookXlsx.SheetNames[0];
const sheet = workbookXlsx.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

const larguraPx = 151; // 4 cm
const alturaPx = 76;   // 2 cm

// Extrair primeira coluna
const bens = data
  .map(row => row[0])
  .filter(Boolean)
  .map(String);

console.log(`Existem ${bens.length} códigos já na planilha.`);

// Pergunta quantos códigos gerar
const quantidadeInput = await perguntar("Quantos códigos deseja gerar? ");
const quantidade = parseInt(quantidadeInput);

// Pergunta se quer salvar na base
const salvarResposta = await perguntar("Deseja salvar na Base de Patrimonios? (s/n) ");
const salvarNoMesmoArquivo = salvarResposta.toLowerCase() === "s";

// Fecha o readline
rl.close();

const spinnerCod = ora("Gerando códigos...").start();

// Gerar 10 códigos de 6 dígitos únicos
const set = new Set();
while (set.size < quantidade) {
  const n = String(Math.floor(Math.random() * 900000) + 100000);
  if (!bens.includes(n)) set.add(n);
}
const codigos = [...set];

spinnerCod.succeed("Códigos gerados!");

// Carregar workbook com ExcelJS
const workbook = new ExcelJS.Workbook();
await workbook.xlsx.load(pastaExcel);
const ws = workbook.getWorksheet(sheetName);

// Ajustar largura das colunas no Excel
ws.getColumn(2).width = 0.4;
ws.getColumn(3).width = 21.6;
ws.getColumn(4).width = 0.4;
ws.getColumn(5).width = 5;
ws.getColumn(6).width = 16;

const linhasImagem = 4;
const bordaPadrao = { style: "thin" };
const alturaBorda = 1;

// Descobrir última linha ocupada
let linha = ws.lastRow ? ws.lastRow.number + 2 : 1;

// 🔵 Armazena buffers das imagens para gerar o PDF depois
const imagens = [];

for (const codigo of codigos) {
  // Criar canvas
  const canvas = createCanvas(larguraPx, alturaPx);
  const ctx = canvas.getContext("2d");

  ctx.fillStyle = "#fff";
  ctx.fillRect(0, 0, larguraPx, alturaPx);

  JsBarcode(canvas, codigo, {
    format: "CODE128",
    displayValue: true,
    lineColor: "#000",
    width: 2,
    height: alturaPx - 20,
    margin: 6,
  });

  const buffer = canvas.toBuffer("image/png");
  imagens.push({ buffer, codigo });

  const imageId = workbook.addImage({ buffer, extension: "png" });

  const startRow = linha;
  const endRow = linha + linhasImagem - 1;

  ws.addImage(imageId, {
    tl: { col: 2, row: startRow - 1 },
    ext: { width: larguraPx, height: alturaPx },
  });

  ws.getRow(startRow - 1).height = alturaBorda;

  for (let c = 2; c <= 4; c++) {
    ws.getCell(`${String.fromCharCode(64 + c)}${startRow - 1}`).border = { top: bordaPadrao };
  }

  for (let r = startRow - 1; r <= endRow + 1; r++) {
    ws.getCell(`B${r}`).border = { left: bordaPadrao };
    ws.getCell(`D${r}`).border = { right: bordaPadrao };
  }

  ws.getRow(endRow + 1).height = alturaBorda;
  for (let c = 2; c <= 4; c++) {
    ws.getCell(`${String.fromCharCode(64 + c)}${endRow + 1}`).border = { bottom: bordaPadrao };
  }

  ws.getCell(`F${startRow + Math.floor(linhasImagem / 2)}`).value = codigo;

  linha = endRow + 4;
}

// Salvar Excel atualizado
if (salvarNoMesmoArquivo) {
  const spinnerPlanilhas = ora("Salvando planilhas...").start();
  await workbook.xlsx.writeFile("./planilhas/BasePatrimonios.xlsx");
  await workbook.xlsx.writeFile("./planilhas/CodigosGerados.xlsx");
  spinnerPlanilhas.succeed("Codigos gerados e salvos na base!");
} else {
  const spinnerPlanilhas = ora("Salvando planilha de codigos...").start();
  await workbook.xlsx.writeFile("./planilhas/CodigosGerados.xlsx");
  spinnerPlanilhas.succeed("Codigos gerados porém não salvos na base");
}

// ------------------------------------------------------------
// 🔴 AGORA GERA O PDF EM GRADE 3xN (SEM drawText)
// ------------------------------------------------------------

const spinnerPDF = ora("Salvando PDF...").start();

const pdf = await PDFDocument.create();
let page = pdf.addPage();

const pageWidth = page.getWidth();
const pageHeight = page.getHeight();

const colWidth = pageWidth / 3;  
const rowHeight = 120;           // altura das etiquetas no PDF

let x = 0;
let y = pageHeight - rowHeight - 20;

for (let i = 0; i < imagens.length; i++) {
  const { buffer } = imagens[i];
  const image = await pdf.embedPng(buffer);

  page.drawImage(image, {
    x: x + 10,
    y: y,
    width: larguraPx,
    height: alturaPx
  });

  x += colWidth;

  // Fecha a linha (3 colunas)
  if ((i + 1) % 3 === 0) {
    x = 0;
    y -= rowHeight + 40;

    // Cria nova página se necessário
    if (y < 50) {
      page = pdf.addPage();
      y = pageHeight - rowHeight - 20;
    }
  }
}

// Salvar PDF
const pdfBytes = await pdf.save();
fs.writeFileSync("./Etiquetas.pdf", pdfBytes);

spinnerPDF.succeed("PDF criado com sucesso: Etiquetas.pdf");
