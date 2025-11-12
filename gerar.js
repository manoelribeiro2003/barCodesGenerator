import { createCanvas } from "canvas";
import JsBarcode from "jsbarcode";
import fs from "fs";
import * as xlsx from "xlsx";

// Ler o arquivo Excel
const fileBuffer = fs.readFileSync("./Pasta5.xlsx");
const workbook = xlsx.read(fileBuffer, { type: "buffer" });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = xlsx.utils.sheet_to_json(sheet, { header: 1 });

// Extrair primeira coluna e limpar
const bens = data
  .map((row) => row[0])
  .filter(Boolean)
  .map(String);

// Gerar 2 novos códigos de 6 dígitos que não existam no Excel
const set = new Set();

while (set.size < 2) {
  const n = String(Math.floor(Math.random() * 900000) + 100000); // 6 dígitos
  if (!bens.includes(n)) {
    set.add(n);
  }
}

const codigos = [...set];
console.log("Códigos gerados:", codigos);

// Criar códigos de barras com JsBarcode (CODE128)
codigos.forEach((codigo) => {
  const canvas = createCanvas();
  JsBarcode(canvas, codigo, {
    format: "CODE128", // permite qualquer sequência numérica
    displayValue: true,
    lineColor: "#000",
    width: 2,
    height: 80,
  });

  const buffer = canvas.toBuffer("image/png");
  fs.writeFileSync(`codigo_${codigo}.png`, buffer);
  console.log(`✅ Código ${codigo} gerado`);
});
