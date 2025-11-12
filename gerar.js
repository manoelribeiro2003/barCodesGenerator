import { createCanvas } from "canvas";
import JsBarcode from "jsbarcode";
import fs from "fs";
import * as xsls from 'xlsx';

const fileBuffer = fs.readFileSync('./Pasta5.xlsx');
const workbook = xsls.read(fileBuffer, { type: 'buffer' });
const sheetName = workbook.SheetNames[0];
const sheet = workbook.Sheets[sheetName];
const data = xsls.utils.sheet_to_json(sheet, { header: 1 });
const codigos = data.map((row) => row[0])

const set = new Set()


while (set.size < 10) {
    const n = Math.floor(Math.random() * 900000) + 100000
    if (!codigos.includes(n)) {
        set.add(n)
    }
}
const nCodigos = [...set]

console.log(nCodigos, set.size);



/*codigos.forEach((codigo) => {
    const canvas = createCanvas();
    JsBarcode(canvas, codigo, {
        format: "EAN13",
        displayValue: true,
        lineColor: "#000",
        width: 2,
        height: 80
    });

    const buffer = canvas.toBuffer("image/png");
    fs.writeFileSync(`codigo_${codigo}.png`, buffer);
    console.log(`✅ Código ${codigo} gerado`);
});*/