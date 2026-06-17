const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");

const excelPath = path.join(__dirname, "data.xlsx");
const outputPath = path.join(__dirname, "output.zpl");

if (!fs.existsSync(excelPath)) {
  console.log("No se encontró el archivo data.xlsx en la carpeta src.");
  process.exit(1);
}

const workbook = XLSX.readFile(excelPath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

fs.writeFileSync(outputPath, "", "utf8");

data.forEach((row) => {
  const location = row["Nombre de la ubicación"] || "";

  if (!location) return;

  const zpl = `
^XA
^PW600
^LL400

^FX ===== TEXTO =====
^CF0,170,82
^FO35,50^FB540,1,0,C^FD${location}^FS

^FX ===== BARCODE (más ancho + más a la derecha) =====
^BY3,2,120
^FO70,210^BCN,120,N,N,N
^FD${location}^FS

^XZ
`;

  fs.appendFileSync(outputPath, zpl, "utf8");
});

console.log("Archivo output.zpl generado correctamente.");
console.log(outputPath);
