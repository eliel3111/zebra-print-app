const XLSX = require("xlsx");
const fs = require("fs");
const { exec } = require("child_process");
const os = require("os");

// 📍 CONFIG
const filePath = "data.xlsx";
const printerShare = `\\\\${os.hostname()}\\ZEBRA`;

// 📥 Leer Excel
const workbook = XLSX.readFile(filePath);
const sheet = workbook.Sheets[workbook.SheetNames[0]];
const data = XLSX.utils.sheet_to_json(sheet);

// 🖨️ Función imprimir
function printZPL(zpl) {
  fs.writeFileSync("label.zpl", zpl, "ascii");

  exec(`copy /B label.zpl "${printerShare}"`, (err) => {
    if (err) {
      console.error("Error:", err);
    } else {
      console.log("Etiqueta enviada");
    }
  });
}

data.reverse().forEach((row, index) => {
  const key = Object.keys(row)[0]; // toma la columna automáticamente
  const location = row[key];

  if (!location) return;
  const cleanLocation = location
  .replace(/\s+/g, '')     // quita espacios
  .replace(/[–—]/g, '-')   // normaliza guiones raros
  .trim();

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

  setTimeout(() => {
    printZPL(zpl);
  }, index * 500);
});