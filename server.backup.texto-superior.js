const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx");
const fs = require("fs");
const path = require("path");
const { exec } = require("child_process");

const app = express();
const PORT = 3000;

const srcPath = path.join(__dirname, "src");
const excelPath = path.join(srcPath, "data.xlsx");
const outputPath = path.join(srcPath, "output.zpl");

const upload = multer({
  dest: path.join(__dirname, "uploads"),
});

app.use(express.static(path.join(__dirname, "public")));
app.use(express.json());

function limpiarTexto(valor) {
  if (valor === undefined || valor === null) return "";
  return String(valor).trim();
}

function obtenerUbicacion(row) {
  const posiblesColumnas = [
    "Nombre de la ubicación",
    "Nombre de la ubicacion",
    "Nombre de la ubicaci├│n",
    "Ubicacion",
    "Ubicación",
    "Ubicaci├│n",
    "location",
    "Location",
    "ubicacion",
    "ubicación",
  ];

  for (const columna of posiblesColumnas) {
    const valor = limpiarTexto(row[columna]);
    if (valor) return valor;
  }

  const valores = Object.values(row)
    .map((v) => limpiarTexto(v))
    .filter((v) => v);

  return valores.length ? valores[0] : "";
}

function leerUbicacionesDesdeExcel() {
  if (!fs.existsSync(excelPath)) {
    throw new Error("No se encontro el archivo data.xlsx en la carpeta src.");
  }

  const workbook = XLSX.readFile(excelPath);
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  // Leer el Excel como matriz para no perder la primera fila.
  // Esto permite usar archivos sin encabezado.
  const filas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });

  const ubicaciones = [];

  filas.forEach((fila) => {
    const primeraCelda = limpiarTexto(fila[0]);

    if (!primeraCelda) return;

    // Si algún día el Excel trae encabezado, lo ignoramos.
    const textoNormalizado = primeraCelda.toLowerCase();

    if (
      textoNormalizado === "nombre de la ubicacion" ||
      textoNormalizado === "nombre de la ubicación" ||
      textoNormalizado === "ubicacion" ||
      textoNormalizado === "ubicación"
    ) {
      return;
    }

    ubicaciones.push(primeraCelda);
  });

  return ubicaciones;
}

function filtrarUbicaciones(ubicaciones, modo, cantidad) {
  if (modo === "cantidad") {
    const limite = parseInt(cantidad, 10);

    if (!limite || limite <= 0) {
      throw new Error("Debes indicar una cantidad valida mayor que 0.");
    }

    return ubicaciones.slice(0, limite);
  }

  return ubicaciones;
}

function numero(valor, defecto) {
  const n = parseFloat(valor);
  return Number.isFinite(n) ? n : defecto;
}

function obtenerConfig(config = {}) {
  return {
    labelWidth: 609,
    labelHeight: 406,

    showTopText: config.showTopText !== false,
    showBottomText: config.showBottomText !== false,

    globalX: numero(config.globalX, 0),

    textTopY: numero(config.textTopY, 45),
    textTopSize: numero(config.textTopSize, 68),

    barcodeX: numero(config.barcodeX, 120),
    barcodeY: numero(config.barcodeY, 145),
    barcodeModule: numero(config.barcodeModule, 3.5),
    barcodeRatio: numero(config.barcodeRatio, 2),
    barcodeHeight: numero(config.barcodeHeight, 135),

    textBottomY: numero(config.textBottomY, 310),
    textBottomSize: numero(config.textBottomSize, 30),
  };
}

function crearZPL(location, config = {}) {
  const c = obtenerConfig(config);

  const textoSuperior = c.showTopText
    ? `
^FX ===== TEXTO SUPERIOR =====
^CF0,${c.textTopSize},${c.textTopSize}
^FO${c.globalX},${c.textTopY}^FB${c.labelWidth},1,0,C,0^FD${location}^FS
`
    : "";

  const textoInferior = c.showBottomText
    ? `
^FX ===== TEXTO INFERIOR =====
^CF0,${c.textBottomSize},${c.textBottomSize}
^FO${c.globalX},${c.textBottomY}^FB${c.labelWidth},1,0,C,0^FD${location}^FS
`
    : "";

  return `
^XA
^PW${c.labelWidth}
^LL${c.labelHeight}
^LH0,0

${textoSuperior}

^FX ===== CODIGO DE BARRAS =====
^BY${c.barcodeModule},${c.barcodeRatio},${c.barcodeHeight}
^FO${c.barcodeX + c.globalX},${c.barcodeY}^BCN,${c.barcodeHeight},N,N,N
^FD${location}^FS

${textoInferior}

^XZ
`;
}

function generarZPL(modo = "todas", cantidad = null, config = {}) {
  const printScale = numero(config.printScale, 1);
  const todasLasUbicaciones = leerUbicacionesDesdeExcel();
  const ubicaciones = filtrarUbicaciones(todasLasUbicaciones, modo, cantidad);

  fs.writeFileSync(outputPath, "", "utf8");

  ubicaciones.forEach((location) => {
    fs.appendFileSync(outputPath, crearZPL(location, config), "utf8");
  });

  return {
    totalExcel: todasLasUbicaciones.length,
    total: ubicaciones.length,
    ubicaciones,
    primeraUbicacion: ubicaciones[0] || "",
    primerZpl: ubicaciones[0] ? crearZPL(ubicaciones[0], config) : "",
    config: obtenerConfig(config),
  };
}

app.post("/subir-excel", upload.single("excel"), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({
        ok: false,
        message: "No se recibio ningun archivo.",
      });
    }

    fs.copyFileSync(req.file.path, excelPath);
    fs.unlinkSync(req.file.path);

    const resultado = generarZPL("todas", null, {});

    if (resultado.total === 0) {
      return res.status(400).json({
        ok: false,
        message: "El Excel se cargo, pero no se encontraron ubicaciones.",
      });
    }

    res.json({
      ok: true,
      message: `Excel cargado correctamente. Total etiquetas en Excel: ${resultado.totalExcel}`,
      ...resultado,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.post("/preparar", (req, res) => {
  try {
    const { modo, cantidad, config } = req.body;
    const resultado = generarZPL(modo, cantidad, config);

    if (resultado.total === 0) {
      return res.status(400).json({
        ok: false,
        message: "No se preparo ninguna etiqueta.",
      });
    }

    res.json({
      ok: true,
      message: `Preparado correctamente. Etiquetas a imprimir: ${resultado.total}`,
      ...resultado,
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.post("/imprimir", (req, res) => {
  try {
    const { modo, cantidad, config } = req.body;

    const resultado = generarZPL(modo, cantidad, config);

    if (resultado.total === 0) {
      return res.status(400).json({
        ok: false,
        message: "No hay etiquetas para imprimir.",
      });
    }

    const command = `powershell -ExecutionPolicy Bypass -File ".\\print-raw.ps1" -PrinterName "ZDesigner ZD421-203dpi ZPL" -FilePath ".\\src\\output.zpl"`;

    exec(command, { cwd: __dirname }, (error, stdout, stderr) => {
      const salida = `${stdout || ""} ${stderr || ""}`.trim();

      if (error) {
        return res.status(500).json({
          ok: false,
          message: `No se pudo imprimir. Detalle: ${salida || error.message}`,
        });
      }

      res.json({
        ok: true,
        message: `Etiquetas enviadas correctamente a la Zebra. Total impreso: ${resultado.total}`,
        stdout,
        ...resultado,
      });
    });
  } catch (error) {
    res.status(500).json({
      ok: false,
      message: error.message,
    });
  }
});

app.listen(PORT, () => {
  console.log(`Interfaz disponible en: http://localhost:${PORT}`);
});


