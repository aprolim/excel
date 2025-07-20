const express = require('express');
const multer = require('multer');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.static('public'));

// Configuración de Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => cb(null, file.originalname)
});
const upload = multer({ storage });

// Función mejorada para asignar códigos
function assignCodes(dataA, dataB) {
  const valuesA = dataA.map(row => row[0]).filter(Boolean);
  const codesB = dataB.map(row => row[0]).filter(Boolean);

  // Contar frecuencia de cada valor
  const frequencyMap = {};
  valuesA.forEach(value => {
    frequencyMap[value] = (frequencyMap[value] || 0) + 1;
  });

  const result = [];
  let codeIndex = 0;

  // Procesar cada valor único
  Object.keys(frequencyMap).forEach(key => {
    const count = frequencyMap[key];
    const codesNeeded = Math.ceil(count / 8);
    const codes = codesB.slice(codeIndex, codeIndex + codesNeeded);
    codeIndex += codesNeeded;

    if (codes.length === 0) {
      console.warn(`No hay suficientes códigos para el valor: ${key}`);
      return;
    }

    const base = Math.floor(count / codes.length);
    const remainder = count % codes.length;

    // Distribuir equitativamente
    codes.forEach((code, i) => {
      const times = i < remainder ? base + 1 : base;
      for (let j = 0; j < times; j++) {
        result.push({ Valor: key, Codigo: code });
      }
    });
  });

  return result;
}

// Ruta para procesar archivos
app.post('/process', upload.fields([
  { name: 'fileA', maxCount: 1 },
  { name: 'fileB', maxCount: 1 }
]), (req, res) => {
  try {
    // Leer archivos
    const workbookA = XLSX.readFile(req.files['fileA'][0].path);
    const workbookB = XLSX.readFile(req.files['fileB'][0].path);

    // Obtener datos
    const sheetA = workbookA.Sheets[workbookA.SheetNames[0]];
    const sheetB = workbookB.Sheets[workbookB.SheetNames[0]];
    const dataA = XLSX.utils.sheet_to_json(sheetA, { header: 1 });
    const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1 });

    // Validar datos
    if (dataA.length === 0 || dataB.length === 0) {
      throw new Error('Los archivos no pueden estar vacíos');
    }

    // Procesar y generar resultado
    const result = assignCodes(dataA, dataB);
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(result);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Resultado");

    const outputPath = path.join(__dirname, 'public', 'resultado.xlsx');
    XLSX.writeFile(newWorkbook, outputPath);

    res.download(outputPath, () => {
      // Limpiar archivos temporales
      fs.unlinkSync(req.files['fileA'][0].path);
      fs.unlinkSync(req.files['fileB'][0].path);
      fs.unlinkSync(outputPath);
    });

  } catch (error) {
    console.error('Error:', error);
    res.status(500).json({ error: error.message });
  }
});

const PORT = 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));
