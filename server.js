const express = require('express');
const multer = require('multer');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');

const app = express();
app.use(cors());
app.use(express.static('public'));

// Configura Multer para subir archivos
const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, 'uploads/'),
  filename: (req, file, cb) => cb(null, file.originalname)
});
const upload = multer({ storage });

// Ruta para subir archivos y procesarlos
app.post('/process', upload.fields([
  { name: 'fileA', maxCount: 1 },
  { name: 'fileB', maxCount: 1 }
]), (req, res) => {
  try {
    // Leer archivos subidos
    const workbookA = XLSX.readFile(req.files['fileA'][0].path);
    const workbookB = XLSX.readFile(req.files['fileB'][0].path);

    // Obtener datos de la primera hoja
    const sheetA = workbookA.Sheets[workbookA.SheetNames[0]];
    const sheetB = workbookB.Sheets[workbookB.SheetNames[0]];

    // Convertir a JSON
    const dataA = XLSX.utils.sheet_to_json(sheetA, { header: 1 });
    const dataB = XLSX.utils.sheet_to_json(sheetB, { header: 1 });

    // Procesamiento de datos
    const result = assignCodes(dataA, dataB);

    // Crear nuevo workbook y guardar
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(result);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Resultado");

    const outputPath = path.join(__dirname, 'public', 'resultado.xlsx');
    XLSX.writeFile(newWorkbook, outputPath);

    res.download(outputPath);
  } catch (error) {
    res.status(500).send('Error al procesar archivos: ' + error.message);
  }
});

// Función para asignar códigos
function assignCodes(dataA, dataB) {
  const valuesA = dataA.map(row => row[0]).filter(Boolean);
  const codesB = dataB.map(row => row[0]).filter(Boolean);

  const frequencyMap = {};
  valuesA.forEach(value => {
    frequencyMap[value] = (frequencyMap[value] || 0) + 1;
  });

  const result = [];
  let codeIndex = 0;

  Object.keys(frequencyMap).forEach(key => {
    const count = frequencyMap[key];
    const codesNeeded = Math.ceil(count / 8);
    const codes = codesB.slice(codeIndex, codeIndex + codesNeeded);
    codeIndex += codesNeeded;

    if (codes.length === 0) return;

    const baseAssignments = Math.floor(count / codes.length);
    const remainder = count % codes.length;

    codes.forEach((code, i) => {
      const assignments = baseAssignments + (i < remainder ? 1 : 0);
      for (let j = 0; j < assignments; j++) {
        result.push({ Valor: key, Codigo: code });
      }
    });
  });

  return result;
}

const PORT = 3000;
app.listen(PORT, () => console.log(`Servidor en http://localhost:${PORT}`));
