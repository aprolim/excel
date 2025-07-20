const express = require('express');
const multer = require('multer');
const cors = require('cors');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
app.use(cors());
app.use(express.static('public'));

// Configuración de Multer
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    if (!fs.existsSync('uploads')) {
      fs.mkdirSync('uploads');
    }
    cb(null, 'uploads/');
  },
  filename: (req, file, cb) => cb(null, Date.now() + '-' + file.originalname)
});
const upload = multer({ storage });

// Función optimizada para asignación perfecta de códigos
function assignCodesPerfectly(dataA, dataB) {
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
  Object.entries(frequencyMap).forEach(([value, count]) => {
    const codesNeeded = Math.ceil(count / 8);
    const availableCodes = codesB.length - codeIndex;
    
    if (availableCodes < codesNeeded) {
      throw new Error(`No hay suficientes códigos para el valor '${value}'. Se necesitan ${codesNeeded} códigos pero solo hay ${availableCodes} disponibles.`);
    }

    const codesToUse = codesB.slice(codeIndex, codeIndex + codesNeeded);
    codeIndex += codesNeeded;

    // Distribución perfectamente equitativa
    const baseCount = Math.floor(count / codesToUse.length);
    const extra = count % codesToUse.length;

    codesToUse.forEach((code, i) => {
      const times = i < extra ? baseCount + 1 : baseCount;
      for (let j = 0; j < times; j++) {
        result.push({ Valor: value, Codigo: code });
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
    if (!req.files['fileA'] || !req.files['fileB']) {
      throw new Error('Debes subir ambos archivos');
    }

    // Leer archivos
    const workbookA = XLSX.readFile(req.files['fileA'][0].path);
    const workbookB = XLSX.readFile(req.files['fileB'][0].path);

    // Obtener datos de la primera hoja
    const getFirstSheetData = (workbook) => {
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      return XLSX.utils.sheet_to_json(sheet, { header: 1 });
    };

    const dataA = getFirstSheetData(workbookA);
    const dataB = getFirstSheetData(workbookB);

    // Validar datos
    if (dataA.length === 0 || dataB.length === 0) {
      throw new Error('Los archivos no pueden estar vacíos');
    }

    // Procesar y generar resultado
    const result = assignCodesPerfectly(dataA, dataB);
    const newWorkbook = XLSX.utils.book_new();
    const newSheet = XLSX.utils.json_to_sheet(result);
    XLSX.utils.book_append_sheet(newWorkbook, newSheet, "Resultado");

    const outputDir = path.join(__dirname, 'public');
    if (!fs.existsSync(outputDir)) {
      fs.mkdirSync(outputDir);
    }

    const outputPath = path.join(outputDir, 'resultado.xlsx');
    XLSX.writeFile(newWorkbook, outputPath);

    // Configurar respuesta
    res.download(outputPath, 'resultado_codigos.xlsx', (err) => {
      // Limpiar archivos temporales
      [req.files['fileA'][0].path, req.files['fileB'][0].path, outputPath].forEach(file => {
        if (fs.existsSync(file)) {
          fs.unlinkSync(file);
        }
      });
      
      if (err) {
        console.error('Error al descargar:', err);
      }
    });

  } catch (error) {
    console.error('Error en el servidor:', error);
    res.status(500).json({ 
      success: false,
      message: error.message
    });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Servidor corriendo en http://localhost:${PORT}`));
