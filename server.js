const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

const MAX_POR_GRUPO = 8;
const PORT = process.env.PORT || 3000;

app.use(express.static('public'));
app.use(express.json());

// Función CORREGIDA para asignar CIs
function asignarCIsCorrectamente(totalRepeticiones, cis) {
    const resultado = [];
    let repeticionesAsignadas = 0;
    let ciIndex = 0;
    const totalCIs = cis.length;

    if (totalCIs === 0) return resultado;

    while (repeticionesAsignadas < totalRepeticiones) {
        const ciActual = cis[ciIndex % totalCIs]; // Rotación correcta de CIs
        const repeticiones = Math.min(
            MAX_POR_GRUPO,
            totalRepeticiones - repeticionesAsignadas
        );

        // Agregar el CI repetido
        for (let i = 0; i < repeticiones; i++) {
            resultado.push(ciActual);
        }

        repeticionesAsignadas += repeticiones;
        ciIndex++; // Avanzar al siguiente CI
    }

    return resultado;
}

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    try {
        if (!req.files['datosFile'] || !req.files['ciFile']) {
            return res.status(400).json({ error: 'Debe subir ambos archivos' });
        }

        // Leer archivos CORRECTAMENTE
        const datosWorkbook = XLSX.readFile(req.files['datosFile'][0].path);
        const datosSheet = datosWorkbook.Sheets[datosWorkbook.SheetNames[0]];
        const datosJson = XLSX.utils.sheet_to_json(datosSheet, { header: 1 });

        const ciWorkbook = XLSX.readFile(req.files['ciFile'][0].path);
        const ciSheet = ciWorkbook.Sheets[ciWorkbook.SheetNames[0]];
        const cis = XLSX.utils.sheet_to_json(ciSheet, { header: 1 })
            .flat()
            .filter(ci => ci !== null && ci !== undefined && ci.toString().trim() !== '')
            .map(ci => ci.toString().trim());

        if (cis.length === 0) {
            return res.status(400).json({ error: 'El archivo de CIs no contiene datos válidos' });
        }

        // Procesar datos
        const resultado = [];
        
        for (const row of datosJson.slice(1)) {
            const valor = row[0]?.toString().trim() || 'Sin nombre';
            const totalRepeticiones = parseInt(row[1]) || 1;

            const asignaciones = asignarCIsCorrectamente(totalRepeticiones, cis);

            asignaciones.forEach(ci => {
                resultado.push([valor, ci]);
            });
        }

        // Generar Excel
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet([
            ['Dato Original', 'CI Asignado'],
            ...resultado
        ]);
        XLSX.utils.book_append_sheet(workbook, worksheet, "Resultado");
        
        const outputPath = path.join(__dirname, 'resultado.xlsx');
        XLSX.writeFile(workbook, outputPath);

        res.download(outputPath, () => {
            fs.unlinkSync(req.files['datosFile'][0].path);
            fs.unlinkSync(req.files['ciFile'][0].path);
            fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
