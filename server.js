const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

const MAX_REPETICIONES = 8;
const PORT = process.env.PORT || 3000;

app.use(express.static('public'));
app.use(express.json());

function asignarCIs(totalRepeticiones, cis) {
    const resultado = [];
    let repeticionesAsignadas = 0;
    let ciIndex = 0;
    
    while (repeticionesAsignadas < totalRepeticiones) {
        const repeticiones = Math.min(MAX_REPETICIONES, totalRepeticiones - repeticionesAsignadas);
        const ci = cis[ciIndex % cis.length];
        
        for (let i = 0; i < repeticiones; i++) {
            resultado.push(ci);
        }
        
        repeticionesAsignadas += repeticiones;
        ciIndex++;
    }
    
    return resultado;
}

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    try {
        // Leer archivo de datos
        const datosWorkbook = XLSX.readFile(req.files['datosFile'][0].path);
        const datosSheet = datosWorkbook.Sheets[datosWorkbook.SheetNames[0]];
        const datos = XLSX.utils.sheet_to_json(datosSheet, { header: 1 });
        
        // Leer archivo de CIs
        const ciWorkbook = XLSX.readFile(req.files['ciFile'][0].path);
        const ciSheet = ciWorkbook.Sheets[ciWorkbook.SheetNames[0]];
        const cis = XLSX.utils.sheet_to_json(ciSheet, { header: 1 })
            .flat()
            .filter(ci => ci && ci.toString().trim() !== '')
            .map(ci => ci.toString().trim());
        
        if (cis.length === 0) {
            throw new Error('No hay CIs válidos en el archivo');
        }

        // Procesar datos
        const resultado = [];
        
        for (let i = 1; i < datos.length; i++) {
            const valor = datos[i][0]?.toString().trim() || 'Sin nombre';
            const total = parseInt(datos[i][1]) || 1;
            
            const asignaciones = asignarCIs(total, cis);
            
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
        res.status(500).json({ error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor listo en http://localhost:${PORT}`);
});
