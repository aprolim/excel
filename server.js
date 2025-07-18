const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

// Configuración
const MAX_POR_GRUPO = 8;
const PORT = process.env.PORT || 3000;

// Middleware
app.use(express.static('public'));
app.use(express.json());

// Función para distribución equilibrada
function distribuirEquilibrado(total, maxPorGrupo = MAX_POR_GRUPO) {
    if (total <= maxPorGrupo) return [total];
    
    const numGrupos = Math.ceil(total / maxPorGrupo);
    const base = Math.floor(total / numGrupos);
    let resto = total % numGrupos;
    
    const grupos = Array(numGrupos).fill(base);
    
    // Distribuir el resto
    for (let i = 0; i < resto; i++) {
        grupos[i]++;
    }
    
    // Ajustar para no exceder el máximo
    return grupos.map(g => Math.min(g, maxPorGrupo));
}

// Ruta principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Ruta para procesar
app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    try {
        // Validar archivos
        if (!req.files['datosFile'] || !req.files['ciFile']) {
            return res.status(400).json({ error: 'Debe subir ambos archivos' });
        }

        // Leer archivos
        const [datosJson, cis] = await Promise.all([
            leerExcel(req.files['datosFile'][0].path, 0),
            leerExcel(req.files['ciFile'][0].path, 0, true)
        ]);

        // Procesar datos
        const resultado = [];
        let ciIndex = 0;
        const ciCount = cis.length;

        for (const row of datosJson.slice(1)) { // Saltar encabezado
            const valor = row[0];
            const total = parseInt(row[1]) || 1;

            const grupos = distribuirEquilibrado(total);

            for (const cantidad of grupos) {
                const ciGrupo = [];
                for (let i = 0; i < cantidad; i++) {
                    ciGrupo.push(cis[ciIndex % ciCount]);
                    ciIndex++;
                }
                resultado.push([valor, cantidad, ciGrupo.join(', ')]);
            }
        }

        // Crear y enviar Excel
        const outputPath = await generarExcel(resultado);
        res.download(outputPath, () => {
            // Limpiar archivos temporales
            fs.unlinkSync(req.files['datosFile'][0].path);
            fs.unlinkSync(req.files['ciFile'][0].path);
            fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Funciones auxiliares
async function leerExcel(filePath, sheetIndex, flatten = false) {
    const workbook = XLSX.readFile(filePath);
    const sheet = workbook.Sheets[workbook.SheetNames[sheetIndex]];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });
    return flatten ? json.flat() : json;
}

async function generarExcel(data) {
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.aoa_to_sheet([
        ['Dato Original', 'Cantidad', 'CIs Asignados'],
        ...data
    ]);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Resultado");
    
    const outputPath = path.join(__dirname, 'resultado.xlsx');
    XLSX.writeFile(workbook, outputPath);
    return outputPath;
}

// Iniciar servidor
app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
