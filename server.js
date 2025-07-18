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

// Función para asignar CIs correctamente
function asignarCIs(totalRepeticiones, cis) {
    const resultado = [];
    let repeticionesAsignadas = 0;
    let ciIndex = 0;
    const totalCIs = cis.length;

    if (totalCIs === 0) return resultado;

    while (repeticionesAsignadas < totalRepeticiones) {
        const ciActual = cis[ciIndex % totalCIs];
        const repeticiones = Math.min(
            MAX_POR_GRUPO,
            totalRepeticiones - repeticionesAsignadas
        );

        // Agregar el CI las veces necesarias
        for (let i = 0; i < repeticiones; i++) {
            resultado.push(ciActual);
        }

        repeticionesAsignadas += repeticiones;
        ciIndex++;
    }

    return resultado;
}

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

        // Filtrar y limpiar CIs
        const cisFiltrados = cis.filter(ci => 
            ci !== null && ci !== undefined && ci.toString().trim() !== ''
        ).map(ci => ci.toString().trim());

        // Verificar que hay CIs
        if (cisFiltrados.length === 0) {
            return res.status(400).json({ error: 'El archivo de CIs está vacío o no tiene datos válidos' });
        }

        // Procesar datos
        const resultado = [];
        
        for (const row of datosJson.slice(1)) { // Saltar encabezado
            const valor = row[0]?.toString().trim() || 'Sin nombre';
            const totalRepeticiones = parseInt(row[1]) || 1;

            // Asignar CIs
            const asignaciones = asignarCIs(totalRepeticiones, cisFiltrados);

            // Agregar al resultado
            asignaciones.forEach(ci => {
                resultado.push([valor, ci]);
            });
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
        ['Dato Original', 'CI Asignado'],
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
