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

// Función para distribución perfectamente equilibrada
function distribuirCIs(totalRepeticiones, cis) {
    const resultado = [];
    const totalCIs = cis.length;
    
    if (totalCIs === 0) return resultado;
    
    // Calcular repeticiones por CI de forma equilibrada
    let repeticionesPorCI = Math.floor(totalRepeticiones / totalCIs);
    let repeticionesExtra = totalRepeticiones % totalCIs;
    
    // Ajustar para no exceder el máximo por grupo
    while (repeticionesPorCI + (repeticionesExtra > 0 ? 1 : 0) > MAX_POR_GRUPO) {
        repeticionesPorCI = MAX_POR_GRUPO;
        repeticionesExtra = totalRepeticiones - (repeticionesPorCI * totalCIs);
        
        if (repeticionesExtra < 0) {
            repeticionesPorCI = Math.floor(totalRepeticiones / totalCIs);
            repeticionesExtra = totalRepeticiones % totalCIs;
            break;
        }
    }

    // Asignar repeticiones
    for (let i = 0; i < totalCIs; i++) {
        let repeticiones = repeticionesPorCI;
        if (i < repeticionesExtra) {
            repeticiones++;
        }

        // Dividir en grupos de máximo 8
        while (repeticiones > 0) {
            const asignar = Math.min(repeticiones, MAX_POR_GRUPO);
            resultado.push({
                ci: cis[i],
                veces: asignar
            });
            repeticiones -= asignar;
        }
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

        // Filtrar CIs válidos
        const cisFiltrados = cis.filter(ci => ci && ci.toString().trim() !== '');

        // Procesar datos
        const resultado = [];
        
        for (const row of datosJson.slice(1)) { // Saltar encabezado
            const valor = row[0];
            const totalRepeticiones = parseInt(row[1]) || 1;

            // Distribuir CIs
            const asignaciones = distribuirCIs(totalRepeticiones, cisFiltrados);

            // Agregar al resultado (sin columna de repeticiones)
            asignaciones.forEach(asig => {
                for (let i = 0; i < asig.veces; i++) {
                    resultado.push([valor, asig.ci]);
                }
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
