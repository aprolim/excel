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
    
    console.log(`\nAsignando ${totalRepeticiones} repeticiones con ${cis.length} CIs disponibles`);
    console.log("Lista de CIs:", cis);
    
    while (repeticionesAsignadas < totalRepeticiones) {
        const repeticiones = Math.min(MAX_REPETICIONES, totalRepeticiones - repeticionesAsignadas);
        const ci = cis[ciIndex % cis.length];
        
        console.log(`- Asignando bloque de ${repeticiones} repeticiones al CI: ${ci}`);
        
        for (let i = 0; i < repeticiones; i++) {
            resultado.push(ci);
        }
        
        repeticionesAsignadas += repeticiones;
        ciIndex++;
    }
    
    // Calcular estadísticas de asignación
    const conteo = {};
    resultado.forEach(ci => {
        conteo[ci] = (conteo[ci] || 0) + 1;
    });
    
    console.log("Resumen de asignaciones:");
    Object.entries(conteo).forEach(([ci, count]) => {
        console.log(`  ${ci}: ${count} repeticiones`);
    });
    
    return resultado;
}

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    try {
        console.log("\n===== INICIO DE PROCESAMIENTO =====");
        
        // Leer archivo de datos
        const datosWorkbook = XLSX.readFile(req.files['datosFile'][0].path);
        const datosSheet = datosWorkbook.Sheets[datosWorkbook.SheetNames[0]];
        const datos = XLSX.utils.sheet_to_json(datosSheet, { header: 1 });
        
        console.log("\nDatos leídos del archivo principal:");
        console.table(datos.slice(0, 5)); // Mostrar primeras 5 filas

        // Leer archivo de CIs
        const ciWorkbook = XLSX.readFile(req.files['ciFile'][0].path);
        const ciSheet = ciWorkbook.Sheets[ciWorkbook.SheetNames[0]];
        let cis = XLSX.utils.sheet_to_json(ciSheet, { header: 1 })
            .flat()
            .filter(ci => ci && ci.toString().trim() !== '')
            .map(ci => ci.toString().trim());
        
        // Filtrar posibles encabezados
        if (isNaN(cis[0]) && cis.length > 1 && !isNaN(cis[1])) {
            console.warn("Advertencia: Se detectó posible encabezado en CIs. Eliminando primera fila");
            cis = cis.slice(1);
        }
        
        console.log(`\nCIs leídos (${cis.length}):`, cis);
        
        if (cis.length === 0) {
            throw new Error('No hay CIs válidos en el archivo');
        }

        // Procesar datos
        const resultado = [];
        
        for (let i = 1; i < datos.length; i++) {
            const fila = datos[i];
            if (!fila || fila.length < 2) continue;
            
            const valor = fila[0]?.toString().trim() || 'Sin nombre';
            const total = parseInt(fila[1]) || 1;
            
            console.log(`\nProcesando: ${valor} (${total} repeticiones)`);
            
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
        
        console.log("\nArchivo resultado generado:", outputPath);
        console.log("===== FIN DE PROCESAMIENTO =====\n");
        
        res.download(outputPath, () => {
            // Limpieza de archivos temporales
            fs.unlinkSync(req.files['datosFile'][0].path);
            fs.unlinkSync(req.files['ciFile'][0].path);
            fs.unlinkSync(outputPath);
        });

    } catch (error) {
        console.error("Error en procesamiento:", error);
        res.status(500).json({ error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor listo en http://localhost:${PORT}`);
});
