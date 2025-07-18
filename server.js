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
    
    console.log(`[AsignarCIs] Iniciando asignación para ${totalRepeticiones} repeticiones`);
    console.log(`[AsignarCIs] CIs disponibles: ${cis.join(', ')}`);
    
    while (repeticionesAsignadas < totalRepeticiones) {
        const repeticiones = Math.min(MAX_REPETICIONES, totalRepeticiones - repeticionesAsignadas);
        const ci = cis[ciIndex % cis.length];
        
        console.log(`[AsignarCIs] Asignando ${repeticiones} repeticiones a ${ci}`);
        
        for (let i = 0; i < repeticiones; i++) {
            resultado.push(ci);
        }
        
        repeticionesAsignadas += repeticiones;
        ciIndex++;
    }
    
    console.log(`[AsignarCIs] Asignación completada. Total: ${resultado.length} registros`);
    return resultado;
}

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    console.log("\n====== NUEVA SOLICITUD RECIBIDA ======");
    
    try {
        // 1. Verificar archivos recibidos
        console.log("[Procesar] Archivos recibidos:", req.files);
        if (!req.files['datosFile'] || !req.files['ciFile']) {
            throw new Error('Faltan archivos requeridos');
        }

        // 2. Procesar archivo de datos
        const datosPath = req.files['datosFile'][0].path;
        console.log("[Procesar] Leyendo archivo de datos:", datosPath);
        
        const datosWorkbook = XLSX.readFile(datosPath);
        console.log("[Procesar] Hojas disponibles en datosFile:", datosWorkbook.SheetNames);
        
        const datosSheet = datosWorkbook.Sheets[datosWorkbook.SheetNames[0]];
        const datos = XLSX.utils.sheet_to_json(datosSheet, { header: 1 });
        
        console.log("[Procesar] Datos brutos leídos (primeras 5 filas):", datos.slice(0, 5));
        console.log("[Procesar] Total de filas en datos:", datos.length);

        // 3. Procesar archivo de CIs
        const ciPath = req.files['ciFile'][0].path;
        console.log("[Procesar] Leyendo archivo de CIs:", ciPath);
        
        const ciWorkbook = XLSX.readFile(ciPath);
        console.log("[Procesar] Hojas disponibles en ciFile:", ciWorkbook.SheetNames);
        
        const ciSheet = ciWorkbook.Sheets[ciWorkbook.SheetNames[0]];
        const cisBrutos = XLSX.utils.sheet_to_json(ciSheet, { header: 1 });
        
        console.log("[Procesar] CIs brutos leídos:", cisBrutos);
        
        // Procesamiento robusto de CIs
        const cis = cisBrutos
            .flat()
            .filter(ci => ci !== null && ci !== undefined)
            .map(ci => ci.toString().trim())
            .filter(ci => ci !== '');

        console.log("[Procesar] CIs válidos encontrados:", cis);
        
        if (cis.length === 0) {
            throw new Error('No se encontraron CIs válidos en el archivo');
        }

        // 4. Procesar cada registro
        const resultado = [];
        let filasProcesadas = 0;
        
        for (let i = 0; i < datos.length; i++) {
            // Saltar filas vacías o encabezados
            if (i === 0 || !datos[i] || datos[i].length < 2) {
                console.log(`[Procesar] Saltando fila ${i}:`, datos[i]);
                continue;
            }
            
            const valor = datos[i][0]?.toString().trim() || `Sin nombre (fila ${i+1})`;
            const total = parseInt(datos[i][1]);
            
            if (isNaN(total) || total <= 0) {
                console.warn(`[Procesar] Valor inválido en fila ${i}:`, datos[i]);
                continue;
            }
            
            console.log(`\n[Procesar] Procesando fila ${i}: ${valor} (${total} repeticiones)`);
            const asignaciones = asignarCIs(total, cis);
            
            asignaciones.forEach(ci => {
                resultado.push([valor, ci]);
            });
            
            filasProcesadas++;
        }

        if (filasProcesadas === 0) {
            throw new Error('No se procesó ninguna fila válida. Verifique el formato del archivo');
        }

        // 5. Generar resultado
        console.log("[Procesar] Generando archivo de resultado...");
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet([
            ['Dato Original', 'CI Asignado'],
            ...resultado
        ]);
        
        XLSX.utils.book_append_sheet(workbook, worksheet, "Resultado");
        
        const outputPath = path.join(__dirname, 'resultado.xlsx');
        XLSX.writeFile(workbook, outputPath);
        
        console.log("[Procesar] Archivo generado:", outputPath);
        console.log("[Procesar] Total de registros generados:", resultado.length);
        console.log("====== PROCESAMIENTO COMPLETADO ======\n");

        // 6. Enviar respuesta
        res.download(outputPath, 'asignaciones.xlsx', (err) => {
            // Limpieza de archivos temporales
            fs.unlinkSync(datosPath);
            fs.unlinkSync(ciPath);
            fs.unlinkSync(outputPath);
            
            if (err) {
                console.error("[Procesar] Error al descargar:", err);
            } else {
                console.log("[Procesar] Archivos temporales eliminados");
            }
        });

    } catch (error) {
        console.error("[Procesar] ERROR:", error.message);
        res.status(500).json({ 
            error: error.message,
            detalle: "Verifique los archivos y consulte los logs del servidor"
        });
    }
});

app.listen(PORT, () => {
    console.log(`\nServidor iniciado en http://localhost:${PORT}`);
    console.log("Esperando archivos...");
});
