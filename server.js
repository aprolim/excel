const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const upload = multer({ dest: 'uploads/' });

const MAX_GROUP_SIZE = 8;
const PORT = process.env.PORT || 3000;

app.use(express.static('public'));
app.use(express.json());

function distribuirCIs(totalRepeticiones, cis) {
    const grupos = [];
    let repAsignadas = 0;
    let ciIndex = 0;
    
    while (repAsignadas < totalRepeticiones) {
        const repsEnGrupo = Math.min(MAX_GROUP_SIZE, totalRepeticiones - repAsignadas);
        const ci = cis[ciIndex % cis.length];
        
        grupos.push({
            ci,
            repeticiones: repsEnGrupo
        });
        
        repAsignadas += repsEnGrupo;
        ciIndex++;
    }
    
    return grupos;
}

function obtenerPrimeraColumna(datos) {
    return datos
        .flatMap(row => {
            // Buscar el primer valor válido en la fila
            for (const cell of row) {
                if (cell !== null && cell !== undefined && cell !== '') {
                    return [cell];
                }
            }
            return [];
        })
        .map(ci => ci.toString().trim());
}

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    console.log("\n====== NUEVA SOLICITUD RECIBIDA ======");
    console.log("Procesando archivos...");
    
    try {
        // 1. Verificar archivos recibidos
        if (!req.files['datosFile'] || !req.files['ciFile']) {
            throw new Error('Faltan archivos requeridos');
        }

        // 2. Procesar archivo de datos
        const datosPath = req.files['datosFile'][0].path;
        console.log("Leyendo archivo de datos:", datosPath);
        
        const datosWorkbook = XLSX.readFile(datosPath);
        const datosSheetName = datosWorkbook.SheetNames[0];
        const datosSheet = datosWorkbook.Sheets[datosSheetName];
        
        // Leer como matriz de valores
        const datos = XLSX.utils.sheet_to_json(datosSheet, { header: 1, defval: null });
        console.log("Datos brutos leídos:", datos);
        
        // Filtrar filas vacías
        const datosFiltrados = datos.filter(row => 
            Array.isArray(row) && 
            row.some(cell => cell !== null && cell !== undefined && cell !== '')
        );
        
        console.log("Datos filtrados:", datosFiltrados);
        
        if (datosFiltrados.length === 0) {
            throw new Error('El archivo de datos está vacío o no se reconocieron filas válidas');
        }

        // 3. Procesar archivo de CIs
        const ciPath = req.files['ciFile'][0].path;
        console.log("Leyendo archivo de CIs:", ciPath);
        
        const ciWorkbook = XLSX.readFile(ciPath);
        const ciSheetName = ciWorkbook.SheetNames[0];
        const ciSheet = ciWorkbook.Sheets[ciSheetName];
        
        const cisBrutos = XLSX.utils.sheet_to_json(ciSheet, { header: 1, defval: null });
        console.log("CIs brutos:", cisBrutos);
        
        const cis = obtenerPrimeraColumna(cisBrutos);
        console.log("CIs válidos:", cis);
        
        if (cis.length === 0) {
            throw new Error('No se encontraron CIs válidos. Asegúrese de que haya valores en la primera columna');
        }

        // 4. Procesar cada registro
        const resultado = [
            ['Dato Original', 'CI Asignado', 'Repeticiones Asignadas']
        ];
        let filasProcesadas = 0;
        
        for (const [index, row] of datosFiltrados.entries()) {
            // Obtener valor y repeticiones de las primeras dos columnas
            const valor = row[0] !== null && row[0] !== undefined 
                ? row[0].toString().trim() 
                : `Dato ${index + 1}`;
            
            // Buscar el valor numérico en la fila (puede estar en cualquier columna)
            let repeticiones = null;
            for (let i = 1; i < row.length; i++) {
                const num = parseInt(row[i]);
                if (!isNaN(num)) {
                    repeticiones = num;
                    break;
                }
            }
            
            if (repeticiones === null) {
                console.warn(`Fila ${index + 1} ignorada: No se encontró número de repeticiones`);
                continue;
            }
            
            if (repeticiones <= 0) {
                console.warn(`Fila ${index + 1} ignorada: Repeticiones debe ser mayor a 0 (${repeticiones})`);
                continue;
            }
            
            console.log(`Procesando: ${valor} (${repeticiones} repeticiones)`);
            const grupos = distribuirCIs(repeticiones, cis);
            
            grupos.forEach(grupo => {
                resultado.push([valor, grupo.ci, grupo.repeticiones]);
            });
            
            filasProcesadas++;
        }

        if (filasProcesadas === 0) {
            throw new Error('No se procesaron filas válidas. Verifique que: \n- La primera columna contiene valores \n- Hay un número válido en la segunda columna (o en cualquier otra columna)');
        }

        // 5. Generar resultado
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(resultado);
        
        XLSX.utils.book_append_sheet(workbook, worksheet, "Grupos Asignados");
        
        const outputPath = path.join(__dirname, 'resultado.xlsx');
        XLSX.writeFile(workbook, outputPath);
        
        console.log("Archivo generado:", outputPath);
        console.log("Total registros:", resultado.length - 1);

        // 6. Enviar respuesta
        res.download(outputPath, 'grupos_asignados.xlsx', (err) => {
            // Limpieza de archivos temporales
            [datosPath, ciPath, outputPath].forEach(file => {
                if (fs.existsSync(file)) {
                    try {
                        fs.unlinkSync(file);
                        console.log("Archivo temporal eliminado:", file);
                    } catch (cleanError) {
                        console.error("Error eliminando temporal:", cleanError);
                    }
                }
            });
        });

    } catch (error) {
        console.error("ERROR:", error.message);
        res.status(500).json({ 
            error: error.message,
            recomendacion: "Verifique el formato de los archivos: \n- Datos: Primera columna=valores, Segunda (o cualquier columna)=números \n- CIs: Primera columna=valores"
        });
    }
});

app.listen(PORT, () => {
    console.log(`\n=================================`);
    console.log(`  Servidor iniciado en puerto ${PORT}`);
    console.log(`  Formato esperado para datos:`);
    console.log(`    Columna 1: Valor (texto)`);
    console.log(`    Columna 2: Repeticiones (número)`);
    console.log(`  Formato esperado para CIs:`);
    console.log(`    Columna 1: CI (texto o número)`);
    console.log(`=================================\n`);
});
