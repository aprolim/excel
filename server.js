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

app.post('/procesar', upload.fields([
    { name: 'datosFile', maxCount: 1 },
    { name: 'ciFile', maxCount: 1 }
]), async (req, res) => {
    console.log("\n====== NUEVA SOLICITUD RECIBIDA ======");
    
    try {
        // 1. Verificar archivos recibidos
        if (!req.files['datosFile'] || !req.files['ciFile']) {
            throw new Error('Faltan archivos requeridos');
        }

        // 2. Procesar archivo de datos
        const datosPath = req.files['datosFile'][0].path;
        const datosWorkbook = XLSX.readFile(datosPath);
        const datosSheetName = datosWorkbook.SheetNames[0];
        const datosSheet = datosWorkbook.Sheets[datosSheetName];
        const datos = XLSX.utils.sheet_to_json(datosSheet, { header: 1, blankrows: false });
        
        // Filtrar filas vacías
        const datosFiltrados = datos.filter(row => row.length > 0 && row[0] !== undefined);
        
        // 3. Procesar archivo de CIs
        const ciPath = req.files['ciFile'][0].path;
        const ciWorkbook = XLSX.readFile(ciPath);
        const ciSheetName = ciWorkbook.SheetNames[0];
        const ciSheet = ciWorkbook.Sheets[ciSheetName];
        const cisData = XLSX.utils.sheet_to_json(ciSheet, { header: 1, blankrows: false });
        
        // Extraer CIs válidos
        const cis = cisData
            .flat()
            .filter(ci => ci !== null && ci !== undefined && ci !== '')
            .map(ci => ci.toString().trim());
        
        if (cis.length === 0) {
            throw new Error('No se encontraron CIs válidos en el archivo');
        }

        // 4. Procesar cada registro
        const resultado = [['Dato Original', 'CI Asignado', 'Repeticiones Asignadas']];
        let filasProcesadas = 0;
        
        for (let i = 0; i < datosFiltrados.length; i++) {
            const row = datosFiltrados[i];
            
            // Saltar encabezados si existen
            if (i === 0 && isNaN(parseInt(row[1]))) continue;
            
            const valor = row[0]?.toString().trim() || `Dato ${i+1}`;
            const total = parseInt(row[1]);
            
            if (isNaN(total) || total <= 0) {
                console.warn(`Fila ${i+1} ignorada: repeticiones inválidas (${row[1]})`);
                continue;
            }
            
            const grupos = distribuirCIs(total, cis);
            
            grupos.forEach(grupo => {
                resultado.push([valor, grupo.ci, grupo.repeticiones]);
            });
            
            filasProcesadas++;
        }

        if (filasProcesadas === 0) {
            throw new Error('No se procesó ninguna fila válida');
        }

        // 5. Generar resultado
        const workbook = XLSX.utils.book_new();
        const worksheet = XLSX.utils.aoa_to_sheet(resultado);
        
        XLSX.utils.book_append_sheet(workbook, worksheet, "Grupos Asignados");
        
        const outputPath = path.join(__dirname, 'resultado.xlsx');
        XLSX.writeFile(workbook, outputPath);
        
        console.log("Archivo generado:", outputPath);

        // 6. Enviar respuesta
        res.download(outputPath, 'grupos_asignados.xlsx', (err) => {
            // Limpieza de archivos temporales
            [datosPath, ciPath, outputPath].forEach(file => {
                if (fs.existsSync(file)) fs.unlinkSync(file);
            });
        });

    } catch (error) {
        console.error("ERROR:", error.message);
        res.status(500).json({ error: error.message });
    }
});

app.listen(PORT, () => {
    console.log(`Servidor iniciado en http://localhost:${PORT}`);
});
