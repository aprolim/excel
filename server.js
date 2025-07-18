const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = 3000;

// Configurar multer para manejar archivos subidos
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Middleware para servir archivos estáticos
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

// Ruta principal - muestra el formulario HTML
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Ruta para procesar los archivos
app.post('/procesar', upload.fields([
    { name: 'archivoB', maxCount: 1 },
    { name: 'archivoA', maxCount: 1 }
]), (req, res) => {
    try {
        // Obtener los archivos subidos
        const archivoB = req.files['archivoB'][0];
        const archivoA = req.files['archivoA'][0];
        
        if (!archivoB || !archivoA) {
            return res.status(400).send('Debe subir ambos archivos');
        }
        
        // Procesar archivo B (códigos únicos)
        const wbB = XLSX.read(archivoB.buffer, { type: 'buffer' });
        const wsB = wbB.Sheets[wbB.SheetNames[0]];
        const dataB = XLSX.utils.sheet_to_json(wsB, { header: 1 });
        const codigos = dataB.flat().filter(Boolean);
        let codigoIndex = 0;
        
        // Procesar archivo A (valores repetidos)
        const wbA = XLSX.read(archivoA.buffer, { type: 'buffer' });
        const wsA = wbA.Sheets[wbA.SheetNames[0]];
        const dataA = XLSX.utils.sheet_to_json(wsA, { header: 1 });
        const valores = dataA.flat().filter(Boolean);
        
        // Contar ocurrencias y posiciones
        const grupos = {};
        valores.forEach((valor, idx) => {
            if (!grupos[valor]) grupos[valor] = { total: 0, posiciones: [] };
            grupos[valor].total++;
            grupos[valor].posiciones.push(idx);
        });
        
        // Preparar resultados (manteniendo orden original)
        const resultado = new Array(valores.length);
        
        // Procesar cada grupo
        Object.entries(grupos).forEach(([valor, data]) => {
            const total = data.total;
            const numGrupos = Math.ceil(total / 8);
            
            // Calcular distribución
            const tamBase = Math.floor(total / numGrupos);
            const extra = total % numGrupos;
            
            // Obtener códigos necesarios
            const cods = codigos.slice(codigoIndex, codigoIndex + numGrupos);
            codigoIndex += numGrupos;
            
            // Asignar a cada grupo
            let start = 0;
            for (let i = 0; i < numGrupos; i++) {
                const tamGrupo = tamBase + (i < extra ? 1 : 0);
                const end = start + tamGrupo;
                
                // Asignar código a todas las posiciones del grupo
                for (let j = start; j < end; j++) {
                    const pos = data.posiciones[j];
                    resultado[pos] = [valor, cods[i]];
                }
                start = end;
            }
        });
        
        // Crear libro de salida
        const wsResult = XLSX.utils.aoa_to_sheet(resultado);
        const wbResult = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbResult, wsResult, 'Resultado');
        
        // Generar buffer de salida
        const buffer = XLSX.write(wbResult, { type: 'buffer', bookType: 'xlsx' });
        
        // Enviar el archivo
        res.setHeader('Content-Disposition', 'attachment; filename="resultado.xlsx"');
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        
    } catch (error) {
        console.error(error);
        res.status(500).send('Error procesando los archivos: ' + error.message);
    }
});

// Iniciar servidor
app.listen(port, () => {
    console.log(`Servidor corriendo en http://localhost:${port}`);
});
