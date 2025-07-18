const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;

// Configurar multer
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Middleware
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

// Ruta principal
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// Ruta para procesar archivos
app.post('/procesar', upload.fields([
    { name: 'archivoB', maxCount: 1 },
    { name: 'archivoA', maxCount: 1 }
]), (req, res) => {
    try {
        // Validar archivos subidos
        if (!req.files || !req.files['archivoB'] || !req.files['archivoA']) {
            return res.status(400).send('Debe subir ambos archivos');
        }
        
        const archivoB = req.files['archivoB'][0];
        const archivoA = req.files['archivoA'][0];
        
        // Leer archivo B
        const wbB = XLSX.read(archivoB.buffer, { type: 'buffer' });
        const wsB = wbB.Sheets[wbB.SheetNames[0]];
        const dataB = XLSX.utils.sheet_to_json(wsB, { header: 1 });
        const codigos = dataB.flat().filter(Boolean);
        let codigoIndex = 0;
        
        // Leer archivo A
        const wbA = XLSX.read(archivoA.buffer, { type: 'buffer' });
        const wsA = wbA.Sheets[wbA.SheetNames[0]];
        const dataA = XLSX.utils.sheet_to_json(wsA, { header: 1 });
        const valores = dataA.flat().filter(Boolean);
        
        // Contar ocurrencias
        const grupos = {};
        valores.forEach((valor, idx) => {
            if (!grupos[valor]) grupos[valor] = { total: 0, posiciones: [] };
            grupos[valor].total++;
            grupos[valor].posiciones.push(idx);
        });
        
        // Preparar resultados
        const resultado = new Array(valores.length);
        
        // Procesar cada grupo
        Object.entries(grupos).forEach(([valor, data]) => {
            const total = data.total;
            const numGrupos = Math.ceil(total / 8);
            
            // Distribución balanceada
            const tamBase = Math.floor(total / numGrupos);
            const extra = total % numGrupos;
            
            // Obtener códigos
            const cods = codigos.slice(codigoIndex, codigoIndex + numGrupos);
            codigoIndex += numGrupos;
            
            // Asignar códigos
            let start = 0;
            for (let i = 0; i < numGrupos; i++) {
                const tamGrupo = tamBase + (i < extra ? 1 : 0);
                const end = start + tamGrupo;
                
                for (let j = start; j < end; j++) {
                    const pos = data.posiciones[j];
                    resultado[pos] = [valor, cods[i]];
                }
                start = end;
            }
        });
        
        // Crear archivo de salida
        const wsResult = XLSX.utils.aoa_to_sheet(resultado);
        const wbResult = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wbResult, wsResult, 'Resultado');
        const buffer = XLSX.write(wbResult, { type: 'buffer', bookType: 'xlsx' });
        
        // Enviar respuesta
        res.setHeader('Content-Disposition', 'attachment; filename="resultado.xlsx"');
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
        
    } catch (error) {
        console.error('Error:', error);
        res.status(500).send('Error procesando los archivos: ' + error.message);
    }
});

// Iniciar servidor
app.listen(port, () => {
    console.log(`Servidor corriendo en http://localhost:${port}`);
});
