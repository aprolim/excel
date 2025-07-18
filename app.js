const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const app = express();
const port = process.env.PORT || 3000;

// Configuración de Multer
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Middleware
app.use(express.static('public'));
app.use(express.urlencoded({ extended: true }));

// Rutas
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.post('/procesar', upload.fields([
    { name: 'archivoB', maxCount: 1 },
    { name: 'archivoA', maxCount: 1 }
]), async (req, res) => {
    try {
        // Validación de archivos
        if (!req.files?.archivoB?.[0] || !req.files?.archivoA?.[0]) {
            return res.status(400).json({ error: 'Debe subir ambos archivos' });
        }

        // Procesar Archivo B
        const wbB = XLSX.read(req.files.archivoB[0].buffer, { type: 'buffer' });
        const codigos = XLSX.utils.sheet_to_json(wbB.Sheets[wbB.SheetNames[0]], { header: 1 })
                      .flat().filter(Boolean);

        // Procesar Archivo A
        const wbA = XLSX.read(req.files.archivoA[0].buffer, { type: 'buffer' });
        const valores = XLSX.utils.sheet_to_json(wbA.Sheets[wbA.SheetNames[0]], { header: 1 })
                      .flat().filter(Boolean);

        // Lógica de asignación
        const resultado = [];
        const grupos = {};
        let codigoIndex = 0;

        // Contar ocurrencias
        valores.forEach((valor, idx) => {
            if (!grupos[valor]) grupos[valor] = { indices: [] };
            grupos[valor].indices.push(idx);
        });

        // Asignar códigos
        Object.entries(grupos).forEach(([valor, data]) => {
            const total = data.indices.length;
            const numGrupos = Math.ceil(total / 8);
            const tamBase = Math.floor(total / numGrupos);
            const extra = total % numGrupos;

            // Verificar códigos disponibles
            if (codigoIndex + numGrupos > codigos.length) {
                throw new Error('No hay suficientes códigos únicos en el Archivo B');
            }

            const codigosGrupo = codigos.slice(codigoIndex, codigoIndex + numGrupos);
            codigoIndex += numGrupos;

            // Distribuir
            let start = 0;
            for (let i = 0; i < numGrupos; i++) {
                const tamGrupo = tamBase + (i < extra ? 1 : 0);
                const end = start + tamGrupo;

                for (let j = start; j < end; j++) {
                    resultado[data.indices[j]] = [valor, codigosGrupo[i]];
                }
                start = end;
            }
        });

        // Crear archivo de salida
        const ws = XLSX.utils.aoa_to_sheet(resultado);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Resultado');
        
        const buffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });

        // Enviar respuesta
        res.setHeader('Content-Disposition', 'attachment; filename="resultado.xlsx"');
        res.type('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);

    } catch (error) {
        console.error('Error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Iniciar servidor
app.listen(port, () => {
    console.log(`Servidor corriendo en http://localhost:${port}`);
});
