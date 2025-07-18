document.addEventListener('DOMContentLoaded', () => {
    const form = document.getElementById('uploadForm');
    const resultContainer = document.getElementById('resultContainer');
    const btnSubmit = form.querySelector('button[type="submit"]');
    const btnText = btnSubmit.querySelector('.btn-text');
    const spinner = btnSubmit.querySelector('.spinner');

    form.addEventListener('submit', async (e) => {
        e.preventDefault();
        
        // Validar archivos
        const datosFile = document.getElementById('datosFile').files[0];
        const ciFile = document.getElementById('ciFile').files[0];
        
        if (!datosFile || !ciFile) {
            showAlert('Debe seleccionar ambos archivos', 'error');
            return;
        }

        // Mostrar loading
        btnText.textContent = 'Procesando...';
        spinner.classList.remove('hidden');
        btnSubmit.disabled = true;

        try {
            const formData = new FormData();
            formData.append('datosFile', datosFile);
            formData.append('ciFile', ciFile);

            const response = await fetch('/procesar', {
                method: 'POST',
                body: formData
            });

            if (response.ok) {
                // Mostrar mensaje de éxito
                resultContainer.classList.remove('hidden');
                form.reset();
                
                // Descargar archivo
                const blob = await response.blob();
                const url = window.URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = 'resultado_equilibrado.xlsx';
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                a.remove();
            } else {
                const error = await response.json();
                throw new Error(error.error || 'Error al procesar los archivos');
            }
        } catch (error) {
            showAlert(error.message, 'error');
            console.error('Error:', error);
        } finally {
            // Restaurar botón
            btnText.textContent = 'Procesar Archivos';
            spinner.classList.add('hidden');
            btnSubmit.disabled = false;
        }
    });

    function showAlert(message, type = 'success') {
        const alert = document.createElement('div');
        alert.className = `alert alert-${type}`;
        alert.textContent = message;
        
        // Estilos para el alert (podrías mover esto a CSS)
        alert.style.position = 'fixed';
        alert.style.top = '20px';
        alert.style.right = '20px';
        alert.style.padding = '15px';
        alert.style.borderRadius = '5px';
        alert.style.color = 'white';
        alert.style.backgroundColor = type === 'error' ? 'var(--error-color)' : 'var(--success-color)';
        alert.style.boxShadow = '0 2px 10px rgba(0,0,0,0.2)';
        alert.style.zIndex = '1000';
        alert.style.animation = 'fadeIn 0.3s';
        
        document.body.appendChild(alert);
        
        setTimeout(() => {
            alert.style.animation = 'fadeOut 0.3s';
            setTimeout(() => alert.remove(), 300);
        }, 3000);
    }
});
