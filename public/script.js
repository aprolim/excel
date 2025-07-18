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
        resultContainer.classList.add('hidden');  // Ocultar resultados previos

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
                a.download = 'grupos_asignados.xlsx';  // Nombre actualizado
                document.body.appendChild(a);
                a.click();
                window.URL.revokeObjectURL(url);
                a.remove();
            } else {
                // Manejar errores específicos del servidor
                try {
                    const errorData = await response.json();
                    throw new Error(errorData.error || 'Error al procesar los archivos');
                } catch (jsonError) {
                    throw new Error('Error en el servidor: ' + response.statusText);
                }
            }
        } catch (error) {
            showAlert(error.message, 'error');
            console.error('Error:', error);
        } finally {
            // Restaurar botón
            btnText.textContent = 'Generar Grupos';
            spinner.classList.add('hidden');
            btnSubmit.disabled = false;
        }
    });

    function showAlert(message, type = 'success') {
        // Eliminar alertas anteriores
        const existingAlerts = document.querySelectorAll('.alert');
        existingAlerts.forEach(alert => alert.remove());
        
        const alert = document.createElement('div');
        alert.className = `alert ${type === 'error' ? 'alert-error' : 'alert-success'}`;
        alert.textContent = message;
        
        // Estilos
        alert.style.position = 'fixed';
        alert.style.top = '20px';
        alert.style.right = '20px';
        alert.style.padding = '15px 20px';
        alert.style.borderRadius = '5px';
        alert.style.color = 'white';
        alert.style.backgroundColor = type === 'error' ? 'var(--error-color)' : 'var(--success-color)';
        alert.style.boxShadow = '0 2px 15px rgba(0,0,0,0.2)';
        alert.style.zIndex = '1000';
        alert.style.fontSize = '16px';
        alert.style.display = 'flex';
        alert.style.alignItems = 'center';
        alert.style.animation = 'fadeIn 0.3s ease-out';
        
        // Icono
        const icon = document.createElement('span');
        icon.textContent = type === 'error' ? '⚠️ ' : '✅ ';
        icon.style.marginRight = '10px';
        icon.style.fontSize = '20px';
        alert.prepend(icon);
        
        document.body.appendChild(alert);
        
        // Animación de salida
        setTimeout(() => {
            alert.style.animation = 'fadeOut 0.3s ease-in forwards';
            setTimeout(() => alert.remove(), 300);
        }, 4000);
    }

    // Definir animaciones CSS dinámicamente
    const style = document.createElement('style');
    style.textContent = `
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(-20px); }
            to { opacity: 1; transform: translateY(0); }
        }
        @keyframes fadeOut {
            from { opacity: 1; transform: translateY(0); }
            to { opacity: 0; transform: translateY(-20px); }
        }
    `;
    document.head.appendChild(style);
});
