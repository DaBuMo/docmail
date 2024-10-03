export function mostrarSpinner(idSpinner) {
    document.getElementById(idSpinner).style.visibility = 'visible';
}

export function ocultarSpinner(idSpinner) {
    document.getElementById(idSpinner).style.visibility = 'hidden';
}

export function configurarEventos() {
    document.getElementById('enviarBtn').addEventListener('click', function() {
        mostrarSpinner('enviarSpinner');
        setTimeout(() => ocultarSpinner('enviarSpinner'), 2000); // Ocultar el spinner después de 2 segundos (ejemplo)
    });

    document.getElementById('imprimirBtn').addEventListener('click', function() {
        mostrarSpinner('imprimirSpinner');
        setTimeout(() => ocultarSpinner('imprimirSpinner'), 2000); // Ocultar el spinner después de 2 segundos (ejemplo)
    });
}