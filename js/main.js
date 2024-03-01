const btnAbrir = document.getElementById('btnAbrirExcel');

btnAbrir.addEventListener('click', abrir);


function abrir(e) {
    let archivo = e.target.files[0];
    let leer = new FileReader();

    leer.onload = function(event) {
        let dato = new Uint8Array(event.target.result);
        var libro = XLSX.read(dato, { type: 'array' });

        var firstSheetName = libro.SheetNames[0]; // Nombre de la primera hoja de
        var firstSheet = libro.Sheets[firstSheetName];

        console.log("Hoja de cálculo:", firstSheetName);

        // Llamar a una función de devolución de llamada para procesar los datos
        procesarDatos(firstSheet);
    }

    leer.readAsArrayBuffer(archivo);
}

function procesarDatos(firstSheet) {
    // Variable para verificar si la hoja de cálculo está vacía
    var isEmpty = true;

    Object.keys(firstSheet).forEach(function(sheetItem) {
        // Verificar si la celda es una celda de datos (no empieza con '!')
        if (sheetItem[0] !== '!') {
            isEmpty = false; // La hoja de cálculo no está vacía
            var cell = firstSheet[sheetItem];
            console.log("Celda:", sheetItem, "Valor:", cell.v);
        }
    });

    // Verificar si la hoja de cálculo está vacía
    if (isEmpty) {
        console.log("Advertencia: El archivo Excel está vacío.");
    } else {
        console.log("La hoja de cálculo se ha leído completamente.");
    }

        // Forzar la visualización en la consola después de que se complete la iteración del bucle
        setTimeout(function() {
            console.log("Proceso de lectura completado.");
        }, 0);
}