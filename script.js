const VERSION = "1.1.0";
let listaA = [];
let listaB = [];
let resultadosGlobales = [];
let filtroActual = "todos";

// =====================================
// LEER EXCEL
// =====================================

function leerExcel(archivo) {

    return new Promise((resolve, reject) => {

        const reader = new FileReader();

        reader.onload = function(e) {

            const data = e.target.result;

            const workbook = XLSX.read(data, {
                type: 'binary'
            });

            const hoja =
                workbook.Sheets[
                    workbook.SheetNames[0]
                ];

            const datos =
                XLSX.utils.sheet_to_json(hoja);

            resolve(datos);
        };

        reader.onerror = reject;

        reader.readAsBinaryString(archivo);

    });

}

function validarColumnas(datos) {

    if (!datos.length) {
        return false;
    }

    const columnas =
        Object.keys(datos[0])
            .map(c => c.toUpperCase());

    return (
        columnas.includes("ORIGEN") &&
        columnas.includes("DESTINO") &&
        columnas.includes("TIPO") &&
        columnas.includes("COLOR") &&
        columnas.includes("SECCION")
    );

}

function convertirFormato(datos) {

    return datos.map(fila => ({

        origen:
            fila.ORIGEN ??
            fila.Origen ??
            fila.origen,

        destino:
            fila.DESTINO ??
            fila.Destino ??
            fila.destino,

        tipo:
            fila.TIPO ??
            fila.Tipo ??
            fila.tipo,

        color:
            fila.COLOR ??
            fila.Color ??
            fila.color,

        seccion:
            fila.SECCION ??
            fila.Seccion ??
            fila.seccion

    }));

}

// =====================================
// EXPORTAR EXCEL
// =====================================

async function exportarExcel() {

    if (resultadosGlobales.length === 0) {

        alert("No hay resultados para exportar");

        return;
    }

    const workbook = new ExcelJS.Workbook();

    const worksheet =
        workbook.addWorksheet("Comparativa");

    worksheet.columns = [

        { header: "ESTADO", width: 18 },
        { header: "ORIGEN", width: 25 },
        { header: "DESTINO", width: 25 },
        { header: "TIPO", width: 20 },
        { header: "COLOR", width: 25 },
        { header: "SECCION", width: 15 }

    ];

    resultadosGlobales.forEach(r => {

        let fila;

        if (r.estado === "igual") {

            fila = worksheet.addRow([
                "SIN CAMBIOS",
                r.nuevo.origen,
                r.nuevo.destino,
                r.nuevo.tipo,
                r.nuevo.color,
                r.nuevo.seccion
            ]);

            colorearFila(
                fila,
                "FFD4EDDA"
            );

        }

        else if (r.estado === "nuevo") {

            fila = worksheet.addRow([
                "NUEVO",
                r.cable.origen,
                r.cable.destino,
                r.cable.tipo,
                r.cable.color,
                r.cable.seccion
            ]);

            colorearFila(
                fila,
                "FFCFE2FF"
            );

        }

        else if (r.estado === "eliminado") {

            fila = worksheet.addRow([
                "ELIMINADO",
                r.cable.origen,
                r.cable.destino,
                r.cable.tipo,
                r.cable.color,
                r.cable.seccion
            ]);

            colorearFila(
                fila,
                "FFF8D7DA"
            );

        }

        else if (r.estado === "modificado") {

            fila = worksheet.addRow([

                "MODIFICADO",

                r.nuevo.origen,

                r.nuevo.destino,

                generarTextoCambio(
                    r.antiguo.tipo,
                    r.nuevo.tipo
                ),

                generarTextoCambio(
                    r.antiguo.color,
                    r.nuevo.color
                ),

                generarTextoCambio(
                    r.antiguo.seccion,
                    r.nuevo.seccion
                )

            ]);

            colorearFila(
                fila,
                "FFFFF3CD"
            );

            if (r.antiguo.tipo !== r.nuevo.tipo)
                colorearCelda(
                    fila.getCell(4)
                );

            if (r.antiguo.color !== r.nuevo.color)
                colorearCelda(
                    fila.getCell(5)
                );

            if (r.antiguo.seccion !== r.nuevo.seccion)
                colorearCelda(
                    fila.getCell(6)
                );
        }

    });

    worksheet.autoFilter = {
        from: 'A1',
        to: 'F1'
    };

    const buffer =
        await workbook.xlsx.writeBuffer();

    const blob = new Blob(
        [buffer],
        {
            type:
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        }
    );

    const url =
        window.URL.createObjectURL(blob);

    const a =
        document.createElement("a");

    a.href = url;

const ahora = new Date();

const fechaHora =
    ahora.getFullYear() +
    "-" +
    String(ahora.getMonth() + 1).padStart(2, "0") +
    "-" +
    String(ahora.getDate()).padStart(2, "0") +
    "_" +
    String(ahora.getHours()).padStart(2, "0") +
    "-" +
    String(ahora.getMinutes()).padStart(2, "0") +
    "-" +
    String(ahora.getSeconds()).padStart(2, "0");

a.download =
    `Comparativa_Cableado_${fechaHora}.xlsx`;

    a.click();

    window.URL.revokeObjectURL(url);

}

function colorearFila(fila, color) {

    fila.eachCell(cell => {

        cell.fill = {

            type: 'pattern',

            pattern: 'solid',

            fgColor: {
                argb: color
            }

        };

    });

}

function colorearCelda(cell) {

    cell.fill = {

        type: 'pattern',

        pattern: 'solid',

        fgColor: {
            argb: "FFFFD54F"
        }

    };

    cell.font = {
        bold: true
    };

}

function generarTextoCambio(
    antiguo,
    nuevo
) {

    if (antiguo === nuevo) {
        return nuevo;
    }

    return `${antiguo} → ${nuevo}`;

}

// =====================================
// NORMALIZACION
// =====================================

function normalizarTexto(valor) {

    if (valor === null || valor === undefined) {
        return "";
    }

    return String(valor)

        // Quitar espacios extremos
        .trim()

        // Convertir a mayúsculas
        .toUpperCase()

        // Reducir espacios múltiples
        .replace(/\s+/g, " ")

        // Quitar espacios alrededor de :
        .replace(/\s*:\s*/g, ":")

        // Normalizar decimales
        .replace(/,/g, ".");
}

function normalizarCamposCable(cable) {

    return {

        origen: normalizarTexto(cable.origen),

        destino: normalizarTexto(cable.destino),

        tipo: normalizarTexto(cable.tipo),

        color: normalizarTexto(cable.color),

        seccion: normalizarTexto(cable.seccion)

    };

}

function normalizarCable(cable) {

    cable = normalizarCamposCable(cable);

    const puntos = [
        cable.origen,
        cable.destino
    ].sort((a, b) => a.localeCompare(b));

    return {
        ...cable,
        origen: puntos[0],
        destino: puntos[1]
    };
}

// =====================================
// PREPARAR LISTA
// =====================================

function prepararLista(lista) {

    const normalizada = lista.map(normalizarCable);

    // Agrupar por origen+destino
    const grupos = {};

    normalizada.forEach(cable => {

        const claveBase =
            `${cable.origen}|${cable.destino}`;

        if (!grupos[claveBase]) {
            grupos[claveBase] = [];
        }

        grupos[claveBase].push(cable);

    });

    const resultado = [];

    Object.values(grupos).forEach(grupo => {

        // Orden estable para que el índice sea consistente
        grupo.sort((a, b) => {

            const claveA =
                `${a.tipo}|${a.color}|${a.seccion}`;

            const claveB =
                `${b.tipo}|${b.color}|${b.seccion}`;

            return claveA.localeCompare(claveB);
        });

        // Sólo añadir índice si hay repetidos
        if (grupo.length > 1) {

            grupo.forEach((cable, index) => {

                resultado.push({
                    ...cable,
                    destino:
                        `${cable.destino} (${index + 1})`
                });

            });

        } else {

            resultado.push(grupo[0]);

        }

    });

    return resultado;

}

// =====================================
// CLAVE
// =====================================

function obtenerClave(cable) {
    return `${cable.origen}|${cable.destino}`;
}

// =====================================
// COMPARACION
// =====================================

function compararListas(listaA, listaB) {

    const mapaA = new Map();
    const mapaB = new Map();

    prepararLista(listaA)
        .forEach(cable =>
            mapaA.set(obtenerClave(cable), cable)
        );

    prepararLista(listaB)
        .forEach(cable =>
            mapaB.set(obtenerClave(cable), cable)
        );

    const resultados = [];

    const claves = new Set([
        ...mapaA.keys(),
        ...mapaB.keys()
    ]);

    for (const clave of claves) {

        const antiguo = mapaA.get(clave);
        const nuevo = mapaB.get(clave);

        if (!antiguo) {

            resultados.push({
                estado: "nuevo",
                cable: nuevo
            });

            continue;
        }

        if (!nuevo) {

            resultados.push({
                estado: "eliminado",
                cable: antiguo
            });

            continue;
        }

        let modificado = false;

        if (
            antiguo.tipo !== nuevo.tipo ||
            antiguo.color !== nuevo.color ||
            antiguo.seccion !== nuevo.seccion
        ) {
            modificado = true;
        }

        resultados.push({
            estado: modificado ? "modificado" : "igual",
            antiguo,
            nuevo
        });
    }

    resultados.sort((a, b) => {

        const origenA =
            a.nuevo?.origen ||
            a.antiguo?.origen ||
            a.cable?.origen;

        const origenB =
            b.nuevo?.origen ||
            b.antiguo?.origen ||
            b.cable?.origen;

        const destinoA =
            a.nuevo?.destino ||
            a.antiguo?.destino ||
            a.cable?.destino;

        const destinoB =
            b.nuevo?.destino ||
            b.antiguo?.destino ||
            b.cable?.destino;

        return (
            origenA.localeCompare(origenB) ||
            destinoA.localeCompare(destinoB)
        );
    });

    return resultados;
}

// =====================================
// FORMATEAR CAMBIO
// =====================================

function pintarCambio(viejo, nuevo) {

    if (viejo === nuevo) {
        return {
            html: viejo,
            cambiado: false
        };
    }

    return {
        html: `
            <span class="valor-antiguo">${viejo}</span>
            →
            <span class="valor-nuevo">${nuevo}</span>
        `,
        cambiado: true
    };
}

// =====================================
// RESUMEN
// =====================================

function pintarResumen(resultados) {

    const resumen = {
        nuevos: 0,
        eliminados: 0,
        modificados: 0,
        iguales: 0
    };

    resultados.forEach(r => {

        if (r.estado === "nuevo") resumen.nuevos++;
        if (r.estado === "eliminado") resumen.eliminados++;
        if (r.estado === "modificado") resumen.modificados++;
        if (r.estado === "igual") resumen.iguales++;

    });

    document.getElementById("resumen").innerHTML = `
    
    <div class="summary-container">

        <div class="card listaA">
            <span class="numero">${listaA.length}</span>
            Lista A
        </div>

        <div class="card listaB">
            <span class="numero">${listaB.length}</span>
            Lista B
        </div>

        <div class="card nuevos">
            <span class="numero">${resumen.nuevos}</span>
            Nuevos
        </div>

        <div class="card eliminados">
            <span class="numero">${resumen.eliminados}</span>
            Eliminados
        </div>

        <div class="card modificados">
            <span class="numero">${resumen.modificados}</span>
            Modificados
        </div>

        <div class="card iguales">
            <span class="numero">${resumen.iguales}</span>
            Sin cambios
        </div>

    </div>
    
    `;
}

// =====================================
// TABLA
// =====================================

function pintarTabla() {
    let resultados;


    if (filtroActual === "todos") {

        resultados = resultadosGlobales;

    }
    else if (filtroActual === "cambios") {

        resultados = resultadosGlobales.filter(
            r => r.estado !== "igual"
        );

    }
    else {

        resultados = resultadosGlobales.filter(
            r => r.estado === filtroActual
        );

    }

    let html = `
    
    <table>

        <thead>
            <tr>
                <th>Estado</th>
                <th>Origen</th>
                <th>Destino</th>
                <th>Tipo</th>
                <th>Color</th>
                <th>Sección</th>
            </tr>
        </thead>

        <tbody>
    
    `;

    resultados.forEach(r => {

        // SIN CAMBIOS

        if (r.estado === "igual") {

            html += `
            <tr class="fila-igual">
                <td class="estado">✅</td>
                <td>${r.nuevo.origen}</td>
                <td>${r.nuevo.destino}</td>
                <td>${r.nuevo.tipo}</td>
                <td>${r.nuevo.color}</td>
                <td>${r.nuevo.seccion}</td>
            </tr>
            `;
        }

        // NUEVO

        if (r.estado === "nuevo") {

            html += `
            <tr class="fila-nuevo">
                <td class="estado">➕</td>
                <td>${r.cable.origen}</td>
                <td>${r.cable.destino}</td>
                <td>${r.cable.tipo}</td>
                <td>${r.cable.color}</td>
                <td>${r.cable.seccion}</td>
            </tr>
            `;
        }

        // ELIMINADO

        if (r.estado === "eliminado") {

            html += `
            <tr class="fila-eliminado">
                <td class="estado">❌</td>
                <td>${r.cable.origen}</td>
                <td>${r.cable.destino}</td>
                <td>${r.cable.tipo}</td>
                <td>${r.cable.color}</td>
                <td>${r.cable.seccion}</td>
            </tr>
            `;
        }

        // MODIFICADO

        if (r.estado === "modificado") {

            const tipo = pintarCambio(
                r.antiguo.tipo,
                r.nuevo.tipo
            );

            const color = pintarCambio(
                r.antiguo.color,
                r.nuevo.color
            );

            const seccion = pintarCambio(
                r.antiguo.seccion,
                r.nuevo.seccion
            );

            html += `
            
            <tr class="fila-modificado">

                <td class="estado">⚠️</td>

                <td>${r.nuevo.origen}</td>

                <td>${r.nuevo.destino}</td>

                <td class="${tipo.cambiado ? 'celda-cambio' : ''}">
                    ${tipo.html}
                </td>

                <td class="${color.cambiado ? 'celda-cambio' : ''}">
                    ${color.html}
                </td>

                <td class="${seccion.cambiado ? 'celda-cambio' : ''}">
                    ${seccion.html}
                </td>

            </tr>
            
            `;
        }

    });

    html += `
        </tbody>
    </table>
    `;

    document.getElementById("resultados").innerHTML = html;
}

// =====================================
// EJECUCION
// =====================================

async function ejecutarComparacion() {

    const archivoA =
        document.getElementById("archivoA")
        .files[0];

    const archivoB =
        document.getElementById("archivoB")
        .files[0];

    if (!archivoA || !archivoB) {

        alert(
            "Seleccione ambos archivos Excel"
        );

        return;
    }

    try {

        const excelA =
            await leerExcel(archivoA);

        const excelB =
            await leerExcel(archivoB);

        if (!validarColumnas(excelA)) {

            alert(
                "El Excel A no contiene las columnas requeridas"
            );

            return;
        }

        if (!validarColumnas(excelB)) {

            alert(
                "El Excel B no contiene las columnas requeridas"
            );

            return;
        }

        listaA =
            convertirFormato(excelA);

        listaB =
            convertirFormato(excelB);

        resultadosGlobales =
            compararListas(
                listaA,
                listaB
            );

        pintarResumen(
            resultadosGlobales
        );

        pintarTabla();

    }
    catch(error) {

        console.error(error);

        alert(
            "Error leyendo los archivos Excel"
        );

    }

}

document
    .getElementById("btnComparar")
    .addEventListener(
        "click",
        ejecutarComparacion
    );

document
    .querySelectorAll(".filtro")
    .forEach(btn => {

        btn.addEventListener("click", function () {

            document
                .querySelectorAll(".filtro")
                .forEach(b => b.classList.remove("activo"));

            this.classList.add("activo");

            filtroActual =
                this.dataset.filtro;

            pintarTabla();
        });

    });

document
    .getElementById("btnExportar")
    .addEventListener("click", exportarExcel);

document.getElementById("versionFooter").textContent =
    `Comparador de Cableado v${VERSION}`;
