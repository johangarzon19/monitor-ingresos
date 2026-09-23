// ============================================================
// MONITOR DE INGRESOS
// ============================================================

let datosProgramacion = [];
let datosIngresos = [];
let resultados = [];


// ============================================================
// 1. CARGAR ARCHIVO EXCEL
// ============================================================

document
    .getElementById("archivoExcel")
    .addEventListener("change", function (evento) {

        const archivo = evento.target.files[0];

        if (!archivo) {
            return;
        }

        const mensaje = document.getElementById("mensaje");

        mensaje.textContent = "Procesando archivo...";

        const lector = new FileReader();

        lector.onload = function (e) {

            try {

                const datos = new Uint8Array(e.target.result);

                // IMPORTANTE:
                // Dejamos cellDates en false para poder interpretar
                // correctamente las fechas y horas de Excel.
                const libro = XLSX.read(datos, {
                    type: "array",
                    cellDates: false
                });

                console.log("Hojas encontradas:", libro.SheetNames);


                // ------------------------------------------------
                // VALIDAR HOJAS
                // ------------------------------------------------

                if (!libro.Sheets["programacion"]) {

                    mensaje.textContent =
                        "Error: no se encontró la hoja 'programacion'.";

                    return;
                }

                if (!libro.Sheets["ingresos"]) {

                    mensaje.textContent =
                        "Error: no se encontró la hoja 'ingresos'.";

                    return;
                }


                // ------------------------------------------------
                // CONVERTIR HOJAS A JSON
                // ------------------------------------------------

                datosProgramacion =
                    XLSX.utils.sheet_to_json(
                        libro.Sheets["programacion"],
                        {
                            defval: ""
                        }
                    );


                datosIngresos =
                    XLSX.utils.sheet_to_json(
                        libro.Sheets["ingresos"],
                        {
                            defval: ""
                        }
                    );


                // ------------------------------------------------
                // MOSTRAR DATOS EN CONSOLA
                // ------------------------------------------------

                console.log(
                    "PROGRAMACION:",
                    datosProgramacion
                );

                console.log(
                    "INGRESOS:",
                    datosIngresos
                );


                if (datosProgramacion.length > 0) {

                    console.log(
                        "Columnas PROGRAMACION:",
                        Object.keys(datosProgramacion[0])
                    );
                }


                if (datosIngresos.length > 0) {

                    console.log(
                        "Columnas INGRESOS:",
                        Object.keys(datosIngresos[0])
                    );
                }


                // ------------------------------------------------
                // ACTUALIZAR INDICADORES
                // ------------------------------------------------

                document.getElementById(
                    "totalProgramados"
                ).textContent =
                    datosProgramacion.length;


                document.getElementById(
                    "totalIngresos"
                ).textContent =
                    datosIngresos.length;


                // ------------------------------------------------
                // REALIZAR CRUCE
                // ------------------------------------------------

                resultados =
                    cruzarDatos(
                        datosProgramacion,
                        datosIngresos
                    );


                console.log(
                    "RESULTADOS:",
                    resultados
                );


                // ------------------------------------------------
                // MOSTRAR RESULTADOS
                // ------------------------------------------------

                mostrarResultados(resultados);


                mensaje.textContent =
                    "Archivo cargado correctamente.";

            }
            catch (error) {

                console.error(
                    "ERROR:",
                    error
                );

                mensaje.textContent =
                    "Ocurrió un error al procesar el archivo.";

                document.getElementById(
                    "resultado"
                ).innerHTML = `
                    <p>
                        ❌ No se pudo procesar el archivo.
                        Revisa la consola del navegador.
                    </p>
                `;
            }

        };


        lector.readAsArrayBuffer(archivo);

    });


// ============================================================
// 2. NORMALIZAR TEXTO
// ============================================================

function normalizarTexto(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";
    }

    return String(valor)
        .trim()
        .toUpperCase();

}


// ============================================================
// 3. NORMALIZAR DOCUMENTO
// ============================================================

function normalizarDocumento(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";
    }

    return String(valor)
        .replace(/\D/g, "")
        .trim();

}


// ============================================================
// 4. NORMALIZAR PDV
// ============================================================

function normalizarPDV(valor) {

    if (
        valor === null ||
        valor === undefined
    ) {

        return "";
    }

    return String(valor)
        .trim()
        .replace(/\.0$/, "");

}


// ============================================================
// 5. OBTENER FECHA
// ============================================================

function obtenerFecha(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {

        return null;
    }


    // --------------------------------------------------------
    // Excel almacena las fechas como números.
    // --------------------------------------------------------

    if (typeof valor === "number") {

        try {

            const fechaExcel =
                XLSX.SSF.parse_date_code(valor);

            if (!fechaExcel) {

                return null;
            }

            return new Date(
                fechaExcel.y,
                fechaExcel.m - 1,
                fechaExcel.d
            );

        }
        catch (error) {

            return null;
        }
    }


    // --------------------------------------------------------
    // Si ya es Date
    // --------------------------------------------------------

    if (valor instanceof Date) {

        return new Date(
            valor.getFullYear(),
            valor.getMonth(),
            valor.getDate()
        );
    }


    // --------------------------------------------------------
    // Si viene como texto
    // --------------------------------------------------------

    const texto =
        String(valor).trim();


    // DD/MM/YYYY

    if (texto.includes("/")) {

        const partes =
            texto.split("/");

        if (partes.length >= 3) {

            const dia =
                parseInt(partes[0], 10);

            const mes =
                parseInt(partes[1], 10);

            const anio =
                parseInt(
                    partes[2].split(" ")[0],
                    10
                );

            if (
                !isNaN(dia) &&
                !isNaN(mes) &&
                !isNaN(anio)
            ) {

                return new Date(
                    anio,
                    mes - 1,
                    dia
                );
            }
        }
    }


    // YYYY-MM-DD

    if (texto.includes("-")) {

        const partes =
            texto.split("-");

        if (partes.length >= 3) {

            const anio =
                parseInt(partes[0], 10);

            const mes =
                parseInt(partes[1], 10);

            const dia =
                parseInt(
                    partes[2].split(" ")[0],
                    10
                );

            if (
                !isNaN(anio) &&
                !isNaN(mes) &&
                !isNaN(dia)
            ) {

                return new Date(
                    anio,
                    mes - 1,
                    dia
                );
            }
        }
    }


    return null;
}


// ============================================================
// 6. OBTENER CLAVE DE FECHA
// ============================================================

function obtenerClaveFecha(valor) {

    const fecha =
        obtenerFecha(valor);

    if (!fecha) {

        return "";
    }

    const anio =
        fecha.getFullYear();

    const mes =
        String(
            fecha.getMonth() + 1
        ).padStart(2, "0");

    const dia =
        String(
            fecha.getDate()
        ).padStart(2, "0");


    return `${anio}-${mes}-${dia}`;
}


// ============================================================
// 7. OBTENER MINUTOS DE UNA HORA
// ============================================================

function obtenerMinutosHora(valor) {

    if (
        valor === null ||
        valor === undefined ||
        valor === ""
    ) {
        return null;
    }


    // ========================================================
    // 1. EXCEL GUARDA FECHA/HORA COMO NÚMERO
    // ========================================================

    if (typeof valor === "number") {

        const fraccionDia = valor % 1;

        return Math.round(
            fraccionDia * 24 * 60
        );
    }


    // ========================================================
    // 2. SI YA ES UN OBJETO DATE
    // ========================================================

    if (valor instanceof Date) {

        return (
            valor.getHours() * 60 +
            valor.getMinutes()
        );
    }


    // ========================================================
    // 3. SI VIENE COMO TEXTO
    // ========================================================

    let texto = String(valor).trim();

    if (!texto) {
        return null;
    }


    // ========================================================
    // BUSCAR UNA HORA DENTRO DEL TEXTO
    //
    // Ejemplos que puede recibir:
    //
    // 22/09/2026 08:28
    // 22/09/2026 08:28:15
    // 2026-09-22 08:28
    // 2026-09-22T08:28:00
    // 08:28
    // 08:28:00
    // ========================================================

    const coincidenciaHora =
        texto.match(
            /(\d{1,2}):(\d{2})(?::(\d{2}))?\s*(AM|PM)?/i
        );


    if (coincidenciaHora) {

        let horas =
            parseInt(
                coincidenciaHora[1],
                10
            );

        const minutos =
            parseInt(
                coincidenciaHora[2],
                10
            );

        const periodo =
            coincidenciaHora[4]
                ? coincidenciaHora[4].toUpperCase()
                : "";


        // Validar minutos

        if (
            minutos < 0 ||
            minutos > 59
        ) {
            return null;
        }


        // AM / PM

        if (
            periodo === "PM" &&
            horas < 12
        ) {
            horas += 12;
        }


        if (
            periodo === "AM" &&
            horas === 12
        ) {
            horas = 0;
        }


        // Validar hora

        if (
            horas < 0 ||
            horas > 23
        ) {
            return null;
        }


        return (
            horas * 60 +
            minutos
        );
    }


    // ========================================================
    // 4. SI NO ENCONTRÓ UNA HORA
    // ========================================================

    return null;
}

function minutosAHora(minutos) {

    if (minutos === null || minutos === undefined) {
        return "";
    }

    const horas =
        Math.floor(minutos / 60);

    const minutosRestantes =
        minutos % 60;

    return (
        String(horas).padStart(2, "0") +
        ":" +
        String(minutosRestantes).padStart(2, "0")
    );
}


// ============================================================
// 8. FORMATEAR HORA
// ============================================================

function formatearHora(minutos) {

    if (
        minutos === null ||
        minutos === undefined
    ) {

        return "--:--";
    }


    minutos =
        Math.round(minutos);


    const horas =
        Math.floor(minutos / 60);

    const mins =
        minutos % 60;


    return (
        String(horas).padStart(2, "0") +
        ":" +
        String(mins).padStart(2, "0")
    );
}


// ============================================================
// 9. FORMATEAR FECHA
// ============================================================

function formatearFecha(fecha) {

    if (!fecha) {

        return "-";
    }


    const dia =
        String(
            fecha.getDate()
        ).padStart(2, "0");


    const mes =
        String(
            fecha.getMonth() + 1
        ).padStart(2, "0");


    const anio =
        fecha.getFullYear();


    return `${dia}/${mes}/${anio}`;
}


// ============================================================
// 10. FORMATEAR DIFERENCIA
// ============================================================

function formatearDiferencia(
    diferencia
) {

    if (
        diferencia === null ||
        diferencia === undefined
    ) {

        return "--";
    }


    if (diferencia < 0) {

        return `${Math.abs(diferencia)} min antes`;
    }


    if (diferencia === 0) {

        return "0 min";
    }


    return `+${diferencia} min`;
}


// ============================================================
// 11. BUSCAR COLUMNA DE HORA PROGRAMADA
// ============================================================
//
// Como todavía estamos verificando el nombre exacto de la
// columna en tu Excel, buscamos varias posibilidades.
// ============================================================

function obtenerColumnaHoraProgramada(
    registro
) {

    const posibles = [

        "Hora Entrada",

        "Hora de Entrada",

        "Hora Entrada Programada",

        "Hora de Entrada Programada",

        "Hora Ingreso",

        "Hora Programada",

        "Hora inicio",

        "Hora Inicio"

    ];


    for (
        const nombre of posibles
    ) {

        if (
            Object.prototype.hasOwnProperty.call(
                registro,
                nombre
            )
        ) {

            return registro[nombre];
        }
    }


    // Búsqueda flexible

    const columnas =
        Object.keys(registro);


    const columnaEncontrada =
        columnas.find(function (columna) {

            const texto =
                normalizarTexto(columna);

            return (
                texto.includes("HORA") &&
                (
                    texto.includes("ENTRADA") ||
                    texto.includes("INGRESO") ||
                    texto.includes("INICIO")
                )
            );

        });


    if (columnaEncontrada) {

        return registro[columnaEncontrada];
    }


    return null;
}


// ============================================================
// 12. CREAR CLAVE DE CRUCE
// ============================================================

function crearClave(
    documento,
    pdv,
    fecha
) {

    return (
        normalizarDocumento(documento) +
        "|" +
        normalizarPDV(pdv) +
        "|" +
        obtenerClaveFecha(fecha)
    );
}


// ============================================================
// 13. CRUZAR DATOS
// ============================================================

function cruzarDatos(programacion,ingresos) {

    const mapaIngresos =
        new Map();


    // --------------------------------------------------------
    // INDEXAR INGRESOS
    // --------------------------------------------------------

    ingresos.forEach(function (ingreso) {

        const documento =
            ingreso["Doc. Identidad"];

        const pdv =
            ingreso["Id Pdv"];

        const fechaRegistro =
            ingreso["Fecha registro"];


        const fecha =
            obtenerFecha(
                fechaRegistro
            );


        if (!fecha) {

            return;
        }


        const minutosIngreso =
            obtenerMinutosHora(
                fechaRegistro
            );


        if (
            minutosIngreso === null
        ) {

            return;
        }


        const clave =
            crearClave(
                documento,
                pdv,
                fechaRegistro
            );


        // Si existen varios registros para el mismo
        // empleado + PDV + fecha, tomamos el primero.

        if (!mapaIngresos.has(clave)) {

            mapaIngresos.set(
                clave,
                {
                    documento:
                        normalizarDocumento(
                            documento
                        ),

                    pdv:
                        normalizarPDV(
                            pdv
                        ),

                    fecha: fecha,

                    minutosIngreso:
                        minutosIngreso,

                    registroOriginal:
                        ingreso
                }
            );

        }
        else {

            const existente =
                mapaIngresos.get(clave);


            // Conservamos el ingreso más temprano.

            if (
                minutosIngreso <
                existente.minutosIngreso
            ) {

                mapaIngresos.set(
                    clave,
                    {
                        documento:
                            normalizarDocumento(
                                documento
                            ),

                        pdv:
                            normalizarPDV(
                                pdv
                            ),

                        fecha: fecha,

                        minutosIngreso:
                            minutosIngreso,

                        registroOriginal:
                            ingreso
                    }
                );
            }
        }

    });


    console.log(
        "Mapa de ingresos:",
        mapaIngresos
    );


    // --------------------------------------------------------
    // RECORRER PROGRAMACIÓN
    // --------------------------------------------------------

    const resultado = [];


    programacion.forEach(
        function (programado) {

            const documento =
                programado["Nro Doc Ident"];

            const pdv =
                programado["ID PDV"];

            const fechaProgramacion =
                programado[
                    "Fecha de Programación"
                ];


            // Buscar hora programada
            const horaProgramada =
                obtenerColumnaHoraProgramada(
                    programado
                );


            const fecha =
                obtenerFecha(
                    fechaProgramacion
                );


            const minutosProgramados =
                obtenerMinutosHora(
                    horaProgramada
                );


            const clave =
                crearClave(
                    documento,
                    pdv,
                    fechaProgramacion
                );


            const ingreso =
                mapaIngresos.get(
                    clave
                );


            const fila = {
            documento:
                normalizarDocumento(
                    documento
                ),

            pdv:
                normalizarPDV(
                    pdv
                ),

            nombrePDV:
                programado["Nombre alias"] || "",

            funcionario:
                programado["Funcionario"] || "",

            fecha:
                fecha,

            horaProgramada:
                minutosProgramados,

            horaReal:
                null,

            diferencia:
                null,

            estado:
                "AÚN NO MARCA"
        };


            // ------------------------------------------------
            // SI NO MARCÓ
            // ------------------------------------------------

            if (!ingreso) {

                resultado.push(fila);

                return;
            }


            // ------------------------------------------------
            // SI MARCÓ
            // ------------------------------------------------

            fila.horaReal =
                ingreso.minutosIngreso;


            if (
                minutosProgramados !== null
            ) {

                fila.diferencia =
                    ingreso.minutosIngreso -
                    minutosProgramados;


                if (
                    fila.diferencia < 0
                ) {

                    fila.estado =
                        "TEMPRANO";

                }
                else if (
                    fila.diferencia === 0
                ) {

                    fila.estado =
                        "A TIEMPO";

                }
                else {

                    fila.estado =
                        "TARDE";
                }

            }


            resultado.push(fila);

        }
    );


    return resultado;
}


// ============================================================
// 14. MOSTRAR RESULTADOS
// ============================================================

function mostrarResultados(datos, actualizarPDV = true) {

    const contenedor =
        document.getElementById(
            "resultado"
        );

    


    if (!contenedor) {

        return;
    }
    if (actualizarPDV) {

        cargarFiltroPDV();

        cargarFiltroFuncionario();

    }

    actualizarAlertasRecurrencia(datos);

    function cargarFiltroFuncionario() {

        const selectFuncionario =
            document.getElementById("filtroFuncionario");

        if (!selectFuncionario) return;


        const funcionarios = [
            ...new Set(
                resultados
                    .map(
                        fila => fila.funcionario
                    )
                    .filter(
                        funcionario =>
                            funcionario !== ""
                    )
            )
        ];


        funcionarios.sort((a, b) =>
            String(a).localeCompare(
                String(b),
                undefined,
                { numeric: true }
            )
        );


        selectFuncionario.innerHTML = `
            <option value="">
                Todos los funcionarios
            </option>
        `;


        funcionarios.forEach(
            function (funcionario) {

            const opcion =
                document.createElement("option");

            opcion.value = funcionario;

            opcion.textContent = funcionario;

            selectFuncionario.appendChild(
                opcion
            );
            }
            );
    }
    // --------------------------------------------------------
    // CALCULAR INDICADORES
    // --------------------------------------------------------

    const totalProgramados =
    datos.length;


    const totalIngresos =
    datos.filter(
        function (fila) {
            return fila.horaReal !== null;
        }
    ).length;

    const totalTarde =
        datos.filter(
            function (fila) {

                return fila.estado === "TARDE";

            }
        ).length;


    const totalPendientes =
        datos.filter(
            function (fila) {

                return (
                    fila.estado ===
                    "AÚN NO MARCA"
                );

            }
        ).length;


    // Actualizar tarjetas

    document.getElementById(
    "totalProgramados"
    ).textContent =
    totalProgramados;


    document.getElementById(
    "totalIngresos"
    ).textContent =
    totalIngresos;

    document.getElementById(
        "totalTarde"
    ).textContent =
        totalTarde;


    document.getElementById(
        "totalPendientes"
    ).textContent =
        totalPendientes;


    // --------------------------------------------------------
    // SI NO HAY RESULTADOS
    // --------------------------------------------------------

    if (datos.length === 0) {

        contenedor.innerHTML = `
            <p>
                No se encontraron registros
                de programación.
            </p>
        `;

        return;
    }


    // --------------------------------------------------------
    // TABLA
    // --------------------------------------------------------

    let html = `

        <div class="tabla-contenedor">

            <table>

                <thead>

                    <tr>

                        <th>Funcionario</th>

                        <th>PDV</th>

                        <th>Fecha</th>

                        <th>Hora programada</th>

                        <th>Hora ingreso</th>

                        <th>Diferencia</th>

                        <th>Estado</th>

                    </tr>

                </thead>

                <tbody>

    `;


    datos.forEach(
        function (fila) {

            let clase = "";
            let icono = "";


            if (
                fila.estado ===
                "TEMPRANO"
            ) {

                clase = "temprano";
                icono = "🟢";

            }
            else if (
                fila.estado ===
                "A TIEMPO"
            ) {

                clase = "a-tiempo";
                icono = "🔵";

            }
            else if (
                fila.estado ===
                "TARDE"
            ) {

                clase = "tarde";
                icono = "🔴";

            }
            else {

                clase = "pendiente";
                icono = "🟡";

            }


            html += `

                <tr>

                    <td>
                        ${fila.funcionario || "-"}
                    </td>

                    <td>
                        ${fila.nombrePDV || "-"}
                    </td>

                    <td>
                        ${formatearFecha(
                            fila.fecha
                        )}
                    </td>

                    <td>
                        ${formatearHora(
                            fila.horaProgramada
                        )}
                    </td>

                    <td>
                        ${formatearHora(
                            fila.horaReal
                        )}
                    </td>

                    <td>
                        ${formatearDiferencia(
                            fila.diferencia
                        )}
                    </td>

                    <td>

                        <span class="estado ${clase}">

                            ${icono}

                            ${fila.estado}

                        </span>

                    </td>

                </tr>

            `;

        }
    );


    html += `

                </tbody>

            </table>

        </div>

    `;


    contenedor.innerHTML =
        html;

}

// ============================================================
// FILTROS
// ============================================================


// ------------------------------------------------------------
// 1. LLENAR LISTA DE PDV
// ------------------------------------------------------------

function cargarFiltroPDV() {

    const selectPDV =
        document.getElementById("filtroPDV");

    if (!selectPDV) return;


    // Obtener código y nombre de cada PDV
    const pdvs = [
        ...new Map(
            resultados
                .filter(
                    fila => fila.pdv !== ""
                )
                .map(fila => [
                    fila.pdv,
                    fila.nombrePDV
                ])
        ).entries()
    ];


    // Ordenar los PDV por nombre
    pdvs.sort((a, b) =>
        String(a[1]).localeCompare(
            String(b[1]),
            undefined,
            { numeric: true }
        )
    );


    // Limpiar el listado actual
    selectPDV.innerHTML = `
        <option value="">
            Todos los PDV
        </option>
    `;


    // Crear las opciones
    pdvs.forEach(function ([codigo, nombre]) {

        const opcion =
            document.createElement("option");

        // Internamente seguimos utilizando el código
        opcion.value = codigo;

        // El usuario ve el nombre
        opcion.textContent =
            nombre || codigo;

        selectPDV.appendChild(opcion);
    });
}


// ------------------------------------------------------------
// 2. APLICAR FILTROS
// ------------------------------------------------------------

function aplicarFiltros() {

    const filtroFechaInicio =
        document.getElementById("filtroFechaInicio").value;

    const filtroFechaFin =
        document.getElementById("filtroFechaFin").value;


    const filtroPDV =
        document.getElementById(
            "filtroPDV"
        ).value;


    const filtroEstado =
        document.getElementById(
            "filtroEstado"
        ).value;


    const filtroFuncionario =
    document.getElementById("filtroFuncionario").value;

    document.getElementById("btnExportarExcel")
    .addEventListener(
        "click",
        exportarExcel
    );


    // --------------------------------------------------------
    // Filtrar resultados
    // --------------------------------------------------------

    const filtrados =
        resultados.filter(function (fila) {


            // -----------------------------------------------
            // FILTRO FECHA
            // -----------------------------------------------

            if (filtroFechaInicio) {

    if (
        fila.fecha &&
        obtenerClaveFecha(fila.fecha) <
            filtroFechaInicio
            ) {
                return false;
            }
        }


        if (filtroFechaFin) {

            if (
                fila.fecha &&
                obtenerClaveFecha(fila.fecha) >
                    filtroFechaFin
            ) {
                return false;
            }
        }


            // -----------------------------------------------
            // FILTRO PDV
            // -----------------------------------------------

            if (filtroPDV) {

                if (
                    fila.pdv !== filtroPDV
                ) {

                    return false;
                }

            }


            // -----------------------------------------------
            // FILTRO ESTADO
            // -----------------------------------------------

            if (filtroEstado) {

                if (
                    fila.estado !==
                    filtroEstado
                ) {

                    return false;
                }

            }


            // -----------------------------------------------
            // FILTRO DOCUMENTO
            // -----------------------------------------------

            if (filtroFuncionario) {

                if (
                    fila.funcionario !==
                    filtroFuncionario
                ) {
                    return false;
                }
            }


            return true;

        });


    // Mostrar resultados filtrados
    mostrarResultados(
        filtrados,
        false
    );

}


// ------------------------------------------------------------
// 3. EVENTOS DE LOS FILTROS
// ------------------------------------------------------------

document.getElementById("filtroFechaInicio")
    .addEventListener(
        "change",
        aplicarFiltros
    );


document.getElementById("filtroFechaFin")
    .addEventListener(
        "change",
        aplicarFiltros
    );


document
    .getElementById("filtroPDV")
    .addEventListener(
        "change",
        aplicarFiltros
    );


document
    .getElementById("filtroEstado")
    .addEventListener(
        "change",
        aplicarFiltros
    );


document.getElementById("filtroFuncionario")
    .addEventListener(
        "change",
        aplicarFiltros
    );


// ------------------------------------------------------------
// 4. LIMPIAR FILTROS
// ------------------------------------------------------------

document
    .getElementById("btnLimpiarFiltros")
    .addEventListener(
        "click",
        function () {

            document.getElementById(
            "filtroFechaInicio"
            ).value = "";

            document.getElementById(
            "filtroFechaFin"
            ).value = "";


            document.getElementById(
                "filtroPDV"
            ).value = "";


            document.getElementById(
                "filtroEstado"
            ).value = "";


            document.getElementById("filtroFuncionario").value = "";


            mostrarResultados(
                resultados,
                false
            );

        }
    );
    function exportarExcel() {

    // Obtener los filtros actuales
    const filtroFechaInicio =
        document.getElementById("filtroFechaInicio").value;

    const filtroFechaFin =
        document.getElementById("filtroFechaFin").value;

    const filtroPDV =
        document.getElementById("filtroPDV").value;

    const filtroEstado =
        document.getElementById("filtroEstado").value;

    const filtroFuncionario =
        document.getElementById("filtroFuncionario").value;


    // Aplicar los mismos filtros de la pantalla
    const datosExportar =
        resultados.filter(function (fila) {

            if (filtroFechaInicio) {

                if (
                    fila.fecha &&
                    obtenerClaveFecha(fila.fecha) <
                        filtroFechaInicio
                ) {
                    return false;
                }
            }


            if (filtroFechaFin) {

                if (
                    fila.fecha &&
                    obtenerClaveFecha(fila.fecha) >
                        filtroFechaFin
                ) {
                    return false;
                }
            }


            if (filtroPDV) {

                if (fila.pdv !== filtroPDV) {
                    return false;
                }
            }


            if (filtroEstado) {

                if (fila.estado !== filtroEstado) {
                    return false;
                }
            }


            if (filtroFuncionario) {

                if (
                    fila.funcionario !==
                    filtroFuncionario
                ) {
                    return false;
                }
            }


            return true;
        });


    // Validar que existan datos
    if (datosExportar.length === 0) {

        alert(
            "No hay datos para exportar con los filtros seleccionados."
        );

        return;
    }


    // Preparar información para Excel
    const datosExcel =
        datosExportar.map(function (fila) {

            return {

                "Funcionario":
                    fila.funcionario || "",

                "PDV":
                    fila.nombrePDV || "",

                "Fecha":
                    obtenerClaveFecha(fila.fecha),

                "Hora programada":
                    fila.horaProgramada !== null
                        ? minutosAHora(fila.horaProgramada)
                        : "",

                "Hora real":
                    fila.horaReal !== null
                        ? minutosAHora(fila.horaReal)
                        : "",

                "Diferencia minutos":
                    fila.diferencia !== null
                        ? fila.diferencia
                        : "",

                "Estado":
                    fila.estado || ""
            };
        });


    // Crear archivo Excel
    const hoja =
        XLSX.utils.json_to_sheet(datosExcel);

    const libro =
        XLSX.utils.book_new();

    XLSX.utils.book_append_sheet(
        libro,
        hoja,
        "Detalle"
    );


    // Nombre del archivo
    let nombreArchivo =
        "Reporte_Ingresos";

    if (
        filtroFechaInicio &&
        filtroFechaFin
    ) {

        nombreArchivo +=
            "_" +
            filtroFechaInicio +
            "_al_" +
            filtroFechaFin;

    } else if (filtroFechaInicio) {

        nombreArchivo +=
            "_desde_" +
            filtroFechaInicio;

    } else if (filtroFechaFin) {

        nombreArchivo +=
            "_hasta_" +
            filtroFechaFin;
    }


    nombreArchivo += ".xlsx";


    // Descargar Excel
    XLSX.writeFile(
        libro,
        nombreArchivo
    );
}
    function actualizarAlertasRecurrencia(datos) {

    const contenedor =
        document.getElementById("alertasRecurrencia");

    if (!contenedor) return;


    // Agrupar información por funcionario
    const resumen = {};


    datos.forEach(function (fila) {

        const funcionario =
            fila.funcionario || "Sin funcionario";


        if (!resumen[funcionario]) {

            resumen[funcionario] = {
                programados: 0,
                ingresos: 0,
                tardanzas: 0,
                minutosTarde: 0,
                ultimaTardanza: null
            };
        }


        // Día programado
        resumen[funcionario].programados++;


        // Si registró ingreso
        if (fila.horaReal !== null) {

            resumen[funcionario].ingresos++;
        }


        // Si llegó tarde
        if (fila.estado === "TARDE") {

            resumen[funcionario].tardanzas++;


            if (fila.diferencia !== null) {

                resumen[funcionario].minutosTarde +=
                    fila.diferencia;
            }


            // Guardar la última fecha de tardanza
            if (
                !resumen[funcionario].ultimaTardanza ||
                obtenerClaveFecha(fila.fecha) >
                    obtenerClaveFecha(
                        resumen[funcionario].ultimaTardanza
                    )
            ) {

                resumen[funcionario].ultimaTardanza =
                    fila.fecha;
            }
        }
    });


    // Convertir el objeto en arreglo
    const funcionarios =
        Object.entries(resumen)
            .map(function ([nombre, datosFuncionario]) {

                const porcentaje =
                    datosFuncionario.ingresos > 0
                        ? (
                            datosFuncionario.tardanzas /
                            datosFuncionario.ingresos
                        ) * 100
                        : 0;


                const promedio =
                    datosFuncionario.tardanzas > 0
                        ? (
                            datosFuncionario.minutosTarde /
                            datosFuncionario.tardanzas
                        )
                        : 0;


                return {
                    funcionario: nombre,
                    programados:
                        datosFuncionario.programados,
                    ingresos:
                        datosFuncionario.ingresos,
                    tardanzas:
                        datosFuncionario.tardanzas,
                    porcentaje: porcentaje,
                    promedioMinutos: promedio,
                    ultimaTardanza:
                        datosFuncionario.ultimaTardanza
                };
            });


    // Solo funcionarios con 3 o más tardanzas
    const alertas =
        funcionarios.filter(function (funcionario) {

            return funcionario.tardanzas >= 3;

        });


    // Ordenar de mayor a menor cantidad de tardanzas
    alertas.sort(function (a, b) {

        return b.tardanzas - a.tardanzas;

    });


    // Si no existen alertas
    if (alertas.length === 0) {

        contenedor.innerHTML = `
            <div class="sin-alertas">
                ✓ No se identifican funcionarios con
                3 o más llegadas tarde en el período seleccionado.
            </div>
        `;

        return;
    }


    // Construir tabla
    let html = `
        <div class="alerta-resumen">
            Se identificaron
            <strong>${alertas.length}</strong>
            funcionario(s) con 3 o más llegadas tarde
            durante el período seleccionado.
        </div>

        <div class="tabla-alertas-contenedor">

            <table class="tabla-alertas">

                <thead>
                    <tr>
                        <th>Funcionario</th>
                        <th>Programados</th>
                        <th>Ingresos</th>
                        <th>Tardanzas</th>
                        <th>% tardanzas</th>
                        <th>Prom. minutos tarde</th>
                        <th>Última tardanza</th>
                    </tr>
                </thead>

                <tbody>
    `;


    alertas.forEach(function (funcionario) {

        html += `
            <tr>

                <td>
                    ${funcionario.funcionario}
                </td>

                <td>
                    ${funcionario.programados}
                </td>

                <td>
                    ${funcionario.ingresos}
                </td>

                <td class="cantidad-tardanzas">
                    ${funcionario.tardanzas}
                </td>

                <td>
                    ${funcionario.porcentaje.toFixed(1)}%
                </td>

                <td>
                    ${funcionario.promedioMinutos.toFixed(1)} min
                </td>

                <td>
                    ${
                        funcionario.ultimaTardanza
                            ? obtenerClaveFecha(
                                funcionario.ultimaTardanza
                            )
                            : "-"
                    }
                </td>

            </tr>
        `;
    });


    html += `
                </tbody>

            </table>

        </div>
    `;


    contenedor.innerHTML = html;
}

// ================================
// CAMBIO DE TEMA
// ================================

const btnTema =
    document.getElementById("btnTema");

btnTema.addEventListener(
    "change",
    function () {

        document.body.classList.toggle(
            "tema-oscuro",
            btnTema.checked
        );

    }
);