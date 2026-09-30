let subsistemaSeleccionado = "";
let fmrSeleccionado = "";
let poSeleccionada = "";
let threeweekSeleccionado = "";

const estado = {
    cargado: false,
    normalizado: [],
    relaciones: [],
    detalle: [],
    semanas: [],
    fmrs: []
};

const threeweekSelect = document.getElementById("threeweek-select");
const fmrSearch = document.getElementById("fmr-search");
const listaFMRsGeneral = document.getElementById("lista-fmrs-general");
const fmrOptions = document.getElementById("fmr-options");
const infoSubsistema = document.getElementById("info-subsistema");
const listaSubsistemas = document.getElementById("lista-subsistemas");
const subsistemaSeleccionadoDiv = document.getElementById("subsistema-seleccionado");
const nombreSubsistema = document.getElementById("nombre-subsistema");
const listaFMR = document.getElementById("lista-fmr");
const resultadoFMRSearch = document.getElementById("resultado-fmr-search");
const fmrEncontrado = document.getElementById("fmr-encontrado");
const subsistemasDelFMR = document.getElementById("subsistemas-del-fmr");
const listaSubsistemasFMR = document.getElementById("lista-subsistemas-fmr");
const detalleFMR = document.getElementById("detalle-fmr");
const listaPO = document.getElementById("lista-po");
const listaProcurement = document.getElementById("lista-procurement");
const estadoCarga = document.getElementById("estado-carga");
const archivoExcel = document.getElementById("archivo-excel");

infoSubsistema.style.display = "none";
resultadoFMRSearch.style.display = "none";
detalleFMR.style.display = "none";
subsistemaSeleccionadoDiv.style.display = "none";
threeweekSelect.disabled = true;
fmrSearch.disabled = true;

function texto(valor) {
    return valor === null || valor === undefined ? "" : String(valor).trim();
}

function escaparHTML(valor) {
    return texto(valor).replace(/[&<>"']/g, function(caracter) {
        return { "&": "&amp;", "<": "&lt;", ">": "&gt;", "\"": "&quot;", "'": "&#039;" }[caracter];
    });
}

function fecha(valor) {
    if (valor instanceof Date && !Number.isNaN(valor.getTime())) {
        return valor;
    }
    if (typeof valor === "number") {
        const resultado = new Date(1899, 11, 30);
        resultado.setDate(resultado.getDate() + Math.floor(valor));
        return resultado;
    }
    const valorTexto = texto(valor);
    if (!valorTexto) {
        return null;
    }
    const partes = valorTexto.split(/[\/\-.]/).map(Number);
    if (partes.length === 3 && partes.every(Number.isFinite)) {
        return partes[0] > 31
            ? new Date(partes[0], partes[1] - 1, partes[2])
            : new Date(partes[2], partes[1] - 1, partes[0]);
    }
    const resultado = new Date(valorTexto);
    return Number.isNaN(resultado.getTime()) ? null : resultado;
}

function fechaTexto(valor) {
    const valorFecha = fecha(valor);
    if (!valorFecha) {
        return "";
    }
    return [
        String(valorFecha.getDate()).padStart(2, "0"),
        String(valorFecha.getMonth() + 1).padStart(2, "0"),
        valorFecha.getFullYear()
    ].join("/");
}

function formatoETA(valor) {
    const valores = valor instanceof Date || typeof valor === "number"
        ? [valor]
        : texto(valor).replace(/[;\n]/g, ",").split(",").filter(Boolean);
    const meses = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    return valores.map(function(item) {
        const valorFecha = fecha(item);
        if (!valorFecha) {
            return texto(item);
        }
        return `${meses[valorFecha.getMonth()]} ${String(valorFecha.getDate()).padStart(2, "0")} ${valorFecha.getFullYear()}`;
    }).join(", ");
}

function limpiarFila(fila) {
    return Object.fromEntries(Object.entries(fila).map(function([clave, valor]) {
        return [clave.trim(), valor];
    }));
}

function separarPO(valor) {
    return texto(valor).split("\n").map(function(item) {
        return item.trim();
    }).filter(function(item) {
        return item && !item.includes("…");
    });
}

function valoresUnicos(valores) {
    return [...new Set(valores.map(texto).filter(Boolean))].sort();
}

function prepararDatos(workbook) {
    const hojaNormalizada = workbook.Sheets["FMR_Normalizado"];
    if (!workbook.Sheets[workbook.SheetNames[0]] || !hojaNormalizada) {
        throw new Error("El archivo debe contener la hoja FMR_Normalizado.");
    }

    estado.normalizado = XLSX.utils.sheet_to_json(hojaNormalizada, {
        defval: null,
        cellDates: true
    }).map(limpiarFila);
    estado.relaciones = [];
    estado.detalle = [];
    estado.fmrs = [];

    estado.normalizado.forEach(function(fila) {
        const subsistema = texto(fila.SubSistema);
        const fmr = texto(fila.FMR);
        if (!fmr || fmr.includes("…")) {
            return;
        }

        estado.fmrs.push(fmr);
        const pos = separarPO(fila.PO);
        if (subsistema) {
            pos.forEach(function(po) {
                estado.relaciones.push({ subsistema, fmr, po });
            });
        }
        if (pos.length === 0) {
            estado.detalle.push({ ...fila, SubSistema: subsistema, FMR: fmr, PO: "" });
        } else {
            pos.forEach(function(po) {
                estado.detalle.push({ ...fila, SubSistema: subsistema, FMR: fmr, PO: po });
            });
        }
    });

    estado.fmrs = valoresUnicos(estado.fmrs);
    estado.relaciones = estado.relaciones.filter(function(relacion, indice, relaciones) {
        return relaciones.findIndex(function(item) {
            return item.subsistema === relacion.subsistema
                && item.fmr === relacion.fmr
                && item.po === relacion.po;
        }) === indice;
    });
    estado.semanas = valoresUnicos(estado.normalizado.map(function(fila) {
        return fechaTexto(fila["3week"]);
    }));
}

function aplicarColor(elemento, color) {
    elemento.classList.remove("estado-yellow", "estado-green", "estado-red", "estado-orange", "estado-white", "estado-red-intense");
    elemento.classList.add(`estado-${color || "red"}`);
}

function obtenerEstadoPO(fmr, po, subsistema) {
    const datos = estado.detalle.filter(function(fila) {
        return fila.FMR === fmr && fila.PO === po
            && (!subsistema || fila.SubSistema === subsistema);
    });
    if (datos.length === 0) {
        return { color: "red", estado: "sin_eta" };
    }

    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    let tieneETA = false;
    let tieneETAFutura = false;
    let tieneOnSite = false;

    datos.forEach(function(fila) {
        texto(fila.ETA).replace(/[;\n]/g, ",").split(",").filter(Boolean).forEach(function(valor) {
            const fechaETA = fecha(valor);
            if (fechaETA) {
                tieneETA = true;
                tieneETAFutura = tieneETAFutura || fechaETA > hoy;
            }
        });
        tieneOnSite = tieneOnSite || Boolean(texto(fila["ON SITE"]));
    });

    if (tieneETAFutura) {
        return { color: "yellow", estado: "futura" };
    }
    if (!tieneETA) {
        return { color: "red", estado: "sin_eta" };
    }
    return tieneOnSite
        ? { color: "green", estado: "pasada_on_site" }
        : { color: "yellow", estado: "eta_sin_on_site" };
}

function obtenerDatosFMR(fmr, subsistema = "") {
    const filasFMR = estado.detalle.filter(function(fila) {
        return fila.FMR === fmr && (!subsistema || fila.SubSistema === subsistema);
    });
    const pos = valoresUnicos(filasFMR.map(function(fila) { return fila.PO; })).map(function(po) {
        return { po, ...obtenerEstadoPO(fmr, po, subsistema) };
    });
    const colores = [...new Set(pos.map(function(item) { return item.color; }))];
    const procurement = filasFMR.map(function(fila) {
        return {
            po: texto(fila.PO),
            status: texto(fila.STATUS),
            buyer: texto(fila.BUYER),
            expeditor: texto(fila.EXPEDITOR),
            eta: formatoETA(fila.ETA),
            on_site: texto(fila["ON SITE"]),
            entregado: texto(fila.ENTREGADO),
            descripcion: texto(fila.DESCRIPCION),
            comentarios: texto(fila.COMENTARIOS)
        };
    });

    return {
        encontrado: estado.fmrs.includes(fmr),
        fmr,
        subsistemas: valoresUnicos(estado.normalizado.filter(function(fila) {
            return texto(fila.FMR) === fmr;
        }).map(function(fila) { return fila.SubSistema; })),
        pos,
        total_po: pos.length,
        descripcion: texto(procurement.find(function(item) { return item.descripcion; })?.descripcion),
        procurement,
        color: colores.length >= 2 ? "orange" : (colores[0] || "red"),
        estado: colores.length >= 2 ? "combinada" : (colores[0] === "yellow" ? "futura" : colores[0] === "green" ? "pasada_on_site" : "sin_eta")
    };
}

function obtenerColorFMR(datos) {
    const hoy = new Date();
    hoy.setHours(0, 0, 0, 0);
    const diasFuturos = (datos.procurement || []).flatMap(function(registro) {
        return texto(registro.eta).replace(/[;\n]/g, ",").split(",").map(fecha).filter(function(valorFecha) {
            return valorFecha && valorFecha > hoy;
        }).map(function(valorFecha) {
            const fechaUTC = Date.UTC(valorFecha.getFullYear(), valorFecha.getMonth(), valorFecha.getDate());
            const hoyUTC = Date.UTC(hoy.getFullYear(), hoy.getMonth(), hoy.getDate());
            return Math.round((fechaUTC - hoyUTC) / 86400000);
        });
    });

    if (diasFuturos.length) {
        return Math.min(...diasFuturos) <= 20 ? "white" : "red-intense";
    }

    const colores = [...new Set((datos.pos || []).map(function(item) { return item.color; }))];
    if (colores.length >= 2) {
        return "orange";
    }
    return colores[0] === "yellow" ? "red" : (colores[0] || "red");
}

function crearBoton(textoBoton, clase, color, alClick) {
    const boton = document.createElement("button");
    boton.type = "button";
    boton.classList.add(clase);
    boton.textContent = textoBoton;
    if (color) {
        aplicarColor(boton, color);
    }
    boton.addEventListener("click", alClick);
    return boton;
}

function cargarTodasLasFMR() {
    listaFMRsGeneral.innerHTML = "";
    estado.fmrs.forEach(function(fmr) {
        const datos = obtenerDatosFMR(fmr);
        listaFMRsGeneral.appendChild(crearBoton(fmr, "fmr-general-item", obtenerColorFMR(datos), function() {
            fmrSearch.value = fmr;
            buscarFMR(fmr);
        }));
    });
}

function cargar3Week() {
    threeweekSelect.innerHTML = "<option value=\"\">Seleccione una fecha</option>";
    estado.semanas.forEach(function(semana) {
        const option = document.createElement("option");
        option.value = semana;
        option.textContent = semana;
        threeweekSelect.appendChild(option);
    });
}

function cargarSubsistemas3Week(fechaSeleccionada) {
    const subsistemas = valoresUnicos(estado.normalizado.filter(function(fila) {
        return fechaTexto(fila["3week"]) === fechaSeleccionada;
    }).map(function(fila) { return fila.SubSistema; }));

    infoSubsistema.style.display = "block";
    listaSubsistemas.innerHTML = "";
    listaFMR.innerHTML = "";
    listaPO.innerHTML = "";
    listaProcurement.innerHTML = "";
    detalleFMR.style.display = "none";
    subsistemaSeleccionadoDiv.style.display = "none";

    if (subsistemas.length === 0) {
        listaSubsistemas.innerHTML = "<p class=\"mensaje-vacio\">No existen SubSistemas para la fecha seleccionada.</p>";
        return;
    }
    subsistemas.forEach(function(subsistema) {
        listaSubsistemas.appendChild(crearBoton(subsistema, "subsistema-item", null, function() {
            seleccionarSubsistema(this, subsistema);
        }));
    });
}

function seleccionarSubsistema(boton, subsistema) {
    subsistemaSeleccionado = subsistema;
    document.querySelectorAll(".subsistema-item").forEach(function(item) { item.classList.remove("selected"); });
    boton.classList.add("selected");
    nombreSubsistema.textContent = subsistema;
    subsistemaSeleccionadoDiv.style.display = "block";
    cargarFMR(subsistema);
}

function cargarFMR(subsistema) {
    listaFMR.innerHTML = "";
    detalleFMR.style.display = "none";
    const fmrs = valoresUnicos(estado.relaciones.filter(function(item) {
        return item.subsistema === subsistema;
    }).map(function(item) { return item.fmr; }));
    if (fmrs.length === 0) {
        listaFMR.innerHTML = "<p class=\"mensaje-vacio\">No existen FMR asociados a este SubSistema.</p>";
        return;
    }
    fmrs.forEach(function(fmr) {
        const datos = obtenerDatosFMR(fmr, subsistema);
        listaFMR.appendChild(crearBoton(fmr, "fmr-item", obtenerColorFMR(datos), function() {
            seleccionarFMR(this, fmr, subsistema);
        }));
    });
}

function seleccionarFMR(boton, fmr, subsistema) {
    fmrSeleccionado = fmr;
    subsistemaSeleccionado = subsistema;
    document.querySelectorAll(".fmr-item").forEach(function(item) { item.classList.remove("selected"); });
    boton.classList.add("selected");
    cargarDetalleFMR(fmr, subsistema);
}

function cargarDetalleFMR(fmr, subsistema = "") {
    mostrarDetalleFMR(obtenerDatosFMR(fmr, subsistema));
}

function mostrarDetalleFMR(datos) {
    detalleFMR.style.display = "block";
    fmrEncontrado.innerHTML = `
        <div class="fmr-summary-container">
            <div class="total-po-card"><div class="summary-label">FMR</div><div class="summary-number">${escaparHTML(datos.fmr)}</div></div>
            <div class="total-po-card"><div class="summary-label">TOTAL PO</div><div class="summary-number">${datos.total_po}</div></div>
            <div class="descripcion-fmr"><div class="summary-label">DESCRIPCIÓN</div><div class="descripcion-fmr-text">${escaparHTML(datos.descripcion) || "Sin descripción"}</div></div>
        </div>`;
    listaPO.innerHTML = "";
    listaProcurement.innerHTML = "";
    if (!datos.pos.length) {
        listaPO.innerHTML = "<p class=\"mensaje-vacio\">Esta FMR no tiene PO asociada.</p>";
        mostrarProcurementPorPO(datos.procurement, "");
        return;
    }
    datos.pos.forEach(function(item) {
        listaPO.appendChild(crearBoton(item.po, "po-item", item.color, function() {
            seleccionarPO(this, item.po, datos.procurement);
        }));
    });
}

function seleccionarPO(boton, po, procurement) {
    poSeleccionada = po;
    document.querySelectorAll(".po-item").forEach(function(item) { item.classList.remove("selected"); });
    boton.classList.add("selected");
    mostrarProcurementPorPO(procurement, po);
}

function mostrarProcurementPorPO(procurement, po) {
    listaProcurement.innerHTML = "";
    const registros = procurement.filter(function(registro) { return registro.po === po; });
    const comentariosFMR = procurement.filter(function(registro) {
        return !registro.po && registro.comentarios;
    });
    if (!registros.length) {
        listaProcurement.innerHTML = "<p class=\"mensaje-vacio\">No hay información de Procurement para esta PO.</p>";
        return;
    }
    registros.forEach(function(registro) {
        const record = document.createElement("div");
        record.classList.add("procurement-record");
        const comentarios = [];
        if (registro.comentarios) {
            comentarios.push(`<div><strong>${registro.po ? "PO" : "FMR"}:</strong> ${escaparHTML(registro.comentarios)}</div>`);
        }
        if (registro.po) {
            comentariosFMR.forEach(function(comentario) {
                comentarios.push(`<div><strong>FMR:</strong> ${escaparHTML(comentario.comentarios)}</div>`);
            });
        }
        record.innerHTML = `
            <div class="procurement-card">
                <div class="procurement-field"><span class="procurement-label">STATUS</span><span class="procurement-value">${escaparHTML(registro.status) || "-"}</span></div>
                <div class="procurement-field"><span class="procurement-label">BUYER</span><span class="procurement-value">${escaparHTML(registro.buyer) || "-"}</span></div>
                <div class="procurement-field"><span class="procurement-label">EXPEDITOR</span><span class="procurement-value">${escaparHTML(registro.expeditor) || "-"}</span></div>
                <div class="procurement-field"><span class="procurement-label">ETA</span><span class="procurement-value">${escaparHTML(registro.eta) || "-"}</span></div>
                <div class="procurement-field"><span class="procurement-label">ON SITE</span><span class="procurement-value">${escaparHTML(registro.on_site) || "-"}</span></div>
                <div class="procurement-field"><span class="procurement-label">ENTREGADO</span><span class="procurement-value">${escaparHTML(registro.entregado) || "-"}</span></div>
                <div class="descripcion-procurement"><span class="procurement-label">DESCRIPCION</span><div class="descripcion-value">${escaparHTML(registro.descripcion) || "Sin descripción"}</div></div>
            </div>
            <aside class="comments-card">
                <span class="procurement-label">COMENTARIOS</span>
                <div class="comments-value">${comentarios.join("") || "Sin comentarios"}</div>
            </aside>`;
        listaProcurement.appendChild(record);
    });
}

function buscarFMR(fmr) {
    fmr = texto(fmr);
    if (!fmr) {
        return;
    }
    const datos = obtenerDatosFMR(fmr);
    resultadoFMRSearch.style.display = "block";
    infoSubsistema.style.display = "none";
    subsistemaSeleccionadoDiv.style.display = "none";
    threeweekSelect.value = "";
    fmrEncontrado.innerHTML = "";
    listaSubsistemasFMR.innerHTML = "";
    if (!datos.encontrado) {
        fmrEncontrado.innerHTML = `<p class="mensaje-error">No se encontró la FMR ${escaparHTML(fmr)}.</p>`;
        return;
    }
    fmrEncontrado.appendChild(crearBoton(datos.fmr, "fmr-item", obtenerColorFMR(datos), function() {
        cargarDetalleFMR(datos.fmr);
    }));
    datos.subsistemas.forEach(function(subsistema) {
        listaSubsistemasFMR.appendChild(crearBoton(subsistema, "subsistema-item", null, function() {
            subsistemaSeleccionado = subsistema;
            fmrSeleccionado = datos.fmr;
            infoSubsistema.style.display = "block";
            listaSubsistemas.innerHTML = "";
            const boton = crearBoton(subsistema, "subsistema-item", null, function() {});
            boton.classList.add("selected");
            listaSubsistemas.appendChild(boton);
            nombreSubsistema.textContent = subsistema;
            subsistemaSeleccionadoDiv.style.display = "block";
            cargarFMR(subsistema);
            cargarDetalleFMR(datos.fmr, subsistema);
        }));
    });
    cargarDetalleFMR(datos.fmr);
}

function cargarListaFMR() {
    fmrOptions.innerHTML = "";
    estado.fmrs.forEach(function(fmr) {
        const option = document.createElement("option");
        option.value = fmr;
        fmrOptions.appendChild(option);
    });
}

function limpiarPantalla() {
    subsistemaSeleccionado = "";
    fmrSeleccionado = "";
    poSeleccionada = "";
    listaSubsistemas.innerHTML = "";
    listaFMR.innerHTML = "";
    listaPO.innerHTML = "";
    listaProcurement.innerHTML = "";
    infoSubsistema.style.display = "none";
    resultadoFMRSearch.style.display = "none";
    detalleFMR.style.display = "none";
}

async function procesarArchivoLocal(contenido, nombre, guardar = true) {
    estadoCarga.textContent = "Leyendo el archivo en el navegador...";
    try {
        const workbook = XLSX.read(contenido, { type: "array", cellDates: true });
        prepararDatos(workbook);
        estado.cargado = true;
        threeweekSelect.disabled = false;
        fmrSearch.disabled = false;
        estadoCarga.textContent = `${nombre} cargado. El análisis se ejecuta en el navegador.`;
        cargar3Week();
        cargarListaFMR();
        cargarTodasLasFMR();
        limpiarPantalla();

        if (guardar) {
            try {
                await guardarWorkbookLocal(contenido, nombre);
            } catch (storageError) {
                console.warn("No se pudo guardar el archivo en el navegador:", storageError);
            }
        }

        return true;
    } catch (error) {
        estado.cargado = false;
        threeweekSelect.disabled = true;
        fmrSearch.disabled = true;
        estadoCarga.textContent = error.message || "No se pudo leer el archivo Excel.";
        return false;
    }
}

archivoExcel.addEventListener("change", async function() {
    const archivo = this.files[0];
    if (archivo) {
        await procesarArchivoLocal(await archivo.arrayBuffer(), archivo.name);
    }
});

async function restaurarWorkbookLocal() {
    try {
        const guardado = await leerWorkbookLocal();
        if (guardado) {
            await procesarArchivoLocal(guardado.buffer, guardado.name, false);
        }
    } catch (error) {
        console.warn("No se pudo restaurar el archivo guardado:", error);
    }
}

threeweekSelect.addEventListener("change", function() {
    threeweekSeleccionado = this.value.trim();
    if (!threeweekSeleccionado) {
        limpiarPantalla();
        return;
    }
    fmrSearch.value = "";
    resultadoFMRSearch.style.display = "none";
    detalleFMR.style.display = "none";
    cargarSubsistemas3Week(threeweekSeleccionado);
});

fmrSearch.addEventListener("keydown", function(event) {
    if (event.key === "Enter") {
        event.preventDefault();
        buscarFMR(this.value);
    }
});

fmrSearch.addEventListener("change", function() {
    buscarFMR(this.value);
});

restaurarWorkbookLocal();

