// ==UserScript==
// @name         PFA Parte Operativo -> Excel Maestro
// @namespace    hernan-automatizacion
// @version      2.2
// @description  Agrega campos extra al Parte Operativo y, al generar el PDF, suma automáticamente filas a un Excel maestro (formato "PARA CARGAR").
// @match        https://partes.pages.dev/sitios_partes/parte_Operativo/*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js
// @updateURL    https://raw.githubusercontent.com/mendosacosmea/partes-operativos-pfa/main/parte-operativo-a-excel.user.js
// @downloadURL  https://raw.githubusercontent.com/mendosacosmea/partes-operativos-pfa/main/parte-operativo-a-excel.user.js
// ==/UserScript==

(function () {
    'use strict';

    // =========================================================
    // 0. ENCABEZADOS EXACTOS DE LA HOJA "PARA CARGAR"
    // =========================================================
    const HEADERS = ["SUPERINTENDENCIA Y/O DIRECCION GENERAL", "DEPENDENCIA", "SUMARIO ", "PARTE INFORMATIVO ", "DELITO NRO. 1", "DELITO NRO. 2", "DELITO NRO. 3", "MODALIDAD (DEL DELITO)", "DETALLE (DEL DELITO)", "POSEE CAPTURA", "TIPO CAPTURA", "MOTIVO CAPTURA", "TIPO DE INTERVENCIÓN", "FECHA Y HORA", "PAÍS", "PROVINCIA", "PARTIDO", "LOCALIDAD", "DIRECCIÓN", "DETALLE DIRECCIÓN", "LATITUD", "LONGITUD", "JUZGADO", "SECRETARÍA", "DETALLE", "APELLIDO", "NOMBRE", "ALIAS", "SEXO", "EDAD", "NACIONALIDAD", "DNI", "DOMICILIO", "DETENIDO", "TIPO ARMA", "DETALLE ARMA", "MARCA ARMA", "MODELO ARMA", "CALIBRE ARMA", "NRO SERIE ARMA", "CANTIDAD ARMA", "PED. SEC. ARMA", "OBS. ARMA", "TIPO DROGA", "CANTIDAD DROGA", "MEDICIÓN DROGA", "OBS. DROGA", "TIPO ELEMENTO", "DETALLE ELEMENTO", "CANTIDAD ELEMENTO", "MEDICIÓN ELEMENTO", "AFORO ELEMENTO", "OBS. ELEMENTO", "MARCA VEHÍCULO", "MODELO VEHÍCULO", "DOMINIO VEHÍCULO", "TIPO VEHÍCULO", "DETALLES VEHÍCULO", "APELLIDO VÍCTIMA", "NOMBRE VÍCTIMA", "EDAD VÍCTIMA", "SEXO VÍCTIMA", "NACIONALIDAD VÍCTIMA", "DNI VÍCTIMA", "DOMICILIO VÍCTIMA", "CANTIDAD VÍCTIMAS", "PROCEDIMIENTO"];
    const SHEET_NAME = "PARA CARGAR";
    const VACIO = "-";

    const OPCIONES_INTERVENCION = ["ALLANAMIENTO", "CAPTURA", "PREVENCION VIA PUBLICA", "CONTROL TERMINAL DE OMNIBUS", "CONTROL VEHICULAR", "CONTROL POBLACIONAL", "SECUESTRO", "ORDEN DE PRESENTACION", "OTRO"];
    const OPCIONES_DETALLE_DIRECCION = ["VÍA PÚBLICA", "DOMICILIO", "CASA", "COMERCIO", "OTRO"];

    // =========================================================
    // 1. HELPERS
    // =========================================================
    function val(id) {
        const el = document.getElementById(id);
        if (!el) return "";
        return (el.value || "").toString().trim();
    }
    function orGuion(v) {
        return (v === "" || v === null || v === undefined) ? VACIO : v;
    }

    // =========================================================
    // 1.1 FUNCIÓN PARA EXTRAER LAT/LONG DEL CAMPO COORDENADAS
    // =========================================================
    function extraerLatLong(coordenadas) {
        if (!coordenadas || coordenadas === "") return { lat: "", long: "" };
        
        const partes = coordenadas.split(",");
        if (partes.length < 2) return { lat: "", long: "" };
        
        const lat = partes[0].trim();
        const long = partes[1].trim();
        
        return { lat, long };
    }

    // =========================================================
    // 2. INYECCIÓN DE CAMPOS NUEVOS (solo para el Excel)
    // =========================================================

    // --- 2.1 Campos generales (una vez, antes de la sección de domicilios) ---
    function inyectarCamposGenerales() {
        if (document.getElementById("excel-datos-generales")) return;
        const contenedorDirecciones = document.getElementById("direcciones__contenedor");
        if (!contenedorDirecciones) return;

        const bloque = document.createElement("div");
        bloque.id = "excel-datos-generales";
        bloque.className = "info__contenedor";
        bloque.style.border = "2px dashed #217346";
        bloque.style.padding = "12px";
        bloque.style.marginBottom = "12px";
        bloque.innerHTML = `
            <h4 style="color:#217346;margin-top:0;">Datos adicionales (solo para Excel)</h4>
            <div class="form_row_datos form_datos-triple">
                <div class="form_datos">
                    <label for="excel-tipo-intervencion">Tipo de Intervención Policial</label><br>
                    <select id="excel-tipo-intervencion">
                        <option value="">${VACIO}</option>
                        ${OPCIONES_INTERVENCION.map(o => `<option value="${o}">${o}</option>`).join("")}
                    </select>
                </div>
                <div class="form_datos">
                    <label for="excel-detalle-direccion">Detalle de la Dirección</label><br>
                    <select id="excel-detalle-direccion">
                        <option value="">${VACIO}</option>
                        ${OPCIONES_DETALLE_DIRECCION.map(o => `<option value="${o}">${o}</option>`).join("")}
                    </select>
                </div>
                <div class="form_datos">
                    <label for="excel-juzgado">Juzgado</label><br>
                    <input type="text" id="excel-juzgado">
                </div>
            </div>
            <div class="form_row_datos form_datos-triple">
                <div class="form_datos">
                    <label for="excel-secretaria">Secretaría</label><br>
                    <input type="text" id="excel-secretaria">
                </div>
            </div>
        `;
        contenedorDirecciones.parentNode.insertBefore(bloque, contenedorDirecciones);
    }

    // --- 2.2 Campos por domicilio (Detalles dirección, Alias imputados) ---
    function inyectarCamposDomicilio(numDire) {
        const coordInput = document.getElementById(`dir${numDire}-coordenadas`);
        if (!coordInput) return;
        const contenedorCoord = coordInput.closest(".form_datos-unico");
        if (!contenedorCoord || document.getElementById(`dir${numDire}-detalleDireccion`)) return;

        const bloque = document.createElement("div");
        bloque.className = "form_row_datos form_datos-triple";
        bloque.style.border = "1px dashed #217346";
        bloque.style.padding = "8px";
        bloque.style.marginTop = "6px";
        bloque.innerHTML = `
            <div class="form_datos">
                <label for="dir${numDire}-detalleDireccion">Detalles de la dirección (Excel)</label><br>
                <input type="text" id="dir${numDire}-detalleDireccion" placeholder="Piso, depto, entre calles, etc.">
            </div>
            <div class="form_datos" style="flex-basis:100%;">
                <label for="dir${numDire}-aliasImputados">Alias de imputados (Excel) — separados por coma, mismo orden que los imputados cargados</label><br>
                <input type="text" id="dir${numDire}-aliasImputados" placeholder="Ej: EL FLACO, PIPA">
            </div>
        `;
        contenedorCoord.parentNode.insertBefore(bloque, contenedorCoord.nextSibling);
    }

    // --- 2.3 Enganchar creación de domicilios / generales sin tocar el sistema original ---
    function iniciarInyeccion() {
        inyectarCamposGenerales();

        const btnAgrDireccion = document.getElementById("btn-direccion");
        if (btnAgrDireccion) {
            btnAgrDireccion.addEventListener("click", () => {
                setTimeout(() => {
                    const direcciones = document.querySelectorAll('[id^="contenedor__dir"]');
                    const ultima = direcciones[direcciones.length - 1];
                    if (ultima) inyectarCamposDomicilio(ultima.dataset.dir);
                }, 0);
            });
        }
    }

    // =========================================================
    // 3. RECOLECCIÓN DE DATOS -> FILAS DEL EXCEL
    // =========================================================
    function contarFilas(numDire, elem) {
        return document.querySelectorAll(`[id^="dir${numDire}-fila-${elem}"]`).length;
    }
    function obtenerDirecciones() {
        return Array.from(document.querySelectorAll('[id^="contenedor__dir"]'))
            .map(el => el.dataset.dir).filter(Boolean);
    }

    function datosGenerales() {
        const codigo = val("dependencia-codigo") || "SIN-CODIGO";
        const numParte = val("num-PI") || "S-N";
        const anio = new Date().getFullYear();
        return {
            superintendencia: val("superintendencia-nombre"),
            dependencia: val("dependencia-nombre"),
            sumario: val("sumario"),
            parteInformativo: numParte,
            delito1: val("nombre-delito1"),
            delito2: val("nombre-delito2"),
            delito3: val("nombre-delito3"),
            modalidad: val("modalidad-delito1"),
            detalleDelito: val("detalle__input-delito1"),
            tipoIntervencion: val("excel-tipo-intervencion"),
            fecha: val("fechayhora"),
            juzgado: val("excel-juzgado"),
            secretaria: val("excel-secretaria"),
            detalle: val("excel-detalle-direccion"),
            procedimiento: `${codigo}-PO-${numParte}-${anio}`
        };
    }

    function datosImputado(numDire, i, aliasArray) {
        if (i < 1) return null;
        const pref = `dir${numDire}-imputado${i}-`;
        if (!document.getElementById(pref + "nombres")) return null;
        const pedidoCap = val(pref + "pedidoCap");
        const poseeCaptura = pedidoCap.toUpperCase().startsWith("SI") ? "SI" : "NO";
        const capturaTipo = pedidoCap.includes("Nacional") ? "NACIONAL" : (pedidoCap.includes("Internacional") ? "INTERNACIONAL" : VACIO);
        const situacionProc = val(pref + "situacionProc");
        return {
            apellido: val(pref + "apellidos"),
            nombre: val(pref + "nombres"),
            alias: aliasArray[i - 1] || "",
            sexo: val(pref + "genero"),
            edad: val(pref + "edad"),
            nacionalidad: val(pref + "nacionalidad"),
            dni: val(pref + "dni"),
            domicilio: val(pref + "domicilio"),
            detenido: situacionProc === "DETENIDO" ? "SI" : "NO",
            poseeCaptura, capturaTipo,
            motivoCaptura: val(pref + "motivo")
        };
    }
    function datosArma(numDire, i) {
        const pref = `dir${numDire}-arma${i}-`;
        if (!document.getElementById(pref + "tipo")) return null;
        return {
            tipo: val(pref + "tipo"), detalle: val(pref + "detalle"), marca: val(pref + "marca"),
            modelo: val(pref + "modelo"), calibre: val(pref + "calibre"), num: val(pref + "num"),
            cantidad: val(pref + "cantidad"), pedSec: val(pref + "pedSec"), observaciones: val(pref + "observaciones")
        };
    }
    function datosDroga(numDire, i) {
        const pref = `dir${numDire}-droga${i}-`;
        if (!document.getElementById(pref + "tipo")) return null;
        return {
            tipo: val(pref + "tipo"), cantidad: val(pref + "cantidad"),
            medicion: val(pref + "medicion"), observaciones: val(pref + "observaciones")
        };
    }
    function datosElemento(numDire, i) {
        const pref = `dir${numDire}-elemento${i}-`;
        if (!document.getElementById(pref + "tipo")) return null;
        return {
            tipo: val(pref + "tipo"), detalle: val(pref + "subtipo"), cantidad: val(pref + "cantidad"),
            medicion: val(pref + "medicion"), aforo: val(pref + "aforo"), observaciones: val(pref + "observaciones")
        };
    }
    function datosVehiculo(numDire, i) {
        const pref = `dir${numDire}-vehiculo${i}-`;
        if (!document.getElementById(pref + "marca")) return null;
        return {
            marca: val(pref + "marca"), modelo: val(pref + "modelo"), dominio: val(pref + "dominio"),
            tipo: val(pref + "tipo"), detalles: val(pref + "detalles")
        };
    }
    function datosVictima(numDire, i) {
        const pref = `dir${numDire}-victima${i}-`;
        if (!document.getElementById(pref + "nombres")) return null;
        return {
            apellido: val(pref + "apellidos"), nombre: val(pref + "nombres"), edad: val(pref + "edad"),
            sexo: val(pref + "genero"), nacionalidad: val(pref + "nacionalidad"), dni: val(pref + "dni"),
            domicilio: val(pref + "domicilio")
        };
    }

    function filasDelDomicilio(numDire, gen) {
        const pcia = val(`dir${numDire}-pcia`).toUpperCase();
        const esInternacional = pcia === "INTERNACIONAL";
        const pais = esInternacional ? val(`dir${numDire}-localidad`) : "ARGENTINA";

        // Extraer latitud y longitud del campo coordenadas
        const coordenadas = val(`dir${numDire}-coordenadas`);
        const { lat, long } = extraerLatLong(coordenadas);

        const base = {
            pais,
            provincia: esInternacional ? "" : pcia,
            partido: esInternacional ? "" : val(`dir${numDire}-localidad`),
            localidad: val(`dir${numDire}-loc`),
            direccion: val(`dir${numDire}-calle`),
            detalleDireccion: val(`dir${numDire}-detalleDireccion`),
            latitud: lat,
            longitud: long
        };

        const aliasArray = val(`dir${numDire}-aliasImputados`).split(",").map(s => s.trim());

        const nArmas = contarFilas(numDire, "arma");
        const nDrogas = contarFilas(numDire, "droga");
        const nElementos = contarFilas(numDire, "elemento");
        const nImputados = contarFilas(numDire, "imputado");
        const nVehiculos = contarFilas(numDire, "vehiculo");
        const nVictimas = contarFilas(numDire, "victima");
        const totalFilas = Math.max(nArmas, nDrogas, nElementos, nImputados, nVehiculos, nVictimas, 1);

        const filas = [];
        for (let i = 1; i <= totalFilas; i++) {
            const imp = datosImputado(numDire, i, aliasArray);
            const arma = datosArma(numDire, i);
            const droga = datosDroga(numDire, i);
            const elem = datosElemento(numDire, i);
            const veh = datosVehiculo(numDire, i);
            const vic = datosVictima(numDire, i);

            const fila = [
                gen.superintendencia, gen.dependencia, gen.sumario, gen.parteInformativo,
                gen.delito1, gen.delito2, gen.delito3, gen.modalidad, gen.detalleDelito,
                imp ? imp.poseeCaptura : VACIO, imp ? imp.capturaTipo : VACIO, imp ? imp.motivoCaptura : VACIO,
                gen.tipoIntervencion, gen.fecha,
                base.pais, base.provincia, base.partido, base.localidad, base.direccion, base.detalleDireccion,
                base.latitud, base.longitud, gen.juzgado, gen.secretaria,
                gen.detalle,
                imp ? imp.apellido : "", imp ? imp.nombre : "", imp ? imp.alias : "", imp ? imp.sexo : "",
                imp ? imp.edad : "", imp ? imp.nacionalidad : "", imp ? imp.dni : "", imp ? imp.domicilio : "",
                imp ? imp.detenido : VACIO,
                arma ? arma.tipo : "", arma ? arma.detalle : "", arma ? arma.marca : "", arma ? arma.modelo : "",
                arma ? arma.calibre : "", arma ? arma.num : "", arma ? arma.cantidad : "", arma ? arma.pedSec : "",
                arma ? arma.observaciones : "",
                droga ? droga.tipo : "", droga ? droga.cantidad : "", droga ? droga.medicion : "", droga ? droga.observaciones : "",
                elem ? elem.tipo : "", elem ? elem.detalle : "", elem ? elem.cantidad : "", elem ? elem.medicion : "",
                elem ? elem.aforo : "", elem ? elem.observaciones : "",
                veh ? veh.marca : "", veh ? veh.modelo : "", veh ? veh.dominio : "", veh ? veh.tipo : "", veh ? veh.detalles : "",
                vic ? vic.apellido : "", vic ? vic.nombre : "", vic ? vic.edad : "", vic ? vic.sexo : "",
                vic ? vic.nacionalidad : "", vic ? vic.dni : "", vic ? vic.domicilio : "",
                String(nVictimas), gen.procedimiento
            ].map(orGuion);

            filas.push(fila);
        }
        return filas;
    }

    function recolectarTodasLasFilas() {
        const gen = datosGenerales();
        let filas = [];
        obtenerDirecciones().forEach(numDire => {
            filas = filas.concat(filasDelDomicilio(numDire, gen));
        });
        return filas;
    }

    // =========================================================
    // 4. ARCHIVO MAESTRO (File System Access API + IndexedDB)
    // =========================================================
    const DB_NAME = "parte-operativo-excel";
    const STORE_NAME = "handles";
    const HANDLE_KEY = "archivoMaestro";

    function abrirDB() {
        return new Promise((resolve, reject) => {
            const req = indexedDB.open(DB_NAME, 1);
            req.onupgradeneeded = () => req.result.createObjectStore(STORE_NAME);
            req.onsuccess = () => resolve(req.result);
            req.onerror = () => reject(req.error);
        });
    }
    async function guardarHandle(handle) {
        const db = await abrirDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, "readwrite");
            tx.objectStore(STORE_NAME).put(handle, HANDLE_KEY);
            tx.oncomplete = () => resolve();
            tx.onerror = () => reject(tx.error);
        });
    }
    async function leerHandle() {
        const db = await abrirDB();
        return new Promise((resolve, reject) => {
            const tx = db.transaction(STORE_NAME, "readonly");
            const req = tx.objectStore(STORE_NAME).get(HANDLE_KEY);
            req.onsuccess = () => resolve(req.result || null);
            req.onerror = () => reject(req.error);
        });
    }

    async function elegirArchivoMaestro() {
        const [handle] = await window.showOpenFilePicker({
            types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }],
            excludeAcceptAllOption: false,
            multiple: false
        });
        await guardarHandle(handle);
        alert("Archivo maestro seleccionado: " + handle.name);
    }

    async function crearArchivoMaestro() {
        const handle = await window.showSaveFilePicker({
            suggestedName: "partes-operativos-maestro.xlsx",
            types: [{ description: "Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"] } }]
        });
        const wb = XLSX.utils.book_new();
        const ws = XLSX.utils.aoa_to_sheet([HEADERS]);
        XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
        const buffer = XLSX.write(wb, { type: "array", bookType: "xlsx" });
        const writable = await handle.createWritable();
        await writable.write(buffer);
        await writable.close();
        await guardarHandle(handle);
        alert("Archivo maestro creado: " + handle.name);
    }

    async function agregarFilasAlMaestro(filas) {
        const handle = await leerHandle();
        if (!handle) {
            alert("Todavía no seleccionaste ni creaste el archivo maestro. Usá los botones correspondientes primero.");
            return false;
        }
        const permiso = await handle.requestPermission({ mode: "readwrite" });
        if (permiso !== "granted") {
            alert("No se otorgó permiso para escribir en el archivo maestro.");
            return false;
        }

        const file = await handle.getFile();
        const buffer = await file.arrayBuffer();
        const wb = XLSX.read(buffer, { type: "array" });

        let ws = wb.Sheets[SHEET_NAME];
        if (!ws) {
            ws = XLSX.utils.aoa_to_sheet([HEADERS]);
            XLSX.utils.book_append_sheet(wb, ws, SHEET_NAME);
        }

        const rango = XLSX.utils.decode_range(ws['!ref'] || "A1:A1");
        let filaSiguiente = rango.e.r + 1;
        if (filaSiguiente < 1) filaSiguiente = 1;

        XLSX.utils.sheet_add_aoa(ws, filas, { origin: { r: filaSiguiente, c: 0 } });

        const nuevoBuffer = XLSX.write(wb, { type: "array", bookType: "xlsx" });
        const writable = await handle.createWritable();
        await writable.write(nuevoBuffer);
        await writable.close();
        return true;
    }

    // =========================================================
    // 5. BOTONES EN LA PÁGINA
    // =========================================================
    function agregarBotones() {
        const btnPrint = document.getElementById("btn-form-print");
        if (!btnPrint || document.getElementById("excel-botones")) return;

        const cont = document.createElement("div");
        cont.id = "excel-botones";
        cont.style.display = "inline-block";
        cont.style.marginLeft = "10px";

        const btnElegir = document.createElement("button");
        btnElegir.type = "button";
        btnElegir.textContent = "📂 Elegir Excel maestro";
        btnElegir.style.marginRight = "6px";
        btnElegir.addEventListener("click", elegirArchivoMaestro);

        const btnCrear = document.createElement("button");
        btnCrear.type = "button";
        btnCrear.textContent = "🆕 Crear Excel maestro";
        btnCrear.addEventListener("click", crearArchivoMaestro);

        cont.appendChild(btnElegir);
        cont.appendChild(btnCrear);
        btnPrint.parentNode.insertBefore(cont, btnPrint.nextSibling);

        btnPrint.addEventListener("click", async () => {
            const filas = recolectarTodasLasFilas();
            const ok = await agregarFilasAlMaestro(filas);
            if (ok) console.log(`Se agregaron ${filas.length} fila(s) al Excel maestro.`);
        });
    }

    // =========================================================
    // 6. INICIO
    // =========================================================
    const intervalo = setInterval(() => {
        if (document.getElementById("btn-form-print") && document.getElementById("direcciones__contenedor")) {
            iniciarInyeccion();
            agregarBotones();
            clearInterval(intervalo);
        }
    }, 500);
})();
