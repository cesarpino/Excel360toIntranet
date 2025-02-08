// ==UserScript==
// @name         Viaje con Excel
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Rellena viajes a partir el contenido de un excel compartido para usuarios autorizados.
// @author       César del Pino
// @match        http://viajes.cdti.es/*
// @match        https://www.google.com/search?q=redirige_viajes*
// @grant        GM_xmlhttpRequest
// @grant        GM_listValues
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addStyle
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=cdti.es
// ==/UserScript==

(function() {
    const this_page_url = window.location.href;
    const WEBAPP_URI = get_match_from_UserScript("http://");
    const REDIRECT_URI = get_match_from_UserScript("https://");
    if (this_page_url.startsWith(REDIRECT_URI)) {
        // Se ejecuta solo en dominio en google.com para redirigir el token a viajes.cdti.es
        let newUrl = `${WEBAPP_URI}${this_page_url.slice(REDIRECT_URI.length)}`;
        alert("Redirige autorización a "+newUrl);
        window.location.href = newUrl;
        console.error("no debe pasar por aqui");
        return;
    }

    const TENANT_ID = getAuthParameter("TENANT_ID"); // or replace with organization TENANT_ID
    const CLIENT_ID = getAuthParameter("CLIENT_ID"); // or replace with clientId of app configured in Azure. ex "Acceso a OneDrive desde viajes.cdti.es" en Azure
    // you must create a Azure app, and get CLIENT_ID and configure same REDIRECT_URI
    // REDIRECT_URI must be configured also in azure associated with client_id.
    const AUTH_URL_BASE = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
    const SCOPES = ['Files.Read', 'Files.Read.All', 'Calendars.Read'].join(' ');
    const AUTH_URL = `${AUTH_URL_BASE}?client_id=${CLIENT_ID}&response_type=token&redirect_uri=${REDIRECT_URI}&scope=${SCOPES}&response_mode=fragment`;

    function get_match_from_UserScript(http_or_https) {
        let config_uri=GM_info.script.matches
        .find(match => match.startsWith(http_or_https))
        ?.replace(/\*$/, '');
        return config_uri;
    }
    function getAuthParameter(param_name) {
        const value=GM_getValue(param_name, undefined);
        // console.log("AuthParameter value",param_name, value);
        if (!value) {
            alert(`Falta el parametro autorización ${param_name}, solicita al autor la url de autorizacion`);
            console.error("falta parametro de configuracion ",param_name);
        }
        return value;
    }

    let accessToken = getConfig('accessToken', "");
    function getConfig(key, defaultValue) {
        let value = GM_getValue(key, undefined); // recupera clave, valor almacenada en browser
        if (value === undefined) { // Si no existe, la creamos
            GM_setValue(key, defaultValue);
            return defaultValue;
        }
        return value;
    }
    function checkForTokenInURL() {
        console.log("checkForTokenInURL");
        const hash = window.location.hash.substring(1);
        const params = new URLSearchParams(hash);
        const token = params.get('access_token');

        if (token) {
            accessToken = token;
            GM_setValue('accessToken', token); // Guardar el token para futuras solicitudes
            // window.location.hash = ''; // Limpiar el hash de la URL
            fetchCalendar();
            fetchExcelData();
        }
    }
    function invalidateToken(){
        accessToken=null;
        GM_setValue('accessToken', accessToken);
    }
    function authenticate() {
        // cambia de pagina para solicitar autorización, y aprovechar que el usuario está logado ya en microsoft.
        window.location.href=AUTH_URL;
        // GM_openInTab(authUrl, true);  para abrir la app de autenticación en otra pantalla
    }
    function checkAuth() {
        console.log("checkAuth");
        if (!accessToken) {
            console.log("autentificar");
            authenticate();
            console.error("no debe pasar por aqui");
        }
        console.log("autentificado! access token ", accessToken);
    }

    const AUTH_WEBAPP_PARAMETERS={
        "group_name":"auth_webapp",
        "parameter_names":["TENANT_ID","CLIENT_ID"]
    };
    const EXCEL_PARAMETERS={
        "group_name":"excel_config",
        "parameter_names":["EXCEL_FILE_ID","SHEET_NAME","SHEET_RANGE"]
    };
    const AUTH_PARAMETERS={
        "group_name":"auth_config",
        "parameter_names":["TENANT_ID","CLIENT_ID","EXCEL_FILE_ID","SHEET_NAME","SHEET_RANGE"]
    };
    const USER_PARAMETERS={
        "group_name":"user_config",
        "parameter_names":["fuerza técnico"]
    };
    console.log("get config url",getConfigURL(AUTH_PARAMETERS));
    console.log("get config url",getConfigURL(USER_PARAMETERS));
    setConfigFromURL(AUTH_WEBAPP_PARAMETERS);
    setConfigFromURL(USER_PARAMETERS);

    const EXCEL_FILE_ID = getAuthParameter('EXCEL_FILE_ID'); // ID del archivo de Excel con proyectos de todos los tecnicos
    const SHEET_NAME = getAuthParameter('SHEET_NAME'); // hoja preparada con datos para este script
    const SHEET_RANGE = getConfig('SHEET_RANGE','A1:AE10000'); // Reemplaza el rango que deseas consultar

    // Verificar si hay un token en la URL (flujo implícito)
    function fetchMicrosoftGraph(url) {
        console.log("fetchMicrosoftGraph",url);
        function avisa_error(errorReported,texto_aviso, alertar=true){
            let texto_alerta=`${texto_aviso}\nCode: ${errorReported.code}\nMessage: ${errorReported.message}`;
            if (errorReported.innerError) {
                texto_alerta+=`\nCódigo interno: ${errorReported.innerError.code}`;
            }
            if (alertar) {
                alert(texto_alerta);
            }

        }
        function handleGraphError(response, reject) {
            try {
                let errorData = JSON.parse(response.responseText);
                let errorReported=errorData.error;
                if (errorReported) {
                    console.error("Código de error:", errorReported.code);
                    console.error("Mensaje de error:", errorReported.message);

                    if (errorReported.innerError) {
                        console.error("Código interno:", errorReported.innerError);
                        console.error("Código interno:", errorReported.innerError.code);
                        console.error("ID de solicitud:", errorReported.innerError["request-id"]);
                        console.error("Fecha:", errorReported.innerError.date);
                    }

                    // Manejo de errores específicos
                    switch (errorReported.code) {
                        case "InvalidAuthenticationToken":
                        case "unauthorized":
                            avisa_error(errorReported,"El token de acceso no es válido o ha expirado.");
                            invalidateToken();
                            checkAuth(); // Reauthenticate if the token is invalid
                            reject(new Error(`Error fetching data: ${response.statusText}`));
                            break;
                        case "invalidRequest":
                            avisa_error(errorReported,"invalidRequest. solicitud incorrecta.");
                            break;
                        case "badRequest":
                            if (errorReported.innerError.code === "invalidRange") {
                                avisa_error(errorReported,"Estás intentando subir datos en un rango inválido o solapado.");
                            } else {
                                avisa_error(errorReported,"badRequest. solicitud incorrecta.");
                            }
                            break;
                        case "forbidden":
                            avisa_error(errorReported,"No tienes permisos suficientes.");
                            break;
                        case "notFound":
                            avisa_error(errorReported,"El archivo o recurso no existe.");
                            break;
                        default:
                            avisa_error(errorReported,"Error desconocido.");
                    }
                }
            } catch (e) {
                console.error("Error al procesar la respuesta JSON:", e);
            }
        }

        return new Promise((resolve,reject)=>{
            GM_xmlhttpRequest({
                method: 'GET',
                url: url,
                headers: {
                    'Authorization': `Bearer ${accessToken}`
                },
                onload: function(response) {
                    if (response.status >= 200 && response.status < 300) {
                        console.log("Graph ok:", response);
                        resolve(response);
                    } else {
                        handleGraphError(response, reject);
                    }
                },
                onerror: function(error) {
                    reject(new Error(`Error connecting graph: ${error}`));
                }
            });
        });
    };
    function fetchExcelData() {
        console.log("fechExcel");
        const url = `https://graph.microsoft.com/v1.0/me/drive/items/${EXCEL_FILE_ID}/workbook/worksheets/${SHEET_NAME}/range(address='${SHEET_RANGE}')`;
        fetchMicrosoftGraph(url)
            .then(response=>{
            console.log('Data received:', response);
            const data = JSON.parse(response.responseText);
            let rows = data.values;
            let col_names = rows.shift();
            let sheet = {
                "rows":rows,
                "col_names":col_names
            }
            console.log("hoja", sheet);
            // Filtrar las filas correspondientes al técnico
            let tecnico = getConfig("fuerza técnico", "cdgo");
            if (! tecnico) {
                tecnico=$('#ctl00_SheetContentPlaceHolder_UCSolicitante_ddlSolicitante').find('option:selected').val();
            }
            console.log("tecnico", tecnico);
            const index_tecnico=sheet.col_names.indexOf("tecnico");
            const filteredRows = sheet.rows.filter((row) => row[index_tecnico] === tecnico); //.includes(tecnico));
            desplegable={
                "rows":filteredRows,
                "col_names":col_names
            };
            InsertaBuscador();
            console.log('Desplegable:', desplegable);
        })
            .catch(error=>{
            console.error('Excel Data:', error);
        });
    };
    function fetchCalendar(){
        let fechaActual = new Date().toISOString();
        // const apiUrl = "https://graph.microsoft.com/v1.0/me/events?$filter=contains(subject,'idi-20241234')";
        const apiUrl = `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${fechaActual}' and contains(subject,'viaje')`;

        fetchMicrosoftGraph(apiUrl)
            .then(response=>{
            console.log('Calendar Data received:', response);
            const data = JSON.parse(response.responseText);
            if (data.value.length > 0) {
                data.value.forEach(evento => {
                    let titulo = evento.subject;
                    let fechaInicio = evento.start.dateTime;
                    let zonaHoraria = evento.start.timeZone;

                    console.log(`Evento: ${titulo}`);
                    console.log(`Fecha: ${fechaInicio} (Zona horaria: ${zonaHoraria})`);
                });
            } else {
                console.log("No se encontró el evento.");
            }
        })
    };

    let desplegable=[];
    function InsertaActividades_optionValue (optionValue) {
        // $('#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_ddlOrganismosFinanciadores option[value="FE"]').prop('selected', true);
        if (!optionValue) {
            alert("Actividad vacía. Comprueba el excel de actividades.");
            return;
        }

        $('#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_ddlActividades option:contains("SEGUIMIENTO")').prop('selected', true);
        $('#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_txtPorcentajeActividad').attr('value',100);
        if ($('#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_gvActividadesProductos').length === 0) {
            // no añado actividad si ya existe una
            $('#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_btnAnyadirActividad').trigger('click');
        }

        // https://stackoverflow.com/questions/64712065/how-to-insert-mutationobserver-into-jquery-code
        // Now we watch for new elements
        const observerOrganosFinanciadores = new MutationObserver(() => {
            const detActividades='#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_detActividadesProductos_';
            if ($(detActividades+'btnAnyadirProducto').length) {
                console.log('Aparecio el boton de añadir producto');
                //$(detActividades+'ddlProductos option:contains("I+D+i COFINANCIADOS POPE 2")').prop('selected', true);
                $(detActividades+'ddlProductos option[value="'+optionValue+'"]').prop('selected', true);
                // TODO comprobar si optionValue se ha seleccionado
                $(detActividades+'txtPorcentajeProducto').attr('value',100);
                $(detActividades+'btnAnyadirProducto').trigger('click');
            }
            if ($(detActividades+'btnAnyadirOrgPresupCentrosCostes').length) {
                $(detActividades+'txtPorcentajeOrgPresup').attr('value',100);
                $(detActividades+'txtPorcentajeCentroCoste').attr('value',100);
                $(detActividades+'btnAnyadirOrgPresupCentrosCostes').trigger('click');
            }
            if ($(detActividades+'gvProductosOrgPresupCCostes_ctl02_btnEditarFila').length) {
                $(detActividades+'btnGuardar').trigger('click');
                observerOrganosFinanciadores.disconnect();
                beep();
            }
        });
        observerOrganosFinanciadores.observe(document.getElementById('ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_updOtrosDatos'), {
            childList: true,
            subtree: true
        });
    }
    function InsertaDestino(provincia,poblacion) {
        const destino=(poblacion==provincia
                       ? provincia
                       : `${poblacion}, ${provincia}`);
        console.log("InsertaDestino",destino);

        let tipo_dieta="N";
        if (provincia=="MADRID") {
            tipo_dieta="M";
            if (poblacion == "MADRID") {
                // debe ser desplazamiento
                if ($('#ctl00_SheetContentPlaceHolder_UCDesplazamientos_txtDestino').length==0) {
                    alert("Esto debe ser un desplazamiento a Madrid, no un viaje.");
                    throw "error";
                }
            } else {
                // puede ser viaje
                if ($('#ctl00_SheetContentPlaceHolder_UCDGViajes_txtDestino').length!=0) {
                    alert("Si "+poblacion+" está en el municipio de Madrid, debe pedir desplazamiento, no viaje.");
                }
            }
        }
        $('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_ddlTipoDieta option[value="'+tipo_dieta+'"]').prop('selected', true);
        $(['#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtDestino',
           "#ctl00_SheetContentPlaceHolder_UCDesplazamientos_txtDestino",
           "#ctl00_SheetContentPlaceHolder_UCDGViajes_txtDestino"].join(",")).val(function(index, value) {
            console.log("seleccionado destino",index,value);
            return destino;
        });
    };
    function InsertaOrigen() {
        $(['#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtOrigen'].join(",")).val(function(index, value) {
            console.log("seleccionado origen",index,value);
            return "Madrid";
        });
    };
    function InsertaMotivo(motivo){
        console.log("inserta motivo",motivo);
        $(["#ctl00_SheetContentPlaceHolder_UCDesplazamientos_txtMotivo",
           "#ctl00_SheetContentPlaceHolder_UCDGViajes_txtMotivo"].join(",")).val(function(index, value) {
            console.log("motivo",index,value);
            return value + motivo +"\n";
        });
    };
    function InsertaProyectoSeleccionado(texto_desplegable){
        const index_texto_desplegable=desplegable.col_names.indexOf("texto_desplegable");
        const filteredRows = desplegable.rows.filter((row) => row[index_texto_desplegable] === texto_desplegable); //.includes(tecnico));
        if (filteredRows.length != 1) {
            alert("he encontrado más de un elemento correspondiente a ",texto_desplegable);
            return;
        }
        const proyecto_seleccionado = Object.fromEntries(desplegable.col_names.map((col, i) => [col, filteredRows[0][i]]));
        console.log("proyecto seleccionado",proyecto_seleccionado);
        InsertaMotivo(proyecto_seleccionado.motivo);
        InsertaOrigen();
        InsertaDestino(proyecto_seleccionado.provincia_desa,proyecto_seleccionado.localidad_desa);
        InsertaActividades_optionValue(proyecto_seleccionado.tipo_a_desplegable_id_desplegable)
    };
    function beep() {
        (new
         Audio(
            "data:audio/wav;base64,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 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////VEFHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAU291bmRib3kuZGUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMjAwNGh0dHA6Ly93d3cuc291bmRib3kuZGUAAAAAAAAAACU="
        )).play();
    }
    function desactivaAutocomplete() {
        $('input, textarea, select').attr('autocomplete', 'off');
    }
    function InsertaBuscador(){
        function estilos(){
            // Estilos CSS
            GM_addStyle(`
        #contenedor_buscador{
           position: fixed;
           top: 20px;
           right: 20px;
           display: flex;
           flex-direction: column; /* Cambia a columna */
           gap: 0px; /* Espacio entre los inputs */
           z-index: 9999;
        }
        #buscador_excel, #buscador_calendar {
            width: 400px;
            padding: 5px;
            border: 1px solid #ccc;
            border-radius: 5px;
            font-size: 12px;
            z-index: 9999;
        }
        #lista_sugerencias_excel,
        #lista_sugerencias_calendar{
            top: 50px;
            right: 20px;
            width: 400px;
            background: white;
            border: 1px solid #ccc;
            max-height: 1500px;
            overflow-y: auto;
            display: none;
            z-index: 9999;
        }
        .item {
            padding: 5px;
            cursor: pointer;
        }
        .item:hover {
            background-color: #ddd;
        }
        `);
        }
        function buscador_calendar(){
            if ($("#buscador_calendar").length) {
                return;
            }

            // Crear la caja de búsqueda en la página
            $('#contenedor_buscador').append('<input type="text" id="buscador_calendar" autocomplete="off" placeholder="Letras del proyecto, empresa, provincia, ..">');
            $('#contenedor_buscador').append('<div id="lista_sugerencias_calendar"></div>');


            // Filtrar y mostrar sugerencias
            $("#buscador_calendar").on("input", function() {
                var texto = $(this).val().toLowerCase();
                $("#lista_sugerencias_calendar").empty();

                if (texto.length > 0) {
                    var resultados = desplegable.rows.filter(item => {
                        const todos_los_campos=(''+item);
                        // console.log("todos los campos",todos_los_campos);
                        return todos_los_campos.toLowerCase().includes(texto)
                    });

                    if (resultados.length > 0) {
                        const index_texto_desplegable=desplegable.col_names.indexOf("texto_desplegable");
                        resultados.forEach(function(item) {
                            $("#lista_sugerencias_calendar").append("<div class='item'>" + item[index_texto_desplegable] + "</div>");
                        });
                        $("#lista_sugerencias_calendar").show();
                    } else {
                        $("#lista_sugerencias_calendar").hide();
                    }
                } else {
                    $("#lista_sugerencias_calendar").hide();
                }
            });
            // Evento para seleccionar un ítem de la lista
            $(document).on("click", ".item", function() {
                const texto_desplegable=$(this).text();
                $("#buscador_calendar").val(texto_desplegable);
                $("#lista_sugerencias_calendar").hide();
                InsertaProyectoSeleccionado(texto_desplegable);
            });

            // Ocultar la lista si se hace clic fuera
            $(document).click(function(e) {
                if (!$(e.target).closest("#buscador_calendar, #lista_sugerencias_calendar").length) {
                    $("#lista_sugerencias_calendar").hide();
                }
            });
        }
        function buscador_excel(){
            $('#contenedor_buscador').append('<input type="text" id="buscador_excel" autocomplete="off" placeholder="Letras del proyecto, empresa, provincia, ..">');
            $('#contenedor_buscador').append('<div id="lista_sugerencias_excel"></div>');

            // Filtrar y mostrar sugerencias
            $("#buscador_excel").on("input", function() {
                var texto = $(this).val().toLowerCase();
                $("#lista_sugerencias_excel").empty();

                if (texto.length > 0) {
                    var resultados = desplegable.rows.filter(item => {
                        const todos_los_campos=(''+item);
                        // console.log("todos los campos",todos_los_campos);
                        return todos_los_campos.toLowerCase().includes(texto)
                    });

                    if (resultados.length > 0) {
                        const index_texto_desplegable=desplegable.col_names.indexOf("texto_desplegable");
                        resultados.forEach(function(item) {
                            $("#lista_sugerencias_excel").append("<div class='item'>" + item[index_texto_desplegable] + "</div>");
                        });
                        $("#lista_sugerencias_excel").show();
                    } else {
                        $("#lista_sugerencias_excel").hide();
                    }
                } else {
                    $("#lista_sugerencias_excel").hide();
                }
            });
            // Evento para seleccionar un ítem de la lista
            $(document).on("click", ".item", function() {
                const texto_desplegable=$(this).text();
                $("#buscador_excel").val(texto_desplegable);
                $("#lista_sugerencias_excel").hide();
                InsertaProyectoSeleccionado(texto_desplegable);
            });

            // Ocultar la lista si se hace clic fuera
            $(document).click(function(e) {
                if (!$(e.target).closest("#buscador_excel, #lista_sugerencias_excel").length) {
                    $("#lista_sugerencias_excel").hide();
                }
            });
        };

        if ($("#contenedor_buscador").length) {
            return;
        }
        // Crear la caja de búsqueda en la página
        $('body').append('<div id="contenedor_buscador"></div>');
        estilos();
        buscador_excel();
        //buscador_calendar();
    }
    function gestor_configuracion() {
        let isCollapsed = false; // Estado inicial (expandido)

        // Obtener todas las claves almacenadas
        async function getAllKeys() {
            return await GM_listValues();
        }

        // Obtener todos los valores como un objeto clave-valor
        async function getAllValues() {
            const keys = await getAllKeys();
            let values = {};
            keys.forEach(key => {
                values[key] = GM_getValue(key, ""); // Tratar todo como string
            });
            return values;
        }

        // Guardar un nuevo valor
        function saveValue(key, newValue) {
            GM_setValue(key, newValue);
            alert(`✅ Guardado: ${key} = "${newValue}"`);
        }

        // Alternar visibilidad de los valores
        function toggleConfigPanel() {
            const configBody = document.getElementById("configBody");
            if (isCollapsed) {
                configBody.style.display = "block";
            } else {
                configBody.style.display = "none";
            }
            isCollapsed = !isCollapsed;
        }

        // Crear la interfaz
        async function createUI() {
            const values = await getAllValues();

            // Crear el contenedor flotante
            const container = document.createElement("div");
            container.style.position = "fixed";
            container.style.bottom = "10px";
            container.style.right = "10px";
            container.style.background = "white";
            container.style.padding = "10px";
            container.style.border = "1px solid black";
            container.style.zIndex = "10000";
            container.style.maxHeight = "400px";
            container.style.overflowY = "auto";
            container.style.fontFamily = "Arial, sans-serif";
            container.style.fontSize = "14px";
            container.style.borderRadius = "8px";
            container.style.boxShadow = "2px 2px 10px rgba(0, 0, 0, 0.2)";
            container.style.transition = "all 0.3s ease";

            let html = `
            <div id="configHeader" style="cursor: pointer; background: #007bff; color: white; padding: 5px; text-align: center; font-weight: bold; border-radius: 5px;">
                📌 Editor de Configuración
            </div>
            <div id="configBody" style="margin-top: 10px;">
        `;

            Object.keys(values).forEach(key => {
                html += `
                <label>${key}:</label><br>
                <input type="text" id="input_${key}" value="${values[key]}" style="width: 200px;">
                <button id="save_${key}" style="background: green; color: white; border: none; padding: 2px 5px; margin-left: 5px;">✔️</button>
                <br><br>
            `;
            });

            html += `</div>`;
            container.innerHTML = html;
            document.body.appendChild(container);

            // Evento para colapsar/expandir
            document.getElementById("configHeader").addEventListener("click", toggleConfigPanel);

            // Agregar eventos a cada botón de guardar
            Object.keys(values).forEach(key => {
                document.getElementById(`save_${key}`).addEventListener("click", () => {
                    const newValue = document.getElementById(`input_${key}`).value;
                    saveValue(key, newValue);
                });
            });
            toggleConfigPanel();
        }

        createUI();
    };
    function getConfigURL(config_parameters) {
        // takes config and store in local storage
        const baseURL=WEBAPP_URI;
        const config = {};
        config[config_parameters.group_name]=1;
        config_parameters.parameter_names.forEach(param_name => {
            config[param_name] = GM_getValue(param_name, undefined);
        });
        const queryString = new URLSearchParams(config).toString();
        const url= `${baseURL}?${queryString}`;
        console.log("getConfigURL",config, "url",url);
        return url;
    }
    function setConfigFromURL(config_parameters) {
        console.log("setConfigFromURL");
        const url_params = new URLSearchParams(window.location.search);
        const paramsObj = Object.fromEntries(url_params.entries());
        if ("1"!=url_params.get(config_parameters.group_name)) {
            console.log("no es url de config");
            return;
        }
        config_parameters.parameter_names.map(config_variable_name=>{
            let config_value = url_params.get(config_variable_name);
            GM_setValue(config_variable_name, config_value); // Guardar el token para futuras solicitudes
        });
    }

    gestor_configuracion();
    checkForTokenInURL();
    checkAuth();
    //fetchCalendar();
    fetchExcelData();
    desactivaAutocomplete();

})();