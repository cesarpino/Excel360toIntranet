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
// @grant        GM_deleteValue
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=cdti.es
// ==/UserScript==

(function() {
    'use strict';

    const this_page_url = window.location.href;
    const WEBAPP_URI = get_match_from_UserScript("http://");
    const REDIRECT_URI = get_match_from_UserScript("https://");
    if (this_page_url.startsWith(REDIRECT_URI)) {
        // We are inside REDIRECT_URI, we force redirection to WEBAPP_URI with the token parameter recieved.
        let back_to_webapp_uri = `${WEBAPP_URI}${this_page_url.slice(REDIRECT_URI.length)}`;
        alert("Redirects authorization to \n"+back_to_webapp_uri);
        stop_execution_and_jump_to(back_to_webapp_uri);
    }

    // We are in now an WEBAPP_URI page, and want to get an access_token for accessing microsoft.
    const AUTH_WEBAPP_PARAMETERS={
        "group_name":"auth_webapp",
        "parameter_names":["TENANT_ID","CLIENT_ID"]
    };

    const TENANT_ID = getConfigFromSet(AUTH_WEBAPP_PARAMETERS,"TENANT_ID"); // or replace with organization TENANT_ID
    const CLIENT_ID = getConfigFromSet(AUTH_WEBAPP_PARAMETERS,"CLIENT_ID"); // or replace with clientId of app configured in Azure. ex "Acceso a OneDrive desde viajes.cdti.es" en Azure
    // you must create a Azure app, and get CLIENT_ID and configure same REDIRECT_URI
    // REDIRECT_URI must be configured also in azure associated with client_id.
    const AUTH_URL_BASE = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
    const SCOPES = ['Files.Read', 'Files.Read.All', 'Calendars.Read', 'Mail.Read'].join(' ');
    const AUTH_URL = `${AUTH_URL_BASE}?client_id=${CLIENT_ID}&response_type=token&redirect_uri=${REDIRECT_URI}&scope=${SCOPES}&response_mode=fragment`;

    function get_match_from_UserScript(http_or_https) {
        let config_uri=GM_info.script.matches
        .find(match => match.startsWith(http_or_https))
        ?.replace(/\*$/, ''); // le quito la estrella del final @match
        return config_uri;
    }
    function stop_execution_and_jump_to(newUrl){
        window.location.href = newUrl;
        throw new Error("Stop script after asking for redirection");
    }
    function getConfigFromSet(param_set, param_name) {
        function setConfigFromURL(config_parameters) {
            console.log("setConfigFromURL", config_parameters);

            const url_params = new URLSearchParams(window.location.search);
            const paramsObj = Object.fromEntries(url_params.entries());
            switch (url_params.get(config_parameters.group_name)) {
                case "clear_all_and_set":
                    GM_listValues().forEach(key => {
                        GM_deleteValue(key);
                    });
                    alert("Tampermonkey config cleaned");
                    // no break;
                case "set":
                    config_parameters.parameter_names.map(config_variable_name=>{
                        let config_value = url_params.get(config_variable_name);
                        GM_setValue(config_variable_name, config_value); // Guardar el token para futuras solicitudes
                    });
                    break;
                default:
                    console.log("no es url de config",url_params.get(config_parameters.group_name));
                    break;
            }
        }
        function getConfigURL(config_parameters,set_or_clear) {
            // takes config and store in local storage
            const baseURL=WEBAPP_URI;
            const config = {};
            config[config_parameters.group_name]=set_or_clear;
            config_parameters.parameter_names.forEach(param_name => {
                config[param_name] = GM_getValue(param_name, undefined);
            });
            const queryString = new URLSearchParams(config).toString();
            const url= `${baseURL}?${queryString}`;
            return url;
        }

        setConfigFromURL(param_set);

        const value=GM_getValue(param_name, undefined);
        // console.log("AuthParameter value",param_name, value);
        if (!value) {
            const error_text=`Falta el parametro de configuración ${param_name}.\n`+
                  "Solicita al autor la url de autorizacion de la forma \n"+
                  getConfigURL(param_set,"set")+
                  "\n o de la forma \n"+
                  getConfigURL(param_set,"clear_all_and_set");
            console.error(error_text);
            alert(error_text);
        }
        return value;
    }
    function getConfig(key, defaultValue) {
        let value = GM_getValue(key, undefined); // recupera clave, valor almacenada en browser
        if (value === undefined) { // Si no existe, la creamos
            GM_setValue(key, defaultValue);
            return defaultValue;
        }
        return value;
    }

    let accessToken = getAccessToken();
    function getAccessToken(){
        let token = getConfig('accessToken', "");
        function checkForTokenInURL() {
            console.log("checkForTokenInURL");
            console.log("called from",document.referrer);
            const hash = window.location.hash.substring(1);
            const params = new URLSearchParams(hash);
            const token_in_url = params.get('access_token');

            if (token_in_url) {
                token = token_in_url;
                GM_setValue('accessToken', token_in_url); // Guardar el token para futuras solicitudes
            }
        }
        checkForTokenInURL();
        return token;
    }
    function invalidateToken(){
        accessToken=null;
        GM_setValue('accessToken', accessToken);
    }
    function checkAuth() {
        function authenticate() {
            // redirects to AUTH_URL. This url will ask user for permission an then redirect to REDIRECT_URI configured in Azure
            stop_execution_and_jump_to(AUTH_URL);
        }
        console.log("checkAuth");
        if (!accessToken) {
            console.log("autentificar");
            authenticate();
            console.error("no debe pasar por aqui");
        }
        console.log("autentificado! access token ", accessToken);
    }

    const EXCEL_PARAMETERS={
        "group_name":"excel_config",
        "parameter_names":["EXCEL_FILE_ID","SHEET_NAME","SHEET_RANGE"]
    };
    const USER_PARAMETERS={
        "group_name":"user_config",
        "parameter_names":["fuerza técnico","Asunto mail solicitud visita"]
    };

    const EXCEL_FILE_ID = getConfigFromSet(EXCEL_PARAMETERS,'EXCEL_FILE_ID'); // ID del archivo de Excel con proyectos de todos los tecnicos
    const SHEET_NAME = getConfigFromSet(EXCEL_PARAMETERS,'SHEET_NAME'); // hoja preparada con datos para este script
    const SHEET_RANGE = getConfigFromSet(EXCEL_PARAMETERS,'SHEET_RANGE'); // Reemplaza el rango que deseas consultar

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
            desplegable_proyectos={
                "rows":filteredRows,
                "col_names":col_names
            };
            InsertaBuscador();
            console.log('desplegable_proyectos:', desplegable_proyectos);
        })
            .catch(error=>{
            console.error('Excel Data:', error);
        });
    };
    function fetchCalendar(){
        function formatearFecha(fechaUTC) {
            const fecha = new Date(fechaUTC);
            const opciones = { weekday: 'long', day: 'numeric', month: 'long' };
            return fecha.toLocaleDateString('es-ES', opciones);
        }
        let dias_futuros=100;
        let fechaActual = new Date();

        let diasDespues = new Date();
        diasDespues.setDate(fechaActual.getDate() + dias_futuros);

        // const apiUrl = "https://graph.microsoft.com/v1.0/me/events?$filter=contains(subject,'idi-20241234')";
        // const apiUrl = `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${fechaActual.toISOString()}' and contains(subject,'viaje')`;
        const apiUrl = `https://graph.microsoft.com/v1.0/me/events?$filter=start/dateTime ge '${fechaActual.toISOString()}' and start/dateTime le '${diasDespues.toISOString()}'`;

        fetchMicrosoftGraph(apiUrl)
            .then(response=>{
            console.log('Calendar response received:', response);
            const data = JSON.parse(response.responseText);
            console.log('Calendar Data received:', data);
            if (data.value.length > 0) {
                data.value.forEach(evento => {
                    let titulo = evento.subject;
                    let fechaInicio = evento.start.dateTime;
                    let zonaHoraria = evento.start.timeZone;

                    console.log(`Evento: ${titulo}`);
                    console.log(`Fecha: ${fechaInicio} (Zona horaria: ${zonaHoraria})`);
                    console.log(`Fecha: ${formatearFecha(fechaInicio)}`);
                });
            } else {
                console.log("No se encontró el evento.");
            }
        })
    };
    function insertaSelectorFechaSalida(){
        function insertaListaFechas(selector_fecha,opciones){
            // Opciones de fechas

            let input = $(selector_fecha);

            // Crear el contenedor del desplegable
            let dropdown = $("<div id='fecha_dropdown' style='position:absolute; background:white; border:1px solid #ccc; padding:5px; width:320px; max-height:180px; overflow-y:auto; z-index:1000; display:none;'></div>");

            // Agregar opciones al desplegable
            opciones.forEach(opcion => {
                console.log("opcion añadida",opcion);
                let item = $("<div style='padding:5px; cursor:pointer; border-bottom:1px solid #eee;'></div>").text(opcion.display);
                item.on("click", function() {
                    input.val(opcion.fechaTexto); // Insertar la fecha en el input
                    dropdown.hide();
                });
                dropdown.append(item);
            });

            // Insertar el dropdown en el DOM
            $("body").append(dropdown);

            // Mostrar la lista cuando el input recibe foco
            input.on("focus", function() {
                let offset = input.offset();
                let inputWidth = input.outerWidth();
                let dropdownWidth = dropdown.outerWidth();

                dropdown.css({
                    top: offset.top + input.outerHeight(),
                    left: offset.left + inputWidth - dropdownWidth + "px" // Alinear derecha con derecha
                }).show();
            });

            // Ocultar el dropdown si se hace clic fuera
            $(document).on("click", function(e) {
                if (!input.is(e.target) && !dropdown.is(e.target) && dropdown.has(e.target).length === 0) {
                    dropdown.hide();
                }
            });
        };
        function buscaFechasViajeEmail() {
            const twoMonthsAgo = new Date();
            twoMonthsAgo.setMonth(twoMonthsAgo.getMonth() - 5);
            const fechaFiltro = twoMonthsAgo.toISOString(); // Convierte a formato UTC
            const mail_visita_empieza_por = getConfig("Asunto mail solicitud visita ", "Confirmación de visita");

            const apiUrl = `https://graph.microsoft.com/v1.0/me/messages?$filter=receivedDateTime ge ${fechaFiltro} and startswith(subject, '${mail_visita_empieza_por}')&$top=50`;
            let fechas = [];

            fetchMicrosoftGraph(apiUrl)
                .then(response=>{
                let data = JSON.parse(response.responseText);
                let respuestas = data.value.filter(buscaProyectoyFechaEnTexto);

                const fechas_ordenadas=fechas.sort((a, b) => b.fecha - a.fecha);

                console.log("Correos que incluyen fecha:", respuestas);
                insertaListaFechas(
                    "#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtFechaSalida",
                    fechas_ordenadas);
            });

            // Función para convertir las fechas encontradas en formato Date, con ajuste de año
            function convertirFecha(match, fechaCorreo, campos) {
                let dia = parseInt(match.groups.dia);
                let mes = match.groups.mes;
                let año = match.groups.año ? parseInt(match.groups.año) : null;

                // Si no se proporciona año, se infiere del contexto
                if (!año) {
                    let fechaCorreoObj = new Date(fechaCorreo);
                    año = fechaCorreoObj.getFullYear();

                    // Si el mes de la fecha es enero, febrero, etc., y la fecha del correo es en diciembre, se infiere que es el año siguiente
                    const mesesQueSonEnElAñoSiguiente = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre"];
                    if (mesesQueSonEnElAñoSiguiente.includes(mes) && fechaCorreoObj.getMonth() === 11) {
                        año += 1; // Año siguiente
                    }
                }

                // Mapear el mes a su número correspondiente
                const meses = {
                    enero: 0, febrero: 1, marzo: 2, abril: 3, mayo: 4, junio: 5,
                    julio: 6, agosto: 7, septiembre: 8, octubre: 9, noviembre: 10, diciembre: 11
                };

                let mesNum = typeof mes === 'number' ? mes - 1 : meses[mes.toLowerCase()];
                let fecha = new Date(año, mesNum, dia);

                // Asegurarse que la fecha es válida
                if (fecha.getFullYear() !== año || fecha.getMonth() !== mesNum || fecha.getDate() !== dia) {
                    return null; // Fecha no válida
                }

                return fecha;
            }
            function formatearFecha(fechaUTC,numerica=true) {
                const fecha = new Date(fechaUTC);
                const formato = (numerica?
                                 { day: '2-digit', // Día con dos dígitos (ej. "21")
                                  month: '2-digit', // Mes con dos dígitos (ej. "01")
                                  year: 'numeric' // Año completo (ej. "2025")
                                 }:
                                 { weekday: 'short', day: 'numeric', month: 'short' });
                return fecha.toLocaleDateString('es-ES', formato);
            }

            function buscaProyectoyFechaEnTexto(msg){
                const regexProyecto = /(?<proyecto>\b[A-Z]{3}\-\d{8}\b)/i;
                let matchProyecto = regexProyecto.exec(msg.subject);
                if (!matchProyecto) return null;
                let proyecto = matchProyecto.groups.proyecto;
                console.log("proyecto :", proyecto);

                const contenido=msg.bodyPreview;
                const regexFechas = [
                    {
                        // Caso: "lunes, 15 de febrero"
                        regex: /(?<dia_semana>lunes|martes|miércoles|jueves|viernes|sábado|domingo)?,?\s*(?<dia>\d{1,2})\s*(de\s+)?(?<mes>enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)(?:\s*de\s*(?<año>\d{4}))?/gi,
                        campos: ['dia_semana', 'dia', 'mes', 'año']
                    },
                    {
                        // Caso: "15/02/2025"
                        regex: /(?<dia>\d{1,2})[\/\-](?<mes>\d{1,2})[\/\-](?<año>\d{4})/g,
                        campos: ['dia', 'mes', 'año']
                    }
                ];

                for (let { regex, campos } of regexFechas) {
                    let matches = [...contenido.matchAll(regex)];
                    matches.forEach(match => {
                        console.log("match :", match);
                        let fecha = convertirFecha(match, msg.receivedDateTime, campos);
                        if (fecha) fechas.push({"fecha":fecha,"fechaTexto":formatearFecha(fecha,true), "proyecto":proyecto, "subject":msg.subject,
                                                "display":formatearFecha(fecha,false)+", "+msg.subject});
                    });
                }
                if (!fechas) return null;

                console.log("fechas :", fechas);

                return fechas;
            };

        }
        buscaFechasViajeEmail();
    }

    let desplegable_proyectos=[];
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
        else {return;}
        if (this.observerOrganosFinanciadores) {
            console.error("organos financiadores descactiva antiguo");
            this.observerOrganosFinanciadores.disconnect();
            this.observerOrganosFinanciadores=null;
        }

        // https://stackoverflow.com/questions/64712065/how-to-insert-mutationobserver-into-jquery-code
        // Now we watch for new elements
        this.observerOrganosFinanciadores = new MutationObserver(() => {
            console.error("organos financiadores mutó");
            const detActividades='#ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_detActividadesProductos_';
            if ($(detActividades+'btnAnyadirProducto').length) {
                console.log('Aparecio el boton de añadir producto');
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
                console.error("organos financiadores. Dejo de observar");
                this.observerOrganosFinanciadores.disconnect();
                this.observerOrganosFinanciadores=null;
                beep();
            }
        });
        console.error("organos financiadores. instala observador");
        this.observerOrganosFinanciadores.observe(document.getElementById('ctl00_SheetContentPlaceHolder_UCOtrosDatosActividadesProductos_updOtrosDatos'), {
            childList: true,
            subtree: true
        });
    }
    function InsertaHorasyTransportePreferido(){
        if ($('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtHoraSalida').val()=="") {
            $('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtHoraSalida').attr('value','06:00');
        }
        if ($('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtHoraLlegada').val()=="") {
            $('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_txtHoraLlegada').attr('value','20:00');
        }
        if ($('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_ddlLocomocion option[value="-"]').is(':selected')) {
            $('#ctl00_SheetContentPlaceHolder_UCTramos_dgTramos_ctl02_ddlLocomocion option[value="TR"]').prop('selected', true);
        }
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
    function inserta_proyecto_nuevo(match_groups) {
        function prefijo_proyecto_nuevo() {
            // busco en la lista de proyectos el que no tenga el año relleno
            const proyecto_vacio=$('input[id^="ctl00_SheetContentPlaceHolder_UCProyectos_gvProyectos"][id$="txtAnioProy"]').filter(function() { return this.value == ""; });
            var result="";
            if (proyecto_vacio.length) {
                const match=/(?<prefijo>.*)txtAnioProy/.exec(proyecto_vacio.attr("id"));
                result = match.groups.prefijo;
            }
            console.log("prefijo_proyecto_nuevo",result);
            return result;
        }
        function insertaIDI_hueco(match_groups, prefijoProyecto) {
            console.log("insertaIDI_hueco");
            if (! ["area","anioProy","numProy","hito"].every(key => key in match_groups)) {
                console.error("falta algun dato en ",match_groups);
                return;
            };
            $("#"+prefijoProyecto+'ddlAreas option[value="'+match_groups.area+'"]').prop('selected', true);
            $("#"+prefijoProyecto+'txtAnioProy').attr('value',match_groups.anioProy);
            $("#"+prefijoProyecto+'txtNumProy').attr('value',match_groups.numProy);
            $("#"+prefijoProyecto+'txtHito').attr('value',match_groups.hito);
        }

        console.log("inserta_proyecto_nuevo",match_groups);
        var prefijo_nuevo=prefijo_proyecto_nuevo();
        if (prefijo_nuevo!="") {
            insertaIDI_hueco(match_groups, prefijo_nuevo);
        } else {
            const observerProyectos = new MutationObserver(() => {
                console.log("observerProyectos");
                prefijo_nuevo=prefijo_proyecto_nuevo();
                if (prefijo_nuevo!="") {
                    insertaIDI_hueco(match_groups, prefijo_nuevo);
                    observerProyectos.disconnect();
                }
            });
            //no funciona con 'ctl00_SheetContentPlaceHolder_UCProyectos_gvProyectos'
            observerProyectos.observe(document.getElementById('contenedor'), {
                childList: true,
                subtree: true
            });
            console.log("click nuevo proyecto");
            $('#ctl00_SheetContentPlaceHolder_UCProyectos_gvProyectos_ctl01_imgAddProyecto').trigger('click');
        }
    }

    function InsertaProyectoSeleccionado(texto_desplegable){
        const index_texto_desplegable=desplegable_proyectos.col_names.indexOf("texto_desplegable");
        const filteredRows = desplegable_proyectos.rows.filter((row) => row[index_texto_desplegable] === texto_desplegable); //.includes(tecnico));
        if (filteredRows.length != 1) {
            alert("he encontrado más de un elemento correspondiente a ",texto_desplegable);
            return;
        }
        const proyecto_seleccionado = Object.fromEntries(desplegable_proyectos.col_names.map((col, i) => [col, filteredRows[0][i]]));
        console.log("proyecto seleccionado",proyecto_seleccionado);
        InsertaMotivo(proyecto_seleccionado.motivo);
        InsertaOrigen();
        InsertaDestino(proyecto_seleccionado.provincia_desa,proyecto_seleccionado.localidad_desa);
        InsertaHorasyTransportePreferido();
        console.error("proyecto_seleccionado", proyecto_seleccionado);

        let match=/(?<area>\w+)\-(?<anioProy>\d{4})(?<numProy>\d{4})/.exec(proyecto_seleccionado.proyecto);
        if (!match) {
            console.error(`no reconozco el proyecto {proyecto_seleccionado.proyecto}` );
            return;
        }
        inserta_proyecto_nuevo({
            area: match.groups.area,
            anioProy: match.groups.anioProy,
            numProy: match.groups.numProy,
            hito: proyecto_seleccionado.hito_viaje
        });
        setTimeout(function() {
            InsertaActividades_optionValue(proyecto_seleccionado.tipo_a_desplegable_id_desplegable);
        }, 4000);

    };
    function beep() {
        (new Audio(
            "data:audio/wav;base64,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 /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////VEFHAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAU291bmRib3kuZGUAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMjAwNGh0dHA6Ly93d3cuc291bmRib3kuZGUAAAAAAAAAACU="
        )).play();
    }
    function desactivaAutocomplete() {
        $('input, textarea, select').attr('autocomplete', 'off');
    }
    function retocaLayout() {
        // console.error('***** retocaLayout');
        const targetNode = document.getElementById(
            'ctl00_UpdatePanel1');
        // Configura el MutationObserver
        function corrigeAspecto(){
            // console.error('***** corrigeAspecto');
            // organos financiadores tiene colspan3, y al añadir la modificación de adjuntar archivo, no se amplió a colspan 4.
            $('td[colspan="3"]', targetNode).attr('colspan', '4');

            // amplia cuadro de texto de motivo
            $('#ctl00_SheetContentPlaceHolder_UCDGViajes_txtMotivo')
                .css('width', '750px')
                .css('height', '80px')
        }
        const observer = new MutationObserver(function (mutationsList, observer) {
            // console.error('***** hay mutacion');
            corrigeAspecto();
        });

        const config = {
            childList: true,    // Monitoriza cambios en los hijos directos del nodo
            subtree: true      // Monitoriza cambios en el subárbol completo
        };

        // Inicia la observación
        if (targetNode) {
            // console.error('***** inicia la observacion');
            corrigeAspecto();
            observer.observe(targetNode, config);
        } else {
            console.error('El elemento no fue encontrado.');
        }
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
                    var resultados = desplegable_proyectos.rows.filter(item => {
                        const todos_los_campos=(''+item);
                        // console.log("todos los campos",todos_los_campos);
                        return todos_los_campos.toLowerCase().includes(texto)
                    });

                    if (resultados.length > 0) {
                        const index_texto_desplegable=desplegable_proyectos.col_names.indexOf("texto_desplegable");
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
                    var resultados = desplegable_proyectos.rows.filter(item => {
                        const todos_los_campos=(''+item);
                        // console.log("todos los campos",todos_los_campos);
                        return todos_los_campos.toLowerCase().includes(texto)
                    });

                    if (resultados.length > 0) {
                        const index_texto_desplegable=desplegable_proyectos.col_names.indexOf("texto_desplegable");
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
    function insertaBotonMiTampermonkey() {
        // añade boton Mi tampermonkey, para que se vea que la pagina está hackeada
        $('body').append('<input type="button" value="Mi Tampermonkey" id="CP">')
        $("#CP").css("position", "fixed").css("top", 0).css("left", 0);
        $('#CP').click(function(){
            // a veces falla inserta actividades. al picar el boton reintenta
            console.log("mi tampermonkey. reintento inserta actividades");
            retocaLayout();
            insertaSelectorFechaSalida();
            // buscarRespuestasConfirmacion();
        });

    }

    gestor_configuracion();
    checkAuth();
    retocaLayout();
    //fetchCalendar();
    fetchExcelData();
    desactivaAutocomplete();
    insertaBotonMiTampermonkey();

})();