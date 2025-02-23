// Your UserScript needs at least these tags copied on  our main tampermonkey script.
// @match        http://your.intranet.site/*
// @match        https://www.google.com/search?q=redirect_to_your_intranet_site*
// @grant        GM_xmlhttpRequest
// @grant        GM_listValues
// @grant        GM_getValue
// @grant        GM_setValue
// @grant        GM_addStyle
// @grant        GM_deleteValue
// @require      https://cdn.jsdelivr.net/gh/cesarpino/Excel360toIntranet@main/Http-Oauth2-authorize.js
// ==/YOUR UserScript==

'use strict';
//***********
// WEBAPP token purchase stuff
const this_page_url = window.location.href;
const WEBAPP_URI = get_match_from_UserScript("http://");
const REDIRECT_URI = get_match_from_UserScript("https://");
if (this_page_url.startsWith(REDIRECT_URI)) {
    // We are inside REDIRECT_URI, we force redirection to WEBAPP_URI with the token parameter recieved.
    let back_to_webapp_uri = `${WEBAPP_URI}${this_page_url.slice(REDIRECT_URI.length)}`;
    alert("Redirects authorization to \n"+back_to_webapp_uri);
    stop_execution_and_jump_to(back_to_webapp_uri);
}

// We are in now an WEBAPP_URI page, and want to get an access_token to enable microsoft access.
const AUTH_WEBAPP_PARAMETERS={
    "group_name":"auth_webapp",
    "parameter_names":["TENANT_ID","CLIENT_ID","SCOPES"]
};

const TENANT_ID = getConfigFromSet(AUTH_WEBAPP_PARAMETERS,"TENANT_ID"); // or replace with organization TENANT_ID
const CLIENT_ID = getConfigFromSet(AUTH_WEBAPP_PARAMETERS,"CLIENT_ID"); // or replace with clientId of app configured in Azure. ex "Acceso a OneDrive desde viajes.cdti.es" en Azure
const SCOPES = getConfigFromSet(AUTH_WEBAPP_PARAMETERS,"SCOPES"); // or replace with e.g. Files.Read Files.Read.All Calendars.Read Mail.Read

// you must create a Azure app, and get CLIENT_ID and configure same REDIRECT_URI
// REDIRECT_URI must be configured also in azure associated with client_id.
const AUTH_URL_BASE = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/authorize`;
const AUTH_URL = `${AUTH_URL_BASE}?client_id=${CLIENT_ID}&response_type=token&redirect_uri=${REDIRECT_URI}&scope=${SCOPES}&response_mode=fragment`;

// get the stored microsoft api token, or as the user for his approval.  
let accessToken = getAccessToken();
checkAuth();

// helper functions.
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
function getAccessToken(){
    function checkForTokenInURL() {
        const hash = window.location.hash.substring(1);
        const params = new URLSearchParams(hash);
        return params.get('access_token');
    }

    const token_in_url = checkForTokenInURL();
    const previous_token = getConfig('accessToken', "");

    if (token_in_url) {
        console.log("Token received in url, called from ",document.referrer); // REDIRECT_URI
        GM_setValue('accessToken', token_in_url); // Guardar el token para futuras solicitudes
        return token_in_url;
    } else {
        return previous_token;
    }
}
function checkAuth() {
    console.log("checkAuth");
    if (!accessToken) {
        // redirects to AUTH_URL where microsoft will ask user for permission and then redirect to REDIRECT_URI configured in Azure
        stop_execution_and_jump_to(AUTH_URL);
    }
    console.log("autentificado! access token ", accessToken);
}
function invalidateToken(){
    accessToken=null;
    GM_setValue('accessToken', accessToken);
}
function showConfigEditor() {
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
