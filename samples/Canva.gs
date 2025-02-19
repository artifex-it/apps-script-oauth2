/*
Aggiornare le credenziali:

var CANVA_CLIENT_ID = "...";
var CANVA_CLIENT_SECRET = "...";

Fare il 'deploy' dello script.

Il redirect URL da indicare nella pagina di autorizzazione di Canva viene mostrato alla prima esecuzione dello script.

Una volta autorizzato si possono eseguire le funzioni definite qui sotto direttamente con il comando 'Esegui' del menù senza dover ripetere il 'deploy'.
*/


var CANVA_CLIENT_ID = "...";
var CANVA_CLIENT_SECRET = "...";


// Legge il profilo
function get_profile() {
  let result = execQuery("https://api.canva.com/rest/v1/users/me/profile");

  let out = JSON.stringify(result, null, 2);

  Logger.log(out);

}


// Scarica l'elenco di tutti gli elementi della cartella 'root'
function get_folder_items() {
  let result = execQuery("https://api.canva.com/rest/v1/folders/root/items");

  let out = JSON.stringify(result, null, 2);

  Logger.log(out);

}


// Scarica l'elenco di tutti i documenti Canva
function get_designs() {
  let result = execQuery("https://api.canva.com/rest/v1/designs");

  let out = JSON.stringify(result, null, 2);

  Logger.log(out);

}


// Invia la prima immagine del documento Google Docs
function send_first_image() {
  let body = DocumentApp.getActiveDocument().getActiveTab().asDocumentTab().getBody();
  let image = body.getImages()[0].getBlob();

  let result = execQuery("https://api.canva.com/rest/v1/asset-uploads",
                         {
                          "method": "POST",
                          "headers": {
                            "Content-Type": "application/octet-stream",
                            "Asset-Upload-Metadata": JSON.stringify({"name_base64": Utilities.base64Encode(image.getName() || "IMMAGINE DI PROVA")}),
                          },
                          "payload": image
                         });

  let out = JSON.stringify(result, null, 2);

  let job_id = result.job.id;
  let job_status = result.job.status;

  while (job_status == "in_progress") {

    result = execQuery("https://api.canva.com/rest/v1/asset-uploads/"+job_id);
    job_status = result.job.status;

    out += JSON.stringify(result, null, 2);
  }

  Logger.log(out);
}


// Crea un documento in Canva
function create_design() {
  let result = execQuery("https://api.canva.com/rest/v1/designs",
                         {
                          "method": "POST",
                          "headers": {
                            "Content-Type": "application/json",
                          },
                          "payload": JSON.stringify({
                            "title": "DOCUMENTO DI PROVA",
                            "design_type": {
                              "type": "preset",
                              "name": "doc"
                            },
                            "asset_id": "..."  //<-- Indicare l'asset id di un'immagine precedentemte inviata (v. send_first_image())
                          })
                         });

  let out = JSON.stringify(result, null, 2);

  Logger.log(out);

}


// Fa il reset delle autorizzazioni
function reset() {
  let oauth = getOauth_();
  oauth.reset();
}


/* ******************** */
/* FUNZIONI DI SERVIZIO */
/* ******************** */

function execQuery (url, params={}) {
    let oauth = getOauth_();
    if (params["headers"]) {
      params["headers"]["Authorization"] = "Bearer " + oauth.getAccessToken();
    } else {
      params["headers"] = {"Authorization": "Bearer " + oauth.getAccessToken()};
    }
    params["muteHttpException"] = true;

    let response = UrlFetchApp.fetch(url, params);

    let result = {};

    if (response.getResponseCode() >= 400) {
      oauth.reset();
    } else {
      result = JSON.parse(response.getContentText());
    }

    return result;
}


function doGet(request) {
  let oauth = getOauth_(request);

  let out = "";
  if (oauth.hasAccess()) {
    result = execQuery("https://api.canva.com/rest/v1/users/me/profile");
    out += "<p>Sei collegato come utente: <strong>" + result.profile.display_name + "</strong></p>";
    out += "<p>Ora puoi usare la connessione con Canva.</p>";
  } else {
    out += "<p>Url di questo script da inserire come redirect URL in Canva:<br><strong>" + ScriptApp.getService().getUrl() + "</strong></p>";
    // out += "<br>code_verifier: " + oauth.getStorage().getValue("code_verifier_");
    out += '<p><a href="' + oauth.getAuthorizationUrl() +'" target="_blank">Clicca sul seguente link per autorizzare il collegamento</a></p>';
  }
  return HtmlService.createHtmlOutput(out);
}


function getOauth_(request=null) {
  let userProps = PropertiesService.getUserProperties();
  let oauth = OAuth2.createService("Canva")
    // Set the endpoint URLs
    .setAuthorizationBaseUrl("https://www.canva.com/api/oauth/authorize")
    .setTokenUrl("https://api.canva.com/rest/v1/oauth/token")

    .setClientId(CANVA_CLIENT_ID)
    .setClientSecret(CANVA_CLIENT_SECRET)

    // Set the name of the callback function that should be invoked to
    // complete the OAuth flow. Required but not used for Canva API
    .setCallbackFunction("uthCallback")

    .setRedirectUri(ScriptApp.getService().getUrl())

    // Set the property store where authorized tokens should be persisted.
    .setPropertyStore(userProps)

    // Set the scopes to request (space-separated services).
    .setScope("asset:read asset:write design:content:read design:content:write design:meta:read design:permission:read folder:read folder:write folder:permission:read folder:permission:write profile:read")

    .generateCodeVerifier()

    .setTokenHeaders({
      "Authorization": "Basic " + Utilities.base64Encode(CANVA_CLIENT_ID + ":" + CANVA_CLIENT_SECRET),
      "Content-Type": "application/x-www-form-urlencoded"
    });

    if (! oauth.hasAccess()) {
      if (request && request.parameters["code"]) {
        if (! oauth.handleCallback(request)) {
          // Authorization failed
          oauth.reset();
        } else {
          // No request
          oauth.reset;
        }
      }
    }

    return oauth;
}

function authCallback(request) {
  var oauth = getOauth_();
  var authorized = oauth.handleCallback(request);
  if (authorized) {
    return HtmlService.createHtmlOutput("Success!");
  } else {
    return HtmlService.createHtmlOutput("Denied.");
  }
}