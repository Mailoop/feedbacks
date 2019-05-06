/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */


var toRecipientsArray = [];
var ccRecipientsArray = [];
var item;
var token;
var action_status;
var mailbox;
var current_event;
var signature_appended = false;

// The initialize function must be run each time a new page is loaded
Office.initialize = function (reason) {
  mailbox = Office.context.mailbox;
};

// Add any ui-less function here
function activateDeconnection (event) {
    action_status = "activate";
    current_event = event;
    getIdentityToken(makeCorsRequestRule);
};

// Add any ui-less function here
function deactivateDeconnection (event) {
    action_status = "deactivate";
    current_event = event;
    getIdentityToken(makeCorsRequestRule);
};

function prependFeedbackFeedbackButton (event) {
    var unique_token;
    unique_token = "123456";
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: event },
      function callback(result) {
          // Do something with the result
          if (result.value.indexOf("mailoop_tag") === -1){
              Office.context.mailbox.item.body.prependAsync(
              '<div id="mailoop_tag" style="background-color: #EC7D31;text-align: center"><a href="https://app.mailoop.com/votes?token=' + unique_token + '" style="text-decoration: none;"><font style="font-family:&quot;Segoe UI Light&quot;; color: #ffffff;font-size: 14px; text-transform: uppercase;">Exprimer anonymement un feedback avec Mailoop</font></a></div><br>',
              {coercionType: Office.CoercionType.Html, asyncContext: result.asyncContext },
              prepend_callback);
          } else {
              result.asyncContext.completed();
          }
    });
}

function prepend_callback(asyncResult) {
    asyncResult.asyncContext.completed();;
}

function getIdentityToken (successCallback) {
    // récupérer le token d'identification
    mailbox.getUserIdentityTokenAsync(function (asyncResult) {
        token = asyncResult.value;
        successCallback();
    });
};

// Make the XHR request.
function makeCorsRequestRule() {
    var url = 'https://app.mailoop.com/api/v1/rules';
    var xhr = createCORSRequest('PATCH', url);

    if (!xhr) {
        alert('CORS not supported');
        return;
    };

    // Response handlers.
    xhr.onload = function () {
        if (xhr.status == 200) {
            ruleCallback();
        } else {
            errorCallback();
        }
    };

    xhr.onerror = function () {
    };

    var params = JSON.stringify( { emailAddress: mailbox.userProfile.emailAddress, action_status: action_status, token: token, time_zone: convertMsToIanaTimezone(mailbox.userProfile.timeZone), rule: {} } );
    xhr.send(params);
};

// Create the XHR object. Shared for all XHR request.
function createCORSRequest(method, url) {
    var xhr = new XMLHttpRequest();
    if ("withCredentials" in xhr) {
        // XHR for Chrome/Firefox/Opera/Safari.
        xhr.open(method, url, true);

    } else if (typeof XDomainRequest !== "undefined") {
        // XDomainRequest for IE.
        xhr = new XDomainRequest();
        xhr.open(method, url);
    } else {
        // CORS not supported.
        xhr = null;
    }
    xhr.setRequestHeader('Content-type', 'application/json;charset=UTF-8');
    return xhr;
};

function ruleCallback () {
    var message;
    if (action_status === "activate") {
      message = "Smart Deconnexion is ON. No more notifications from 8pm to 8am and during week-ends, except for messages marked as Important.";
    } else {
      message = "Smart Deconnexion is OFF. You will receive all notifications.";
    }

    Office.context.mailbox.item.notificationMessages.replaceAsync("deconnexion_status", {
        type: "informationalMessage",
        message : message,
        icon : "icon16",
        persistent: false
    });
    current_event.completed();
}

function errorCallback () {
    Office.context.mailbox.item.notificationMessages.replaceAsync("error", {
        type: "errorMessage",
        message : "Action couldn't be performed."
    });
    current_event.completed();
}

function convertMsToIanaTimezone(ms_timezone){
  var timezones = {
    "W. Central Africa Standard Time":"Africa/Algiers",
    "Egypt Standard Time":"Africa/Cairo",
    "Morocco Standard Time":"Africa/Casablanca",
    "South Africa Standard Time":"Africa/Harare",
    "South Africa Standard Time":"Africa/Johannesburg",
    "Greenwich Standard Time":"Africa/Monrovia",
    "E. Africa Standard Time":"Africa/Nairobi",
    "Argentina Standard Time":"America/Argentina/Buenos_Aires",
    "SA Pacific Standard Time":"America/Bogota",
    "Venezuela Standard Time":"America/Caracas",
    "Central Standard Time":"America/Chicago",
    "Mountain Standard Time (Mexico)":"America/Chihuahua",
    "Mountain Standard Time":"America/Denver",
    "Greenland Standard Time":"America/Godthab",
    "Central America Standard Time":"America/Guatemala",
    "SA Western Standard Time":"America/Guyana",
    "Atlantic Standard Time":"America/Halifax",
    "US Eastern Standard Time":"America/Indiana/Indianapolis",
    "Alaskan Standard Time":"America/Juneau",
    "SA Western Standard Time":"America/La_Paz",
    "SA Pacific Standard Time":"America/Lima",
    "SA Pacific Standard Time":"America/Lima",
    "Pacific Standard Time":"America/Los_Angeles",
    "Mountain Standard Time (Mexico)":"America/Mazatlan",
    "Central Standard Time (Mexico)":"America/Mexico_City",
    "Central Standard Time (Mexico)":"America/Mexico_City",
    "Central Standard Time (Mexico)":"America/Monterrey",
    "Montevideo Standard Time":"America/Montevideo",
    "Eastern Standard Time":"America/New_York",
    "US Mountain Standard Time":"America/Phoenix",
    "Canada Central Standard Time":"America/Regina",
    "Pacific SA Standard Time":"America/Santiago",
    "E. South America Standard Time":"America/Sao_Paulo",
    "Newfoundland Standard Time":"America/St_Johns",
    "Pacific Standard Time":"America/Tijuana",
    "Central Asia Standard Time":"Asia/Almaty",
    "Arabic Standard Time":"Asia/Baghdad",
    "Azerbaijan Standard Time":"Asia/Baku",
    "SE Asia Standard Time":"Asia/Bangkok",
    "SE Asia Standard Time":"Asia/Bangkok",
    "China Standard Time":"Asia/Chongqing",
    "Sri Lanka Standard Time":"Asia/Colombo",
    "Bangladesh Standard Time":"Asia/Dhaka",
    "Bangladesh Standard Time":"Asia/Dhaka",
    "China Standard Time":"Asia/Hong_Kong",
    "North Asia East Standard Time":"Asia/Irkutsk",
    "SE Asia Standard Time":"Asia/Jakarta",
    "Israel Standard Time":"Asia/Jerusalem",
    "Afghanistan Standard Time":"Asia/Kabul",
    "Russia Time Zone 11":"Asia/Kamchatka",
    "Pakistan Standard Time":"Asia/Karachi",
    "Pakistan Standard Time":"Asia/Karachi",
    "Nepal Standard Time":"Asia/Kathmandu",
    "India Standard Time":"Asia/Kolkata",
    "India Standard Time":"Asia/Kolkata",
    "India Standard Time":"Asia/Kolkata",
    "India Standard Time":"Asia/Kolkata",
    "North Asia Standard Time":"Asia/Krasnoyarsk",
    "Singapore Standard Time":"Asia/Kuala_Lumpur",
    "Arab Standard Time":"Asia/Kuwait",
    "Magadan Standard Time":"Asia/Magadan",
    "Arabian Standard Time":"Asia/Muscat",
    "Arabian Standard Time":"Asia/Muscat",
    "N. Central Asia Standard Time":"Asia/Novosibirsk",
    "Myanmar Standard Time":"Asia/Rangoon",
    "Arab Standard Time":"Asia/Riyadh",
    "Korea Standard Time":"Asia/Seoul",
    "China Standard Time":"Asia/Shanghai",
    "Singapore Standard Time":"Asia/Singapore",
    "Russia Time Zone 10":"Asia/Srednekolymsk",
    "Taipei Standard Time":"Asia/Taipei",
    "West Asia Standard Time":"Asia/Tashkent",
    "Georgian Standard Time":"Asia/Tbilisi",
    "Iran Standard Time":"Asia/Tehran",
    "Tokyo Standard Time":"Asia/Tokyo",
    "Tokyo Standard Time":"Asia/Tokyo",
    "Tokyo Standard Time":"Asia/Tokyo",
    "Ulaanbaatar Standard Time":"Asia/Ulaanbaatar",
    "Central Asia Standard Time":"Asia/Urumqi",
    "Vladivostok Standard Time":"Asia/Vladivostok",
    "Yakutsk Standard Time":"Asia/Yakutsk",
    "Ekaterinburg Standard Time":"Asia/Yekaterinburg",
    "Caucasus Standard Time":"Asia/Yerevan",
    "Azores Standard Time":"Atlantic/Azores",
    "Cape Verde Standard Time":"Atlantic/Cape_Verde",
    "UTC-02":"Atlantic/South_Georgia",
    "Cen. Australia Standard Time":"Australia/Adelaide",
    "E. Australia Standard Time":"Australia/Brisbane",
    "AUS Central Standard Time":"Australia/Darwin",
    "Tasmania Standard Time":"Australia/Hobart",
    "AUS Eastern Standard Time":"Australia/Melbourne",
    "AUS Eastern Standard Time":"Australia/Melbourne",
    "W. Australia Standard Time":"Australia/Perth",
    "AUS Eastern Standard Time":"Australia/Sydney",
    "UTC":"UTC",
    "W. Europe Standard Time":"Europe/Amsterdam",
    "GTB Standard Time":"Europe/Athens",
    "Central Europe Standard Time":"Europe/Belgrade",
    "W. Europe Standard Time":"Europe/Berlin",
    "W. Europe Standard Time":"Europe/Berlin",
    "Central Europe Standard Time":"Europe/Bratislava",
    "Romance Standard Time":"Europe/Brussels",
    "GTB Standard Time":"Europe/Bucharest",
    "Central Europe Standard Time":"Europe/Budapest",
    "Romance Standard Time":"Europe/Copenhagen",
    "GMT Standard Time":"Europe/Dublin",
    "FLE Standard Time":"Europe/Helsinki",
    "Turkey Standard Time":"Europe/Istanbul",
    "Kaliningrad Standard Time":"Europe/Kaliningrad",
    "FLE Standard Time":"Europe/Kiev",
    "GMT Standard Time":"Europe/Lisbon",
    "Central Europe Standard Time":"Europe/Ljubljana",
    "GMT Standard Time":"Europe/London",
    "GMT Standard Time":"Europe/London",
    "Romance Standard Time":"Europe/Madrid",
    "Belarus Standard Time":"Europe/Minsk",
    "Russian Standard Time":"Europe/Moscow",
    "Russian Standard Time":"Europe/Moscow",
    "Romance Standard Time":"Europe/Paris",
    "Central Europe Standard Time":"Europe/Prague",
    "FLE Standard Time":"Europe/Riga",
    "W. Europe Standard Time":"Europe/Rome",
    "Russia Time Zone 3":"Europe/Samara",
    "Central European Standard Time":"Europe/Sarajevo",
    "Central European Standard Time":"Europe/Skopje",
    "FLE Standard Time":"Europe/Sofia",
    "W. Europe Standard Time":"Europe/Stockholm",
    "FLE Standard Time":"Europe/Tallinn",
    "W. Europe Standard Time":"Europe/Vienna",
    "FLE Standard Time":"Europe/Vilnius",
    "Russian Standard Time":"Europe/Volgograd",
    "Central European Standard Time":"Europe/Warsaw",
    "Central European Standard Time":"Europe/Zagreb",
    "Samoa Standard Time":"Pacific/Apia",
    "New Zealand Standard Time":"Pacific/Auckland",
    "New Zealand Standard Time":"Pacific/Auckland",
    "Tonga Standard Time":"Pacific/Fakaofo",
    "Fiji Standard Time":"Pacific/Fiji",
    "Central Pacific Standard Time":"Pacific/Guadalcanal",
    "West Pacific Standard Time":"Pacific/Guam",
    "Hawaiian Standard Time":"Pacific/Honolulu",
    "UTC+12":"Pacific/Majuro",
    "UTC-11":"Pacific/Midway",
    "UTC-11":"Pacific/Midway",
    "Central Pacific Standard Time":"Pacific/Noumea",
    "UTC-11":"Pacific/Pago_Pago",
    "West Pacific Standard Time":"Pacific/Port_Moresby",
    "Tonga Standard Time":"Pacific/Tongatapu"
  };
  return timezones[ms_timezone];
}
