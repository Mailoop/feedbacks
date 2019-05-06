// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.

/// <reference path="../App.js" />

var exchangeIdentityToken;
var userEmail;
var item;
var itemDatas;
var itemId;
var mailbox;
var perimeter;
var _settings;
var _props;
var voted_behaviors;
var behaviors;
var internalDomains;
var userStatus;
var defaultLocale = "fr";
var backendURI = "https://region-normandie.app.mailoop.com";
var dashboardURI = "https://region-normandie.dashboard.mailoop.com";
var msIconNames = {"positive": "Emoji2", "neutral": "EmojiNeutral", "negative": "EmojiDisappointed"};
var deconnexionSettings;
var toggles = {};
var button;
var checkboxes = {};
var defaultDeconnexionSettings =  {'smart_deconnexion_set': false, 'smart_deconnexion_enabled': false, 'country': "FR", 'time_zone': "Europe/Paris", 'working_time': {"1": [9,18],"2": [9,18],"3": [9,18],"4": [9,18],"5": [9,18]}, 'is_enabled': {'nights': true, 'weekends': true, 'vacations': false}};
var analyticsRequestedScopes = "Calendars.Read MailboxSettings.Read User.Read Mail.Read offline_access";
var minimumAnalyticsScopes = "Calendars.Read MailboxSettings.Read User.Read Mail.Read";
var smartDeconnexionRequestedScopes = "Mail.ReadWrite Calendars.Read MailboxSettings.ReadWrite User.Read offline_access";
var minimumDeconnexionScopes = "Mail.ReadWrite Calendars.Read MailboxSettings.ReadWrite";
var current_scopes;
var statusPromise = false;
var domainsPromise = false;
var behaviorsPromise = false;
var apiTestPromise = false;
var lastUpdate;
var productChoice;
var apiAvailable;


(function () {
 	"use strict";

	// The Office initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();
			_settings = Office.context.roamingSettings;
			mailbox = Office.context.mailbox;
			userEmail = _settings.get('userEmail') || mailbox.userProfile.emailAddress;
			item = mailbox.item;
			var SpinnerElements = document.querySelectorAll(".ms-Spinner");
			for (var i = 0; i < SpinnerElements.length; i++) {
			    new fabric['Spinner'](SpinnerElements[i]);
			}
			mailbox.getUserIdentityTokenAsync(function (asyncResult) {
		     	exchangeIdentityToken = asyncResult.value;
				if (isInitialized()) {
					appLoader();
				} else {
					appInitializer();
				}
			});
		});
	};

	// Initialize datas if not persisted or if older than 1 day
	function isInitialized(){
		lastUpdate = undefined;
		// lastUpdate = _settings.get('lastUpdate');
		return !((lastUpdate === undefined || Math.abs(new Date() - lastUpdate) > 86400000 ))
	}

	function appInitializer(){
		withBehaviors(initializationComplete);
		withStatus(initializationComplete);
		withDomains(initializationComplete);
		withApiTest(initializationComplete);
	}

	// Update behaviors list - trigggered at first launch and when behaviors_version differs between Mailoop server and client
	function withBehaviors(callback) {
	    var url = backendURI + '/api/v2/company/behaviors';
	    var params = {};
	    sendCORSRequest('GET', url, behaviorsAllCallback, params, callback);
	}

	function withDomains(callback){
		var url = backendURI + '/api/v2/company/domains';
	    var params = {};
	    sendCORSRequest('GET', url, domainsCallback, params, callback);
	}

	function withStatus(callback){
		var url = backendURI + '/api/v2/employees/' + userEmail + '/status';
	    var params = {};
	    sendCORSRequest('GET', url, statusCallback, params, callback);
	}

	function withApiTest(callback){
		if ( Office.context.requirements.isSetSupported('Mailbox', 1.5) ) {
			mailbox.getCallbackTokenAsync({isRest: true}, function(result){
		  		if (result.status === "succeeded") {
		    		var accessToken = result.value;
		    		// Use the access token
		    		testApiAvailability(accessToken, callback);
		  		}
			});
		} else {
			mailbox.getCallbackTokenAsync(function(result){
		  		if (result.status === "succeeded") {
		    		var accessToken = result.value;
		    		// Use the access token
		    		testApiAvailability(accessToken, callback);
		  		}
			});
		}
	}

	
	function initializationComplete(){
		if ( statusPromise && domainsPromise && behaviorsPromise && apiTestPromise ) {
			lastUpdate = new Date ();
			_settings.set('lastUpdate', lastUpdate);
			_settings.saveAsync();
			appLoader();
		}
	}

	function appLoader() {
		deconnexionSettings = _settings.get('deconnexionSettings') || defaultDeconnexionSettings;
		internalDomains = _settings.get('internalDomains');
		behaviors = _settings.get('behaviors');
		userStatus = _settings.get('userStatus');
		apiAvailable = _settings.get('apiAvailable');
		productChoice = getProductChoiceFromStatus(); //  ||  _settings.get('productChoice');
		if (userStatus.is_opposed) {
			showPage("opposition");
			hideHeader();
		} else if ( productChoice === undefined && (userStatus.personal_grant_company_policy === null || userStatus.personal_grant_company_policy)) {
			showPage("product-choice");
		} else {
			showPage("content-main");
		}

		buildMainPage();
		buildProductChoicePage();
		buildSettingsPage();
		eventsHandlers();

		if (isPersistenceSupported()) {
			// Set up ItemChanged event
			mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem);
		}
	}

	function buildProductChoicePage(){
		$('.product-card').hide();
		userStatus.products_allowed.forEach(function(product){
			$('#product-' + product).show();
		})
		if (productChoice) {
			$("." + productChoice).addClass('active');
			$("." + productChoice + ' i').attr("class", "ms-Icon ms-Icon--CompletedSolid")
		}
	}

	function buildMainPage(){
		var ToggleElement = document.querySelector("#smartdeco-switch");
		var activated = deconnexionSettings.smart_deconnexion_enabled;
		var set = deconnexionSettings.smart_deconnexion_set;
		toggles["smartdeco-switch"] = toggles["smartdeco-switch"] || new fabric['Toggle'](ToggleElement);
		if (activated) {
			$("#smartdeco").addClass('is-selected');
		} else if ((productChoice !== "product-smartdeconnexion") || !(set)) {
			$("#smartdeco-switch").addClass('is-disabled');
			$("#smartdeco-switch").attr('title',"La Smart Deconnexion n'est actuellement pas disponible.");
		}
		if (userStatus.email_validated === undefined || !(userStatus.email_validated)){
			$("#dashboard i").addClass("disabled").attr('title', 'Validez votre adresse email pour accéder au tableau de bord !');
		}
		// Initialize list of behaviors if first launch of addin - else launch addin activation tests
		testIfPlugInActive();
	}

	function buildSettingsPage(){
		// Existing values and update, or default values for first setup
		var CheckBoxElements = document.querySelectorAll(".ms-CheckBox");
		for (var i = 0; i < CheckBoxElements.length; i++) {
			checkboxes[i] = checkboxes[i] || new fabric['CheckBox'](CheckBoxElements[i]);
			if ( deconnexionSettings.working_time && Object.keys(deconnexionSettings.working_time).indexOf( CheckBoxElements[i].querySelector(".ms-CheckBox-field").id ) > -1 ) {
				checkboxes[i].check();
			}
		}

		var ToggleElements = document.querySelectorAll(".ms-Toggle");
		for (var j = 0; j < ToggleElements.length; j++) {
			var switchId = ToggleElements[j].id;
			var activated;
			switch (switchId) {
				case 'smartdeco-switch' :
					activated = deconnexionSettings.smart_deconnexion_enabled;
					break;
				case 'night-switch' :
					activated = deconnexionSettings.is_enabled.nights;
					break;
				case 'weekend-switch' :
					activated = deconnexionSettings.is_enabled.weekends;
					break;
				case 'vacation-switch' :
					activated = deconnexionSettings.is_enabled.vacations;
					break;
			}

			toggles[switchId] = toggles[switchId] || new fabric['Toggle'](ToggleElements[j]);
			if (activated) {
				ToggleElements[j].querySelector(".ms-Toggle-field").setAttribute('class', 'ms-Toggle-field is-selected');
			}
		}

		var ButtonElement = document.querySelector("#deconnexion-button");
		button = button || new fabric['Button'](ButtonElement, function(event) {
		  	// Insert Event Here
			updateDeconnexionSettings();
		});

		$("#start-time select").val(deconnexionSettings.working_time && deconnexionSettings.working_time[Object.keys(deconnexionSettings.working_time)[0]][0]);
		$("#end-time select").val(deconnexionSettings.working_time && deconnexionSettings.working_time[Object.keys(deconnexionSettings.working_time)[0]][1]);

		if (deconnexionSettings.smart_deconnexion_set){
			$("#smartdeco-switch").removeClass('is-disabled');
			$("#smartdeco-switch").attr('title',"Activer pour couper automatiquement les notifications en dehors de vos horaires");
		}
	}

	function eventsHandlers(){
		$("#dashboard").click(function () {
			if (userStatus.email_validated){
				makeCorsRequestMagicLink();
			}
		});

		$(".product-card").click(function () {
			if ($(this).hasClass("active")){
				showPage("content-main");
			} else {
				getConsent($(this).attr('id'));
			}
		});

		/*$("#smartdeco").click(function () {
			if (deconnexionSettings.smart_deconnexion_set) {
				toggleSmartDeconnexion($(this).hasClass('is-selected'));
			}
		});

		$("#smartdeco-switch").click(function () {
			if (!deconnexionSettings.smart_deconnexion_set){
				if (productChoice === "product-smartdeconnexion"){
					withStatus(function(){
						buildSettingsPage();
						showPage("content-settings");
					});
				} else if (apiAvailable && (userStatus.personal_grant_company_policy || userStatus.personal_grant_company_policy === undefined)) {
					withStatus(function(){
						buildSettingsPage();
						showPage("product-choice");
					});
				}
			}
		});

		$("#settings").click(function () {
			var mainIndex = Number($("#content-main").css('z-index'));
			var settingsIndex = Number($("#content-settings").css('z-index'));
			var productsIndex = Number($("#product-choice").css('z-index'));
			var max = Math.max(mainIndex, settingsIndex, productsIndex)
			if (productsIndex === max) {
				showPage("content-main");
			} else if ( settingsIndex < max ) {
				withStatus(function(){
					if (productChoiceDisplay()){
						buildProductChoicePage();
						showPage("product-choice");
					} else {
						buildSettingsPage();
						showPage("content-settings");
					}
				});
			} else {
				showPage("content-main");
			}
		});*/

		$("#opposition").click(function () {
			revokeOpposition();
		});
	}

	function revokeOpposition(){
		var url = backendURI + '/api/v2/employees/' + userEmail + '/set_opposition_right';
	    var params = {"is_opposed": false};
	    sendCORSRequest('POST', url, oppositionCallback, params, function(){});
	}

	function getConsent(product) {
		if ( product !== "product-feedback") {
			var incremental_needed_scopes;
			if ( product === "product-smartdeconnexion" ) {
				incremental_needed_scopes = smartDeconnexionRequestedScopes;
			} else if ( product === "product-analytics" ) {
				incremental_needed_scopes = analyticsRequestedScopes;
			}
			var state = "mailoop";
			// var state = JSON.stringify({ "exchangeIdentityToken": exchangeIdentityToken, "exchangeCallbackToken": asyncResult.value });
			var consentUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize?client_id=15a9849e-dad7-44f1-abb6-ec70a9822043&response_type=code&state=' + state + '&redirect_uri=' + backendURI +'/synchronous_ms_personnal_access_grant&response_mode=query&scope='+ incremental_needed_scopes + '&domain_hint=organizations&login_hint=' + userEmail
            console.log("Getting_consent; ui.displayAsync/ No ApiAvailable check")
			Office.context.ui.displayDialogAsync(consentUrl, function (asyncResult) {
				if (asyncResult.status === "succeeded") {
					showPage("content-main");
				}
			});
		} else {
			productChoice = "product-feedback";
			_settings.set("productChoice", productChoice);
			_settings.saveAsync();
			showPage("content-main");
		}
	}

	function showPage(pageId){
		var mainIndex = Number($("#content-main").css('z-index'));
		var settingsIndex = Number($("#content-settings").css('z-index'));
		var productsIndex = Number($("#product-choice").css('z-index'));
		var oppositionIndex = Number($("#opposition").css('z-index'));
		$("#" + pageId).css({'z-index': Math.max(mainIndex, settingsIndex, productsIndex, oppositionIndex) + 1});
		setTimeout(function(){ 
			$("#loading").css({'z-index': 0});
		}, 300);
	}

	function hideHeader(){
		$("#header-overlay").css({"z-index": Number($("#content-header").css('z-index')) + 1});
		$("#header-overlay").css({"opacity": 0.6});
	}

	function showHeader(){
		$("#header-overlay").css({"z-index": Number($("#content-header").css('z-index')) - 1});
		$("#header-overlay").css({"opacity": 0});
	}

	function productChoiceDisplay(){
		var response = false;
		minimumDeconnexionScopes.split(' ').forEach(function(element){
			if (userStatus.currently_allowed_scopes.indexOf(element) === -1){
				response = true;
			}
		})
		return response;
	}

	function getProductChoiceFromStatus(){
		var missing = false;
		minimumDeconnexionScopes.split(' ').forEach(function(element){
			if (userStatus.currently_allowed_scopes.indexOf(element) === -1){
				missing = true;
			}
		});
		if (!missing){
			return "product-smartdeconnexion"; 
		} else {
			missing = false;
			minimumAnalyticsScopes.split(' ').forEach(function(element){
				if (userStatus.currently_allowed_scopes.indexOf(element) === -1){
					missing = true;
				}
			});
			if (!missing){
				return "product-analytics"; 
			} else {
				return undefined;
			}
		}
	}

	function toggleSmartDeconnexion(isActivated){
		var url = backendURI + '/api/v2/employees/' + userEmail + '/set_smart_deconnexion_enabled';
		var params = {'smart_deconnexion_enabled': isActivated};
		sendCORSRequest('POST', url, toggleSmartDeconnexionCallback, params);
	}

	function toggleSmartDeconnexionCallback(json_response){
		if ( json_response.smart_deconnexion_enabled !==  $("#smartdeco").hasClass("is-selected")) { 
			$("#smartdeco").toggleClass("is-selected");
		}
		deconnexionSettings = buildDeconnexionSettings(json_response);
		_settings.set('deconnexionSettings', deconnexionSettings);
		_settings.saveAsync();
	}

	function withSmartDeconnexionSettings(callback){
		var url = backendURI + '/api/v2/employees/' + userEmail + '/status';
		sendCORSRequest('GET', url, deconnexionSettingsCallback, {}, callback);
	}

	function deconnexionSettingsCallback(json_response, callback){
		deconnexionSettings = buildDeconnexionSettings(json_response);
		_settings.set('deconnexionSettings', deconnexionSettings);
		_settings.saveAsync();
		buildSettingsPage();
		callback();
	}

	function buildDeconnexionSettings(json) {
		if ( json.smart_deconnexion_set ) {
			return ({
				'country': json.country,
				'time_zone': json.time_zone,
				'smart_deconnexion_enabled': json.smart_deconnexion_enabled,
				'is_enabled': {
					'nights': json.is_deconnexion_enabled_at_nights,
					'weekends': json.is_deconnexion_enabled_at_weekends,
					'vacations': json.is_deconnexion_enabled_at_vacations
				},
				'working_time': json.working_time,
				'smart_deconnexion_set': json.smart_deconnexion_set
			});
		} else {
			return defaultDeconnexionSettings;
		}
	}

	function buildDeconnexionSettingsFromForm() {
		var activated = $('.ms-Toggle-field.is-selected').map(function(){ return this.id; }).get();
		return ({
			'country': 'FR',
			'time_zone': 'Europe/Paris',
			'smart_deconnexion_enabled': activated.indexOf('smartdeco') > -1,
			'is_enabled': {
				'nights': activated.indexOf('nights') > -1,
				'weekends': activated.indexOf('weekends') > -1,
				'vacations': activated.indexOf('vacations') > -1,
			},
			'working_time': buildWorkingTime(),
		});
	}

	function buildWorkingTime(){
		var workingTime = {};
		var startTime = $("#start-time select").val();
		var endTime = $("#end-time select").val();
		$(".ms-CheckBox-field.is-checked").map(function(){
			workingTime[Number(this.id)] = [parseFloat(startTime), parseFloat(endTime)];
		});
		return workingTime;
	}

	function updateDeconnexionSettings() {
		var url = backendURI + '/api/v2/employees/'+ userEmail + '/deconnexion_settings';
		var params = buildDeconnexionSettingsFromForm();
		deconnexionSettings = buildDeconnexionSettingsFromForm();
		sendCORSRequest('POST', url, saveDeconnexionSettings, params, function(){});
	}

	function saveDeconnexionSettings(json_response){
		deconnexionSettings['smart_deconnexion_set'] = json_response.smart_deconnexion_set;
		if (json_response.smart_deconnexion_set) { $("#smartdeco-switch").removeClass('is-disabled'); }
		_settings.set('deconnexionSettings', deconnexionSettings);
		_settings.saveAsync();
		showPage("content-main");
	}



	// Validate if addin should be activated and vote possible
	function testIfPlugInActive() {
		// Order in the if is very important to manage all cases
		if  (item.itemClass.indexOf("IPM.Appointment") !== -1 ) {
			if ( meetingVotable() ) {
				item.loadCustomPropertiesAsync(customPropsCallback);
			}
		} else if ( messageVotable() ) {
			item.loadCustomPropertiesAsync(customPropsCallback);
		}
	}

	function meetingVotable() {
		var now = new Date();
		var min_date = new Date();
		// validation on email type
		min_date.setMonth(min_date.getMonth() - 1);
		var end = item.end;

		if (end < min_date){
			$('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Vous ne pouvez pas exprimer un feedback sur une réunion datant de plus d\'un mois.</p>');
			return false;
		} else if ( now < end ) {
			$('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Vous ne pouvez pas exprimer un feedback sur une réunion qui n\'est pas encore terminée.</p>');
			return false;
		} else if ( item.requiredAttendees.length < 2 ) {
			$('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Vous ne pouvez pas exprimer un feedback sur une réunion où vous semblez être le seul participant.</p>');
			return false;
		} else {
			return true;
		}
	}

	function messageVotable() {
		var min_date = new Date();
		// validation on email type
		min_date.setMonth(min_date.getMonth() - 1);
		var mail_date = item.dateTimeCreated;
		var from = item.from.emailAddress;
		var from_domain = from.replace(/.*@/, "");
		var user = mailbox.userProfile.emailAddress;
		if (mail_date < min_date) {
			// No vote for emails older than a month
			$('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Vous ne pouvez pas exprimer un feedback sur un e-mail datant de plus d\'un mois.</p>');
			return false;
		} else if (from === user) {
			// No vote for emails sent to user himself
			$('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Vous ne pouvez pas exprimer de feedbacks sur vos propres e-mails !</p>');
			return false;
		} else if ( internalDomains && internalDomains.indexOf(from_domain) > -1 ) {
			// Email is 'votable' --> load voted behaviors for the current item
			perimeter = "internal";
			return true;
		} else {
			// Email is considered as external
			perimeter = "external";
			return true;
		}
	}

  // Callback function from loading custom properties
  function customPropsCallback(asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          // Handle the failure.
      } else {
          // Successfully loaded custom properties,
          // can get them from the asyncResult argument.
          _props = asyncResult.value;
          voted_behaviors = _props.get('voted_behaviors');
          if (voted_behaviors === undefined) { voted_behaviors = {}; }
          fillTaskpaneWithBehaviors();
      }
  }

	// Fill the taskpane with votable behaviors UI
	function fillTaskpaneWithBehaviors(){
		$('#content-main').empty();
		var listBehaviors;
		// Manage different type of messages : message, meeting_request, meeting_response
		var behaviors = _settings.get('behaviors');
		if (item.itemClass === "IPM.Note") {
			if (perimeter === "internal") {
				listBehaviors = behaviors.filter(function(behavior) { return ( behavior.category === "message" && ["MIXED","INTERNAL"].indexOf(behavior.perimeter) > -1 ); });
			} else {
				listBehaviors = behaviors.filter(function(behavior) { return ( behavior.category === "message" && ["MIXED","EXTERNAL"].indexOf(behavior.perimeter) > -1 ); });
			}
		} else if (item.itemClass === "IPM.Schedule.Meeting.Request") {
			listBehaviors = behaviors.filter(function(behavior) { return ( behavior.category === "message" ); });
		} else if (item.itemClass.indexOf("IPM.Schedule.Meeting.Resp") !== -1 ) {
			listBehaviors =  behaviors.filter(function(behavior) { return ( behavior.category === "message" ); });
		} else if (item.itemClass.indexOf("IPM.Appointment") !== -1 ) {
			listBehaviors =  behaviors.filter(function(behavior) { return ( behavior.category === "meeting" ); });		}

		if ((listBehaviors === undefined) || (behaviors.length === 0)){
		  // Manage case where no behaviors have been chosen by the company
		  $('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Il semble qu\'aucun feedback n\'aient été sélectionnés pour ce type de messages.</p>');
		} else {
		  // Fill the taskpane with votable behaviors from roamingSettings (_settings) and vote status from item customProperties (_props)
		
		var translate = {"0": "neutral", "1": "positive", "-1": "negative"};
		$('#content-main').append('<h1 class="ms-font-l" id="subtitle">Partagez vos feedbacks</h1>');
		["positive", "neutral", "negative"].forEach(function(key) {
			renderBehaviorsList(key, listBehaviors.filter(function(behavior) { return ( translate[behavior.family.toString()] === key ); } ));
		});

	  	// Action triggered on behavior item click
	  	$('.vote-element').click(function () {
	      	if ($('> .check', this).hasClass("voted")) {
	        	makeCorsRequestDeleteVote($('> .check', this).attr('id'));
	      	} else {
	        	makeCorsRequestCreateVote($(this).attr('id'));
	      	}
	  	});
		}
	}

function renderBehaviorsList(family, list) {
	$('#content-main').append('<div class="emoji-family"><i class="icon ms-Icon ms-Icon--' + msIconNames[family] + '" aria-hidden="true"></i></div>');
	list.forEach(function (value) {
    	$('#content-main').append('<div title="' + value.description + '" class="vote-element ' + family + '" id="' + value.behavior_id + '"><img src= "https://mailoop.blob.core.windows.net/assets/' + value.ref_name + '_64.png" alt=""><p class="ms-font-m">' + value.name + '</p><div id="' + ( voted_behaviors[value.behavior_id] || "" ) + '" class="check' + ( voted_behaviors[value.behavior_id] ? " voted" : "" ) + '"><i class="icon ms-Icon ms-Icon--CheckMark" aria-hidden="true"></i></div></div>');
 	});
}



//Success callback for AllBehaviors update - save all behaviors in _settings by category
function behaviorsAllCallback(json_response, callback){
	behaviors = [];
	json_response.forEach( function(rawBehavior) {
		behaviors.push({
			ref_name: rawBehavior.ref_name,
			name: rawBehavior.translations.fr.name,
			description: rawBehavior.translations.fr.description,
			perimeter: rawBehavior.perimeter,
			priority: rawBehavior.priority,
			family: rawBehavior.family,
			category: rawBehavior.category,
			behavior_id: rawBehavior.id,
		});
	});
	_settings.set('behaviors', behaviors);
	behaviorsPromise = true;
	_settings.saveAsync();
	callback();
}

//Success callback for AllBehaviors update - save all behaviors in _settings by category
function domainsCallback(json_response, callback){
	internalDomains = [];
	json_response.forEach( function(rawDomain) {
		internalDomains.push(rawDomain.url);
	});
	_settings.set('internalDomains', internalDomains);
	_settings.saveAsync();
	domainsPromise = true;
	callback();
}

//Success callback for AllBehaviors update - save all behaviors in _settings by category
function statusCallback(json_response, callback){
	deconnexionSettings = buildDeconnexionSettings(json_response);
	userStatus = json_response;
	_settings.set('userStatus', userStatus);
	_settings.set('deconnexionSettings', deconnexionSettings);
	_settings.saveAsync();
	statusPromise = true;
	callback();
}

function oppositionCallback(json_response){
	showHeader();
	userStatus = json_response;
	_settings.set('userStatus', userStatus);
	_settings.saveAsync();
	showPage("content-main");
}

// Make the actual CORS request to vote for a specific behavior.
function makeCorsRequestCreateVote(behavior_id) {
    var url;
    var params;
    if ( item.itemClass === 'IPM.Appointment' ) {
    	url = backendURI + '/api/v2/ms_meeting_votes';
    	params =  meetingPayload(behavior_id);
    } else {
    	url = backendURI + '/api/v2/votes';
    	params =  messagePayload(behavior_id);
    }
    $('#' + behavior_id + '.vote-element > .check').addClass('voted');
    sendCORSRequest('POST', url, voteCreateCallback, params, function(){});
}

function messagePayload(behavior_id){
	return { vote:
				{  	
					behavior_id: behavior_id,
		        	email: {
		          		internet_message_id: item.internetMessageId,
		          		from: item.from.emailAddress,
		          		date: item.dateTimeCreated,
		        	}
				}
			};
}

function meetingPayload(behavior_id){
	return { vote:
				{  	
					behavior_id: behavior_id,
		        	meeting: {
		        		mailbox_meeting_id: item.itemId,
		          		organizer: item.organizer.emailAddress,
		          		required_attendees: item.requiredAttendees.map(function(attendee){ return attendee.emailAddress; }),
		          		start: item.start,
		          		end: item.end,
		          		created: item.dateTimeCreated
		        	}
				}
			};
}

// Make the actual CORS request to vote for a specific behavior.
function makeCorsRequestDeleteVote(vote_id) {
    var url = backendURI + '/api/v2/votes/' + vote_id;
    var params = {};
    $('#' + vote_id + '.check').removeClass('voted');
    sendCORSRequest('DELETE', url, voteDeleteCallback, params, function(){});
}


// Success callback after Vote create
function voteCreateCallback(json_response, callback) {

	// Manage normal behaviors vote
	$('#' + json_response.behavior_id + '.vote-element > .check').attr('id', json_response.id);
	voted_behaviors[json_response.behavior_id] = json_response.id;
	// Save to _props the vote status
	_props.set("voted_behaviors", voted_behaviors);
	_props.saveAsync();
}

// Success callback after Vote destroy
function voteDeleteCallback(json_response, callback) {
    // Manage normal behaviors vote
    $('#' + json_response.id + '.check').removeClass('voted').attr('id', '');
    delete voted_behaviors[json_response.behavior_id];
    // Save to _props the vote status
    _props.set("voted_behaviors", voted_behaviors);
    _props.saveAsync();
  }

  // Success callback after Vote destroy
function vote404Callback(id) {
    // Manage normal behaviors vote
    var behavior_id = $('#' + id).parent().attr('id');
    $('#' + id + '.check').removeClass('voted').attr('id', '');
    delete voted_behaviors[behavior_id];
    // Save to _props the vote status
    _props.set("voted_behaviors", voted_behaviors);
    _props.saveAsync();
}


  function makeCorsRequestMagicLink(){
    var url = backendURI + '/api/v2/magic_link';
    var params =  {};
    sendCORSRequest('GET', url, magicLinkCallback, params, function(){});
  }

function magicLinkCallback(json_response){
    window.open(dashboardURI + '/?email=' + json_response.magic_link.email + '&temporary_password=' + json_response.magic_link.temporary_password);
}

  function saveCallback(asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Failed) {
          console.log(asyncResult.error.message);
      } else {
          // Async call to save custom properties completed.
          // Proceed to do the appropriate for your add-in.
      }
  }

  function isPersistenceSupported() {
    // This feature is part of the preview 1.5 req set
    // Since 1.5 isn't fully implemented, just check that the
    // method is defined.
    // Once 1.5 is implemented, we can replace this with
    // Office.context.requirements.isSetSupported('Mailbox', 1.5)
    return mailbox.addHandlerAsync !== undefined;
  }

  // Auto load when addin is pinned on Outlook desktop
  function loadNewItem(eventArgs) {
      testIfPlugInActive();
  }

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
      xhr.setRequestHeader("X-User-Email", userEmail);
      xhr.setRequestHeader("X-User-Exchange-Identity-Token", exchangeIdentityToken);
      return xhr;
  }


  // Send the request and manage callback with token validated
  function sendCORSRequest(method, url, successCallback, params, callbackCallback) {
    if (method == 'GET' && Object.keys(params).length > 0 ) {
        url += '?' + Object.keys(params).map(function(k) { return encodeURIComponent(k) + "=" + encodeURIComponent(params[k]);}).join('&');
    }

    var xhr = createCORSRequest(method, url);

    if (!xhr) {
        alert('CORS not supported');
        return;
    }

    xhr.onload = function () {
        if ( xhr.status === 404 && ( url.indexOf("/api/v2/votes/") > -1 ) && ( method === "DELETE" ) ) {
      		vote404Callback(url.match(/(\d+)$/)[0]);
      	} else if ((xhr.status !== 200) && (xhr.status !== 201) && (xhr.status !== 204) && (xhr.status !== 202))  {
          	connexionErrorMessage(xhr.status);
      	} else {
          successCallback(JSON.parse(xhr.responseText), callbackCallback);
        }
    };
    xhr.onerror = function () {
      	connexionErrorMessage(xhr.status);
	};

    xhr.send(JSON.stringify(params));
  }

  // Print connexion error message
  function connexionErrorMessage(status) {
      $('#content-main').empty().append('<p class="ms-font-l ms-fontWeight-light comment">Impossible de se connecter au serveur Mailoop. Fermez et relancez l\'addin. Contactez-nous si le problème persiste.</p>');
  }

  function getItemRestId() {
	  	if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
	    	// itemId is already REST-formatted
	    	return item.itemId;
	  	} else {
	    	// Convert to an item ID for API v2.0
	    	return Office.context.mailbox.convertToRestId(
	      		item.itemId,
	      		Office.MailboxEnums.RestVersion.v2_0
	    	);
	  	}
	}

	function testApiAvailability(accessToken, callback) {
	  	// Get the item's REST ID
	  	var itemId = getItemRestId();

	  	// Construct the REST URL to the current item
	  	var getMessageUrl = mailbox.restUrl + '/v2.0/me/messages/' + itemId;
	  	var url = backendURI + '/api/v2/employees/' + userEmail + '/api_available';
	  	$.ajax({
	    	url: getMessageUrl,
	    	dataType: 'json',
	    	headers: { 'Authorization': 'Bearer ' + accessToken }
	  	}).done(function(item){
	    	// Message is passed in `item`
	    	apiAvailable = true;
	    	apiTestPromise = true;
	    	_settings.set("apiAvailable", apiAvailable);
	    	_settings.saveAsync();
	    	sendCORSRequest("POST", url, function(callback){}, {status: true}, function(){});
	    	callback();
	  	}).fail(function(error){
	    	// Handle error
	    	apiAvailable = false;
	    	apiTestPromise = true;
	    	_settings.set("apiAvailable", apiAvailable);
	    	_settings.saveAsync();
	  		sendCORSRequest("POST", url, function(callback){}, {status: false}, function(){});
	  		callback();
	  	});
	}


})();

// MIT License:

// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// ""Software""), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:

// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED ""AS IS"", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
