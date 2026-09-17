/**
* @fileoverview This Google Sheets script will provide Canvas API functionality for other scripts.
* @author james@richland.edu [James Jones]
* @author drcalhoun
* @license Copyright 2015 Standard ISC License
* @OnlyCurrentDoc
*/

/**
* @function This will test the API settings and try to determine the name of
*           the user
*/
function checkApiSettings() {
  var name;
  try {
    var props = getApiSettings();
    if (props === false) {
      return false;
    }
    var ui = SpreadsheetApp.getUi();
    var profile = canvasAPI('GET /api/v1/users/self/profile');
    if (typeof profile === 'undefined') {
      var mesg = 'Unable to connect to ' + props.host;
      ui.alert('Failure', mesg, ui.ButtonSet.OK);
      throw new Error(mesg);
    }
    name = profile.name ? profile.name : 'Unknown User';
    ui.alert('Success!', 'Connected to ' + props.host + ' as ' + name + '\nYou may now use the API calls.', ui.ButtonSet.OK);
  } catch (e) {
    Logger.log(e);
    if (typeof showError === 'function') {
      showError('API Check Error', e);
    }
    return;
  }
  return name;
}

/**
* @function This opens the Configuration Dialog. 
* This should be added to your menu
*/
function configurationDialog() {
  // Removed the deprecated .setSandboxMode() method
  var html = HtmlService.createTemplateFromFile('canvasConfig').evaluate();
  var props = SpreadsheetApp.getUi().showModalDialog(html, 'Canvas API Configuration');
  return;
}

/**
* @function Process the submission from the Configuration form
* @param {Object}
*          formObject - the data returned by the form submission
*/
/**
* @function Process the submission from the Configuration form
*/
function processConfigurationForm(formObject) {
  var userProperties = PropertiesService.getUserProperties();
  var token = formObject.canvas_token;
  
  // 1. Save the token immediately so a bad host doesn't abort the save
  if (typeof token !== 'undefined' && token && token.trim() !== '') {
    userProperties.setProperty('token', token.trim());
  }

  // 2. Validate the host safely
  try {
    var host = determineCanvasHost(formObject.canvas_host);
    if (host !== false) {
      if (host != '') {
        userProperties.setProperty('host', host);
      } else {
        userProperties.deleteProperty('host');
      }
    }
    checkApiSettings();
  } catch(e) {
    Logger.log(e);
    if (typeof showError === 'function') {
      showError('Configuration Error', e);
    }
  }
  return;
}

/**
* @function Tries to parse the URL and determine the Canvas Host
* @param {String}
*          text - The value to parse
* @returns
*/
function determineCanvasHost(text) {
  if (typeof text === 'undefined') {
    return false;
  }
  text = text.toLowerCase().trim();
  if (text == '') {
    return '';
  }
  try {
    var ui = SpreadsheetApp.getUi();
    var hostRegex = new RegExp(
      '^(?:https:\/\/)?(?:([a-z0-9][-a-z0-9]*$)|(?:([a-z0-9][-a-z0-9]*(?:[.][a-z0-9][-a-z0-9]*)+)(?:\/|$)))',
      'i');
    var match = hostRegex.exec(text);
    if (match != null) {
      var value = typeof match[1] === 'undefined' ? match[2] : match[1];
      if (!/[.]/.test(value)) {
        value += '.instructure.com';
      }
      if (value != text) {
        var alertResponse = ui.alert('Notice', 'Using "' + value + '" for your Canvas Hostname.\n If this is not correct, then click Cancel.', ui.ButtonSet.OK_CANCEL);
        if (alertResponse == ui.Button.CANCEL) {
          text = false;
        } else {
          text = value;
        }
      }
    } else {
      text = false;
      ui.alert(
        'Error',
        'Sorry, I did not understand what you entered for the Canvas Hostname',
        ui.ButtonSet.OK);
      throw new Error('Bad Hostname');
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble to UI
  }
  return text;
}

/**
* @function This function fetches your Canvas instance and your access token.
*/
function getApiSettings() {
  var required_properties = [ 'host', 'token' ];
  var properties;
  try {
    var userProperties = PropertiesService.getUserProperties();
    properties = userProperties.getProperties();
    var openDialog = false;
    for (var i = 0; i < required_properties.length; i++) {
      if (typeof properties[required_properties[i]] === 'undefined') {
        openDialog = true;
      }
    }
    if (openDialog) {
      configurationDialog();
      return false;
    }
    var valid = true;
    if (typeof properties.host === 'undefined' || properties.host == '') {
      valid = false;
    }
    if (typeof properties.token === 'undefined' || properties.token == '') {
      valid = false;
    }
    if (!valid) {
      properties = false;
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Prevents the 'undefined === false' bug
  }
  return properties;
}

/**
* @function reset user properties
*/
function resetApiSettings() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
}

/**
* @function convert the parameters into a Query String
*/
function makeQueryString(obj) {
  var q = [];
  var j;
  for ( var i in obj) {
    if (obj.hasOwnProperty(i)) {
      var item = obj[i];
      if (typeof item == 'object') {
        if (Array.isArray(item)) {
          for (j = 0; j < item.length; j++) {
            q.push(i + '[]=' + item[j]);
          }
        } else {
          for (j in item) {
            if (item.hasOwnProperty(j)) {
              q.push(i + '[' + j + ']=' + item[j]);
            }
          }
        }
      } else {
        q.push(i + '=' + item);
      }
    }
  }
  return q.join('&');
}

/**
* @function Maps a Canvas API field name to a friendlier label for error messages.
* @param {String} field - the raw field name from Canvas (e.g. "due_at")
* @returns {String}
*/
function friendlyFieldName(field) {
  var labels = {
    'due_at': 'Due Date',
    'unlock_at': 'Available From',
    'lock_at': 'Available Until',
    'points_possible': 'Points Possible',
    'published': 'Published',
    'show_correct_answers_at': 'Show Answers Date',
    'hide_correct_answers_at': 'Hide Answers Date',
    'invalid_record': 'Canvas Validation'
  };
  return labels[field] || field;
}

/**
* @function This function calls the CanvasAPI and returns any information as an object.
*/
function canvasAPI(endpoint, opts, filter) {
  if (typeof endpoint === 'undefined') {
    return;
  }
  if (typeof opts === 'undefined') {
    opts = {};
  }
  
  var PER_PAGE = 100;
  var API_VERSION = 1;
  
  var endpointRegex = /^(GET|POST|PUT|DELETE|HEAD)\s+(.*)$/i;
  var tokenRegex = new RegExp('^:([a-z_]+)$');
  var data;
  
  try {
    var userProperties = getApiSettings();
    if (userProperties === false) {
      throw new Error('You need to specify a full set of credentials.');
    }
    
    // Mute HTTP exceptions to properly handle non-200 responses
    var parms = { 
      'headers' : { 'Authorization' : 'Bearer ' + userProperties.token },
      'muteHttpExceptions' : true
    };
    
    if (typeof endpoint !== 'string') {
      throw new Error('Endpoint specification must be a string. Received: ' + typeof endpoint);
    }
    var endpointMatches = endpointRegex.exec(endpoint);
    if (endpointMatches === null) {
      throw new Error('Invalid endpoint specified: ' + endpoint);
    }
    parms.method = endpointMatches[1].toLowerCase();
    var routes = endpointMatches[2].split('/');
    var route = [];
    var j;
    for (var i = 0; i < routes.length; i++) {
      if (routes[i]) {
        if (route.length == 0 && routes[i] != 'api') {
          route.push('api');
          route.push('v' + API_VERSION);
          i++;
        }
        var matches = tokenRegex.exec(routes[i]);
        if (matches !== null) {
          if (typeof opts !== 'object') {
            throw new Error('Options is not an object but variable substitutions is needed');
          }
          var tokens = [ routes[i], matches[1], ':sis_' + matches[1],
              'sis_' + matches[1] ];
          var tokenMatch = false;
          j = 0;
          while (!tokenMatch && j < tokens.length) {
            var token = tokens[j++];
            if (typeof opts[token] !== 'undefined') {
              tokenMatch = true;
              route.push(opts[token]);
              opts[token] = null;
            }
          }
          if (!tokenMatch) {
            throw new Error('Unable to find substitution for :' + matches[1] + ' in ' + endpointMatches[2]);
          }
        } else {
          route.push(routes[i]);
        }
      }
    }
    var payload = {};
    if (typeof opts === 'object') {
      for ( var field in opts) {
        if (opts.hasOwnProperty(field) && opts[field] !== null && !/^:/.test(field)) {
          payload[field] = opts[field];
        }
      }
    }
    var url = 'https://' + userProperties.host + '/' + route.join('/');
    data = [];
    if (parms.method == 'get' || parms.method == 'delete') {
      if (parms.method == 'get' && !payload.per_page) {
        payload.per_page = PER_PAGE;
      }
      var queryStr = makeQueryString(payload);
      if (queryStr) {
        url += '?' + queryStr;
        url = encodeURI(url);
      }
    } else if (parms.method == 'post' || parms.method == 'put') {
      parms.payload = payload;
    }
    while (url !== null) {
      var response = UrlFetchApp.fetch(url, parms);
      url = null;
      var code = response.getResponseCode();
      
      // Accept 201 Created and 204 No Content alongside 200 OK
      if (code >= 200 && code < 300) {
        var headers = response.getAllHeaders();
        if (parms.method == 'get' && typeof headers.Link !== 'undefined') {
          var links = headers.Link.split(',');
          if (typeof links === 'object') {
            for (var l = 0; l < links.length; l++) {
              // Resilient pagination regex
              var linkMatch = /<([^>]+)>;\s*rel=["']?next["']?/i.exec(links[l]);
              if (linkMatch !== null) {
                url = linkMatch[1];
              }
            }
          }
        }
        
        // Ensure content exists (protects against 204 Empty Responses)
        var content = response.getContentText();
        if (content) {
          var json = JSON.parse(content);
          var key;
          if (typeof json === 'object') {
            if (Array.isArray(json)) {
              for ( var item in json) {
                if (json.hasOwnProperty(item)) {
                  var entry = json[item];
                  if (typeof filter !== 'undefined') {
                    var row = {};
                    for (j in filter) {
                      if (filter.hasOwnProperty(j)) {
                        key = filter[j];
                        if (typeof entry[key] !== 'undefined') {
                          if (Array.isArray(entry[key])) {
                            row[key] = entry[key].slice(0);
                          } else {
                            row[key] = entry[key];
                          }
                        }
                      }
                    }
                    data.push(row);
                  } else {
                    data.push(entry);
                  }
                }
              }
            } else {
              if (typeof filter === 'undefined') {
                data = json;
              } else {
                for (j in filter) {
                  if (filter.hasOwnProperty(j)) {
                    key = filter[j];
                    if (typeof json[key] !== 'undefined') {
                      if (Array.isArray(json[key])) {
                        data[key] = json[key].slice(0);
                      } else {
                        data[key] = json[key];
                      }
                    }
                  }
                }
              }
            }
          }
        }
      } else {
        // Advanced Error Parsing: Extract clean message if Canvas sends JSON
        var content = response.getContentText();
        var cleanMessage = '';
        try {
                    var errObj = JSON.parse(content);
          if (errObj.errors) {
            if (Array.isArray(errObj.errors)) {
              // Form 1: errors is an array of { message: ... }
              cleanMessage = errObj.errors.map(function(e){ return e.message; }).join('\n');
            } else if (typeof errObj.errors === 'object') {
              // Form 2: errors is an object keyed by field name, e.g.
              // { "due_at": [ { attribute, message, type } ], "invalid_record": [ ... ] }
              var fieldMessages = [];
              for (var errField in errObj.errors) {
                if (errObj.errors.hasOwnProperty(errField) && errField !== 'invalid_record') {
                  var fieldErrors = errObj.errors[errField];
                  if (Array.isArray(fieldErrors)) {
                    for (var fe = 0; fe < fieldErrors.length; fe++) {
                      var feMessage = fieldErrors[fe] && fieldErrors[fe].message ? fieldErrors[fe].message : JSON.stringify(fieldErrors[fe]);
                      fieldMessages.push(friendlyFieldName(errField) + ': ' + feMessage);
                    }
                  }
                }
              }
              // If every field was "invalid_record" (no more specific field errors found), fall back to it
              if (fieldMessages.length === 0 && errObj.errors.invalid_record) {
                var invalidRecord = errObj.errors.invalid_record;
                if (Array.isArray(invalidRecord)) {
                  for (var ir = 0; ir < invalidRecord.length; ir++) {
                    fieldMessages.push(friendlyFieldName('invalid_record') + ': ' + (invalidRecord[ir] && invalidRecord[ir].message ? invalidRecord[ir].message : JSON.stringify(invalidRecord[ir])));
                  }
                }
              }
              cleanMessage = fieldMessages.join('\n');
            }
          } else if (errObj.message) {
             cleanMessage = errObj.message;
          }
        } catch(parseError) {
           // Fallback if the response isn't JSON
        }
        
        if (cleanMessage) {
            throw new Error('Canvas API Error: ' + cleanMessage);
        } else {
            throw new Error('HTTP ' + code + '\n' + content);
        }
      }
    }
  } catch (e) {
    Logger.log(e);
    throw e; // Bubble the error up to the UI-level functions
  }
  return data;
}