// ==UserScript==
// @name         XHR Interceptor
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  try to take over the world!
// @author       You
// @match        https://tasks.office.com/*
// @grant        none
// ==/UserScript==

window.HEADERS = new Headers();
window.HEADERS.append("x-planner-clienttype", "StandAloneWebApp");
window.HEADERS.append("x-plannerclienttype", "Plex");
window.HEADERS.append("x-requested-with", "XMLHttpRequest");
window.HEADERS.append("x-planner-usage", "planboard");


// Reasign the existing setRequestHeader function to
// something else on the XMLHtttpRequest class
XMLHttpRequest.prototype.wrappedSetRequestHeader =
  XMLHttpRequest.prototype.setRequestHeader;

// Override the existing setRequestHeader function so that it stores the headers
XMLHttpRequest.prototype.setRequestHeader = function(header, value) {
    // Call the wrappedSetRequestHeader function first
    // so we get exceptions if we are in an erronous state etc.
    this.wrappedSetRequestHeader(header, value);

    // Create a headers map if it does not exist
    /*if(!this.headers) {
        this.headers = {};
    }

    // Create a list for the header that if it does not exist
    if(!this.headers[header]) {
        this.headers[header] = [];
    }

    // Add the value to the header
    this.headers[header].push(value);*/

    window.HEADERS.set(header, value);
}
