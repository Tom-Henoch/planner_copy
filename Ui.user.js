// ==UserScript==
// @name         Ui
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  try to take over the world!
// @author       You
// @match        https://tasks.office.com/FR/Home/Planner
// @grant       GM_getResourceText
// @grant       GM_addStyle
// @require      https://code.jquery.com/jquery-3.5.1.min.js
// @require      https://code.jquery.com/ui/1.12.1/jquery-ui.min.js
// @resource       jQueryUi  https://cdnjs.cloudflare.com/ajax/libs/jqueryui/1.12.1/jquery-ui.min.css
// ==/UserScript==

const CSS_jQueryUi = GM_getResourceText('jQueryUi');
GM_addStyle(CSS_jQueryUi);

const HTML = `
<div>
<p>Selectionner une équipe puis un plan dans celle-ci, pour copier le plan actuel</p>
<p><label for="groups">Équipe: </label>
<select id="groups"></select></p>

<p><label for="plans">Plan: </label>
<select id="plans"></select></p>


<p><button id="copy">Copier</button></p>
</div>
`

$(() => {

    $('body').append('<div id="dialog" title="Copie Planificateur"></div>');

    $( "#dialog").dialog({modal: true}).html(HTML);


});