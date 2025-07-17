/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    }
});

export async function run() {
    /**
     * Inserisci qui il tuo codice per interagire con l'elemento di Outlook.
     * Ad esempio, per aggiungere del testo al corpo dell'email:
     */
    const item = Office.context.mailbox.item;
    if (item) {
        // Esempio: Aggiunge del testo al corpo dell'email in composizione
        item.body.setSelectedDataAsync("Hello from your add-in!", { coercionType: Office.CoercionType.Text },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                    document.getElementById("output").innerText = "Testo aggiunto al corpo dell'email.";
                } else {
                    document.getElementById("output").innerText = "Errore nell'aggiungere il testo: " + asyncResult.error.message;
                }
            }
        );
    }
}