(function () {
    // Office.onReady viene chiamato quando Office.js è completamente caricato.
    Office.onReady(function (info) {
        if (info.host === Office.HostType.Outlook) {
            console.log("Dialog Office.js is ready.");

            // Aggiungi event listener ai bottoni
            document.getElementById("confirmBtn").addEventListener("click", function () {
                console.log("Confirm button clicked in dialog.");
                // Invia un messaggio all'add-in principale che l'utente ha confermato.
                Office.context.ui.messageParent(JSON.stringify({ confirmSend: true }));
            });

            document.getElementById("cancelBtn").addEventListener("click", function () {
                console.log("Cancel button clicked in dialog.");
                // Invia un messaggio all'add-in principale che l'utente ha annullato.
                Office.context.ui.messageParent(JSON.stringify({ confirmSend: false }));
            });
        }
    });
})();