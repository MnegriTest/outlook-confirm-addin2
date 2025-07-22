// La funzione onMessageSendHandler è il punto di ingresso per l'evento OnMessageSend.
// Viene registrata nel manifest.xml.
Office.addin.onMessageSend(function (event) {
    const item = Office.context.mailbox.item;
    const targetRecipient = "github@mario.it"; // **L'INDIRIZZO EMAIL DA MONITORARE**

    console.log("OnMessageSend event triggered.");

    // Recupera i destinatari in modo asincrono.
    item.recipients.getAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const recipients = asyncResult.value;
            let foundSpecificRecipient = false;

            for (let i = 0; i < recipients.length; i++) {
                if (recipients[i].emailAddress.toLowerCase() === targetRecipient.toLowerCase()) {
                    foundSpecificRecipient = true;
                    break;
                }
            }

            if (foundSpecificRecipient) {
                console.log(`Specific recipient found: ${targetRecipient}. Opening confirmation dialog.`);
                // Apre il dialog di conferma. L'URL è quello definito nel manifest (Dialog.Url).
                Office.context.ui.displayDialogAsync(
                    Office.context.addin.get  Office.context.dialog.url, // Usa la proprietà di Office.context.dialog.url
                    { height: 45, width: 45, displayInIframe: true }, // displayInIframe è utile per il debug
                    function (dialogAsyncResult) {
                        if (dialogAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                            let dialog = dialogAsyncResult.value;
                            console.log("Dialog opened successfully.");
                            // Aggiunge un handler per i messaggi ricevuti dal dialog.
                            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
                                console.log("Message received from dialog:", args.message);
                                dialog.close(); // Chiudi il dialog dopo aver ricevuto la risposta.
                                const messageData = JSON.parse(args.message);
                                const confirmSend = messageData.confirmSend;

                                if (confirmSend) {
                                    console.log("User confirmed send. Allowing event.");
                                    event.completed({ allowEvent: true }); // L'utente ha confermato, permetti l'invio.
                                } else {
                                    console.log("User cancelled send. Blocking event.");
                                    // Aggiunge una notifica all'utente che l'invio è stato annullato.
                                    item.notificationMessages.addAsync("SendCancelled", {
                                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                                        message: "L'invio dell'email è stato annullato.",
                                        icon: "Icon.80x80", // Assicurati che l'icona esista e sia referenziata nel manifest
                                        persistent: false // Non persistente, scompare automaticamente
                                    });
                                    event.completed({ allowEvent: false }); // L'utente ha annullato, blocca l'invio.
                                }
                            });
                            // Aggiunge un handler per la chiusura del dialog dall'utente (es. cliccando la X).
                            dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
                                if (args.error.code === 12006) { // 12006 è il codice per la chiusura del dialog da parte dell'utente
                                    console.log("Dialog closed by user. Blocking event.");
                                    item.notificationMessages.addAsync("SendCancelledByDialog", {
                                        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                                        message: "L'invio è stato annullato perché hai chiuso la finestra di conferma.",
                                        icon: "Icon.80x80",
                                        persistent: false
                                    });
                                    event.completed({ allowEvent: false }); // Blocca l'invio se il dialog è chiuso
                                } else {
                                    console.error("Dialog error:", args.error.message);
                                    event.completed({ allowEvent: true }); // Permetti l'invio in caso di altri errori
                                }
                            });
                        } else {
                            // Errore nell'apertura del dialog.
                            console.error("Failed to open dialog:", dialogAsyncResult.error.message);
                            // Permetti l'invio per evitare di bloccare l'utente in caso di problemi tecnici.
                            event.completed({ allowEvent: true });
                        }
                    }
                );
            } else {
                console.log("Specific recipient not found. Allowing event.");
                // Nessun destinatario specifico trovato, permetti l'invio.
                event.completed({ allowEvent: true });
            }
        } else {
            console.error("Failed to get recipients: " + asyncResult.error.message);
            // In caso di errore nel recupero dei destinatari, permetti l'invio per non bloccare.
            event.completed({ allowEvent: true });
        }
    });
});

// Questa funzione è necessaria per l'inizializzazione di Office JS.
// Viene chiamata automaticamente quando l'add-in è caricato.
Office.onReady(function(info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Office.js is ready for Outlook commands.");
    }
});
