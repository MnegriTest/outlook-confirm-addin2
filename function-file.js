Office.onReady(function(info) {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, onMessageSendHandler);
  }
});

function onMessageSendHandler(eventArgs) {
  Office.context.mailbox.item.to.getAsync(function(result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const recipients = result.value.map(r => r.emailAddress.toLowerCase());
      if (recipients.includes("github@mario.it")) {
        Office.context.ui.displayDialogAsync("https://mnegritest.github.io/outlook-confirm-addin/taskpane.html",
          { height: 30, width: 20, displayInIframe: true },
          function (asyncResult) {
            const dialog = asyncResult.value;
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
              if (arg.message === "yes") {
                dialog.close();
                eventArgs.completed({ allowEvent: true });
              } else {
                dialog.close();
                eventArgs.completed({ allowEvent: false });
              }
            });
          });
      } else {
        eventArgs.completed({ allowEvent: true });
      }
    } else {
      eventArgs.completed({ allowEvent: true });
    }
  });
}
