function confirmSend(confirmed) {
  Office.context.ui.messageParent(confirmed ? "yes" : "no");
}
