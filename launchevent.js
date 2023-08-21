function onNewMessageComposeHandler(event) {
    const signature = "Jakob Christensen";

    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function (result) {
        event.completed();
    });
}

Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
