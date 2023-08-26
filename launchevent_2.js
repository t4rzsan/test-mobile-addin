function onNewMessageComposeHandler(event) {
    console.log("onNewMessageComposeHandler");
    const item = Office.context.mailbox.item;
    
    const notification = {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "Jakob: onNewMessageComposeHandler",
        icon: "none",
        persistent: false                        
    };

    item.notificationMessages.addAsync("signature_notification", notification, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.log(result.error.message);
            event.completed({ allowEvent: false });
            return;
        }});

    const signature = "Signature from Jakob's amazing add-in";

    Office.context.mailbox.item.body.setSignatureAsync(signature, { coercionType: "html" }, function (result) {
        console.log("setSignatureAsync");

        event.completed();
    });
}

console.log("launchevent.js");
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
