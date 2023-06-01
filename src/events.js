
Office.initialize = () => { }

const setSignature = (eventType, event) => {
    Office.context.mailbox.item.body.setSignatureAsync(eventType, { coercionType : 'html' }, (asyncResult)=> {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(asyncResult.error.message)
        }
        console.log(asyncResult.value)
        event.completed()
    })
}

Office.actions.associate('OnNewMessageCompose', (event) => {
    setSignature('OnNewMessageCompose', event)
})
Office.actions.associate('OnMessageRecipientsChanged', (event) => {
    setSignature('OnMessageRecipientsChanged', event)
})

Office.actions.associate('onMessageSendHandler', (event) => {
    Office.context.mailbox.item.body.getAsync(
        "text",
        { asyncContext: event },
        getBodyCallback
      );
})

function getBodyCallback(asyncResult){
    let event = asyncResult.asyncContext;
    let message = "Envoie interdit";
    event.completed({ allowEvent: false ,errorMessage: message });
    return;
  }