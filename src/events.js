
Office.initialize = () => { }

const setSignature = (eventType, event) => {
    Office.context.mailbox.item.body.setSignatureAsync(eventType, { coercionType : 'html' }, (asyncResult)=> {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error(result.error.message)
        }
        console.log(result.value)
        event.completed()
    })
}

Office.actions.associate('OnNewMessageCompose', (event) => {
    setSignature('OnNewMessageCompose', event)
})
Office.actions.associate('OnMessageRecipientsChanged', (event) => {
    setSignature('OnMessageRecipientsChanged', event)
})
