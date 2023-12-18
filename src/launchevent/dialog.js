
Office.onReady().then(() => {
    //  Office JS in the dialog might not be initiallised by the time the host tries to send the email data so send a confirmation message to confirm it is ready.
    Office.context.ui.messageParent(JSON.stringify({ messageType: 'initialise', message: 'Dialog is ready' }))
    
    //  Recieve emails from host page.
    // Office.context.ui.addHandlerAsync(
    //   Office.EventType.DialogParentMessageReceived, createEmailCheckBoxList)
    dontsend = function(){
        const cancelMessage = { messageType: 'cancel' }
        Office.context.ui.messageParent(JSON.stringify(cancelMessage))
    }
    sendanyways = function(){
        const sendanywaysMessage = { messageType: 'sendanyways' }
        Office.context.ui.messageParent(JSON.stringify(sendanywaysMessage))
    }
    
  })
  
  