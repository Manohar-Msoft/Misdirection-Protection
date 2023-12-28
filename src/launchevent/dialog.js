
Office.onReady().then(() => {
    //  Office JS in the dialog might not be initiallised by the time the host tries to send the email data so send a confirmation message to confirm it is ready.
    Office.context.ui.messageParent(JSON.stringify({ messageType: 'initialise', message: 'Dialog is ready' }))
    
    //  Recieve emails from host page.
    Office.context.ui.addHandlerAsync(
      Office.EventType.DialogParentMessageReceived, displaymessage)
    
      dontsend = function(){
        const cancelMessage = { messageType: 'cancel' }
        Office.context.ui.messageParent(JSON.stringify(cancelMessage))
    }
    sendanyways = function(){
        const sendanywaysMessage = { messageType: 'sendanyways' }
        Office.context.ui.messageParent(JSON.stringify(sendanywaysMessage))
    }
    
  })

/**
 * Function that creates the check box form from the recipients.
 * @param {object} arg - The message object from the host pages that contains the recipient data.
 */
function displaymessage(arg){
    const allexternalEmails = JSON.parse(arg.message)
    recipientsTo = allexternalEmails[0]
    recipientsCc = allexternalEmails[1]
    recipientsBcc = allexternalEmails[2] 
    document.getElementById("toContainer").append(recipientsTo)
    document.getElementById("ccContainer").append(recipientsCc)
    document.getElementById("bccContainer").append(recipientsBcc)
        
}       
  