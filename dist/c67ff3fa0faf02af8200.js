Office.onReady().then((function(){Office.context.ui.messageParent(JSON.stringify({messageType:"initialise",message:"Dialog is ready"})),dontsend=function(){Office.context.ui.messageParent(JSON.stringify({messageType:"cancel"}))},sendanyways=function(){Office.context.ui.messageParent(JSON.stringify({messageType:"sendanyways"}))}}));