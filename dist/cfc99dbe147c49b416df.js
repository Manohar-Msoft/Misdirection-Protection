var customerDomain="@microsoft.com";function onMessageSendHandler(e){var o=[],c=!1,n=!1;function t(t,a){t.getAsync((function(t){t.status===Office.AsyncResultStatus.Succeeded&&t.value.forEach((function(e){e.emailAddress.includes(customerDomain)||o.push("".concat(e.emailAddress," in ").concat(a))})),"To"===a?c=!0:"CC"===a&&(n=!0),c&&n&&(o.length>0?window.confirm("Are You Sure?"):e.completed({allowEvent:!0}))}))}t(Office.context.mailbox.item.to,"To"),t(Office.context.mailbox.item.cc,"CC")}Office.context.platform!==Office.PlatformType.PC&&null!=Office.context.platform||Office.actions.associate("onMessageSendHandler",onMessageSendHandler);