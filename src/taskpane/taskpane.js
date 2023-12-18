/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
// import * as appInsights from 'applicationinsights';
// appInsights.setup('a7c07799-bacf-41dd-9b74-2014f06f64ae');
// appInsights.start();
/* global document, Office */
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  // try {
  // //   Your Outlook code here

  // //   Log an event when the run function is executed
  //   appInsights.defaultClient.trackEvent({ name: 'RunFunctionExecuted' });
  // } catch (error) {
  //   // Log exceptions
  //   appInsights.defaultClient.trackException({ exception: error });
  // } finally {
  //   // Flush telemetry data
  //   appInsights.defaultClient.flush();
  // }
}
