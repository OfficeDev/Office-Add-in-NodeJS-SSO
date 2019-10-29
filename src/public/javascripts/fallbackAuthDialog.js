/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use MSAL.js to get an access token to Microsoft Graph an pass it to the task pane.
 */

    Office.initialize = function () { 
        if (Office.context.ui.messageParent)
        {
            userAgentApp.handleRedirectCallback(authCallback);
           // userAgentApp.loginRedirect(requestObj);
            userAgentApp.acquireTokenRedirect(requestObj);
        }
     };

    const msalConfig = {
        auth: {
            clientId: "$application_GUID here$", //This is your client ID
            authority: "https://login.microsoftonline.com/common", 
            redirectURI: "https://localhost:44355/dialog.html", 
            navigateToLoginRequestUrl: false,
            response_type: "access_token"
        },
        cache: {
            cacheLocation: 'localStorage', // Needed to avoid "user login is required" error.
            storeAuthStateInCookie: true  // Recommended to avoid certain IE/Edge issues.
        }
    };

    var requestObj = {
        scopes: ["https://graph.microsoft.com/User.Read", 
        "https://graph.microsoft.com/Files.Read.All"]
    };

    const userAgentApp = new Msal.UserAgentApplication(msalConfig);

    function authCallback(error, response) {
        if (error) {
            console.log(error);
            Office.context.ui.messageParent(JSON.stringify({ status: 'failure', result : error }));
        } else {
            if (response.tokenType === "id_token") {
                console.log(response.idToken.rawIdToken);
            } else {
                console.log("token type is:" + response.tokenType);
                Office.context.ui.messageParent( JSON.stringify({ status: 'success', result : response.accessToken }) );               
            }        
        }
    }
