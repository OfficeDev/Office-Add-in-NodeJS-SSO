/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. 
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */

  // If the add-in is running in Internet Explorer, the code must add support 
 // for Promises.
if (!window.Promise) {
    window.Promise = Office.Promise;
}

Office.onReady(function(info) {
    $(document).ready(function() {
        $('#getGraphDataButton').click(getGraphData);
    });
});

let retryGetAccessToken = 0;

async function getGraphData() {
    try {
        let bootstrapToken = await OfficeRuntime.auth.getAccessToken({ allowSignInPrompt: true, allowConsentPrompt: true, forMSGraphAccess: true });
        let exchangeResponse = await getGraphToken(bootstrapToken);
        if (exchangeResponse.claims) {
            // Microsoft Graph requires an additional form of authentication. Have the Office host 
            // get a new token using the Claims string, which tells AAD to prompt the user for all 
            // required forms of authentication.
            let mfaBootstrapToken = await OfficeRuntime.auth.getAccessToken({ authChallenge: exchangeResponse.claims });
            exchangeResponse = await getGraphToken(mfaBootstrapToken);
        }
        
        if (exchangeResponse.error) {
            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            handleAADErrors(exchangeResponse);
        } 
        else 
        {
            // For debugging:
            // showMessage("ACCESS TOKEN: " + JSON.stringify(exchangeResponse.access_token));

            // makeGraphApiCall makes an AJAX call to the MS Graph endpoint. Errors are caught
            // in the .fail callback of that call, not in the catch block below.
            makeGraphApiCall(exchangeResponse.access_token);
        }
    }
    catch(exception) {
        // The only exceptions caught here are exceptions in your code in the try block
        // and errors returned from the call of `getAccessToken` above.
        if (exception.code) { 
            handleClientSideErrors(exception);
        }
        else {
            showMessage("EXCEPTION: " + JSON.stringify(exception));
        } 
    }
}

async function getGraphToken(bootstrapToken) {
    let response = await $.ajax({type: "GET", 
		url: "/auth",
        headers: {"Authorization": "Bearer " + bootstrapToken }, 
        cache: false
    });
    return response;
}

function handleClientSideErrors(error) {
    switch (error.code) {

        case 13001:
            // No one is signed into Office. If the add-in cannot be effectively used when no one 
            // is logged into Office, then the first call of getAccessToken should pass the 
            // `allowSignInPrompt: true` option. Since this sample does that, you should not see this error
            showMessage("No one is signed into Office. But you can use many of the add-ins functions anyway. If you want to log in, press the Get OneDrive File Names button again.");  
            break;
        case 13002:
            // The user aborted the consent prompt. If the add-in cannot be effectively used when consent
            // has not been granted, then the first call of getAccessToken should pass the `allowConsentPrompt: true` option.
            showMessage("You can use many of the add-ins functions even though you have not granted consent. If you want to grant consent, press the Get OneDrive File Names button again."); 
            break;
        case 13006:
            // Only seen in Office on the web.
            showMessage("Office on the web is experiencing a problem. Please sign out of Office, close the browser, and then start again."); 
            break;
        case 13008:
            // Only seen in Office on the web.
            showMessage("Office is still working on the last operation. When it completes, try this operation again."); 
            break;
        case 13010:
            // Only seen in Office on the web.
            showMessage("Follow the instructions to change your browser's zone configuration.");
            break;
        default:
            // For all other errors, including 13000, 13003, 13005, 13007, 13012, and 50001, fall back
            // to non-SSO sign-in.
            dialogFallback();
            break;
    }
}

function handleAADErrors(exchangeResponse) {
    // On rare occasions the bootstrap token is unexpired when Office validates it,
    // but expires by the time it is sent to AAD for exchange. AAD will respond
    // with "The provided value for the 'assertion' is not valid. The assertion has expired."
    // Retry the call of getAccessToken (no more than once). This time Office will return a 
    // new unexpired bootstrap token. 
    if ((exchangeResponse.error_description.indexOf("AADSTS500133") !== -1)
        &&
        (retryGetAccessToken <= 0)) 
    {
        retryGetAccessToken++;
        getGraphData();
    }
    else 
    {
        // For all other AAD errors, fallback to non-SSO sign-in.
        // For debugging: 
        // showMessage("AAD ERROR: " + JSON.stringify(exchangeResponse));                   
        dialogFallback();
    }
}
