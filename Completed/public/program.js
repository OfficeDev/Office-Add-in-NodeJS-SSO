// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides functions to get ask the Office host to get an access token to the add-in
	and to pass that token to the server to get Microsoft Graph data. 
*/
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add any initialization logic to this function.
        $("#getGraphAccessTokenButton").click(function () {
            getOneDriveFiles();
        });
    });
};

// This value is used to prevent the user from being
// cycled repeatedly through prompts to rerun the operation.
var timesGetOneDriveFilesHasRun = 0;

// This value is used to prevent the user from being
// cycled repeatedly between attempts to get the token with
// forceConsent and without it. 
var triedWithoutForceConsent = false;

// This value is used to prevent the user from being 
// cycled repeatedly if Microsoft Graph keeps erroring.
var timesMSGraphErrorReceived = false;

function getOneDriveFiles() {
    timesGetOneDriveFilesHasRun++;

    // Ask the Office host for an access token to the add-in. If the user is 
    // not signed in, s/he is prompted to sign in.
    triedWithoutForceConsent = true;
    getDataWithToken({ forceConsent: false });
}

function getDataWithToken(options) {
    Office.context.auth.getAccessTokenAsync(options,
        function (result) {
            if (result.status === "succeeded") {
                accessToken = result.value;
                getData("/api/values", accessToken);
            }
            else {
                handleClientSideErrors(result);
            }
        });
}

// Calls the specified URL or route (in the same domain as the add-in) 
// and includes the specified access token.
function getData(relativeUrl, accessToken) {

    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET"
    })
        .done(function (result) {
            showResult(result);
        })
        .fail(function (result) {
            handleServerSideErrors(result);
        });
}

function handleServerSideErrors(result) {

    // There are configurations of Azure Active Directory in which the user is required 
    // to provide additional authentication factor(s) to access some Microsoft Graph 
    // targets (e.g., OneDrive), even if the user can sign on to Office with just a 
    // password. In that case, AAD will return, to the add-in's web service, an error 50076 
    // response that has a Claims property. Server-side code passes this error back to the
    // client Have the Office host get a new token using the Claims string, which tells 
    // AAD to prompt the user for all required forms of authentication.
    if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 50076) {
        getDataWithToken({ authChallenge: result.responseJSON.error.innerError.claims });
    }

    // If consent was not granted (or was revoked) for one or more permissions,
    // the add-in's web service relays the 65001 error. Try to get the token
    // again with the forceConsent option.
    else if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 65001) {

        getDataWithToken({ forceConsent: true });
    }

    // If the add-in asks for an invalid scope (permission),
    // the add-in's web service relays the 70011 error. 
    else if (result.responseJSON.error.innerError
        && result.responseJSON.error.innerError.error_codes
        && result.responseJSON.error.innerError.error_codes[0] === 70011) {
        showResult(['The add-in is asking for a type of permission that is not recognized.']);
    }

    // If for any reason, the access_as_user scope (permission) is not in the token that
    // the client-side sends to the add-in's web service, it will respond with an error 
    // of the form "UnauthorizedError: JWT assertion failed: scp was ... ; expected access_as_user".
    else if (result.responseJSON.error.name
        && result.responseJSON.error.name.indexOf('expected access_as_user') !== -1) {
        showResult(['Microsoft Office does not have permission to get Microsoft Graph data on behalf of the current user.']);
    }

    // MS Graph may return an error; for example, if the token sent to MS Graph is expired or
    // invalid, a 401 error is sent. For any such error, start the whole process over,
    // but no more than once.
    else if (result.responseJSON.error.name
        && result.responseJSON.error.name.indexOf('Microsoft Graph error') !== -1) {
        if (!timesMSGraphErrorReceived) {
            timesMSGraphErrorReceived = true;
            timesGetOneDriveFilesHasRun = 0;
            triedWithoutForceConsent = false;
            getOneDriveFiles();
        } else {
            logError(result);
        }
    }
    // Any other error.  
    else {
        logError(result);
    }
}

function handleClientSideErrors(result) {

    switch (result.error.code) {

        case 13001:
            // The user is not logged in, or the user cancelled without responding a
            // prompt to provide a 2nd authentication factor. (See comment about two-
            // factor authentication in the fail callback of the getData method.)
            // Either way start over and force sign-in. 
            getDataWithToken({ forceAddAccount: true });
            break;
        case 13002:
            // The user's sign-in or consent was aborted. Ask the user to try again
            // but no more than once again.
            if (timesGetOneDriveFilesHasRun < 2) {
                showResult(['Your sign-in or consent was aborted before completion. Please try that operation again.']);
            } else {
                logError(result);
            }
            break;
        case 13003:
            // The user is logged in with an account that is neither work or school, nor Microsoft Account.
            showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
            break;
        case 13005:
            // The Office host has not been authorized to the add-in's web service
            // or the user has not granted the service permission to their `profile`.
            getDataWithToken({ forceConsent: true });
            break;
        case 13006:
            // Unspecified error in the Office host.
            showResult(['Please save your work, sign out of Office, close all Office applications, and restart this Office application.']);
            break;
        case 13007:
            // The Office host cannot get an access token to the add-ins web service/application.
            showResult(['That operation cannot be done at this time. Please try again later.']);
            break;
        case 13008:
            // The user triggered an operation that calls getAccessTokenAsync before a previous call of it completed.
            showResult(['Please try that operation again after the current operation has finished.']);
            break;
        case 13009:
            // The add-in does not support forcing consent. Try signing the user in without forcing consent, unless
            // that's already been tried.
            if (triedWithoutForceConsent) {
                showResult(['Please sign out of Office and sign in again with a work or school account, or Microsoft Account. Other kinds of accounts, like corporate domain accounts do not work.']);
            } else {
                getDataWithToken({ forceConsent: false });
            }
            break;
        default:
            logError(result);
            break;
    }
}

// Displays the data, assumed to be an array.
function showResult(data) {

    // Note that in this sample, the data parameter is an array of OneDrive file/folder
    // names. Encoding/sanitizing to protect against Cross-site scripting (XSS) attacks
    // is not needed because there are restrictions on what characters can be used in 
    // OneDrive file and folder names. These restrictions do not necessarily apply 
    // to other kinds of data including other kinds of Microsoft Graph data. So, to 
    // make this method safely reusable in other contexts, it uses the jQuery text() 
    // method which automatically encodes values that are passed to it.
	$.each(data, function (i) {
        var li = $('<li/>').addClass('ms-ListItem').appendTo($('#file-list'));
        var outerSpan = $('<span/>').addClass('ms-ListItem-secondaryText').appendTo(li);
        $('<span/>').addClass('ms-fontColor-themePrimary').appendTo(outerSpan).text(data[i]);
      });
}

function logError(result) {

    // Error messages can have a variety of structures depending on the ultimate
    // ultimate source and how intervening code restructures it before relaying it.
    console.log("Status: " + result.status);
    if (result.error.code) {
        console.log("Code: " + result.error.code);
    }
    if (result.error.name) {
        console.log("Code: " + result.error.name);
    }
    if (result.error.message) {
        console.log("Code: " + result.error.message);
    }
    if (result.responseJSON.error.name) {
        console.log("Code: " + result.responseJSON.error.name);
    }
    if (result.responseJSON.error.name) {
        console.log("Code: " + result.responseJSON.error.name);
    }
}
