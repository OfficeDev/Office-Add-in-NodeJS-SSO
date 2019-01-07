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
