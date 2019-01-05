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
    // to other kinds of Microsoft Graph data. If you copy this function in a production
    // add-in, be sure to protect against the possibility that data from a data source 
    // contains a Cross-site scripting string by encoding or sanitizing the data before 
    // it is inserted into an HTML page.
	for (var i = 0; i < data.length; i++) {
		$('#file-list').append('<li class="ms-ListItem">' + 
		'<span class="ms-ListItem-secondaryText">' + 
		  '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' + 
		'</span></li>');
	}
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
