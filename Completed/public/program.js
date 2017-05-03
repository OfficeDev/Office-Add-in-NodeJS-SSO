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
                getOneDriveItems();
            });
    });
}

function getOneDriveItems() {

	// Ask the Office host for an access token to the add-in. If the user is 
	// not signed in, s/he is prompted to sign in.
	Office.context.auth.getAccessTokenAsync({ forceConsent: false },
		function (result) {
            if (result.status === "succeeded") {
                accessToken = result.value;
                getData("/api/onedriveitems", accessToken);
			}
			else {
				console.log("Code: " + result.error.code);
				console.log("Message: " + result.error.message);
				console.log("name: " + result.error.name);
				document.getElementById("getGraphAccessTokenButton").disabled = true;
			}
		});
}	

// Calls the specified URL or route (in the same domain as the add-in) 
// and includes the specified access token.
function getData(relativeUrl, accessToken) {

    $.ajax({
        url: relativeUrl,
        headers: { "Authorization": "Bearer " + accessToken },
        type: "GET",
    })
    .done(function (result) {
        showResult(result);
    })
    .fail(function (result) {
        console.log(result.error);
    });
}

// Displays the data, assumed to be an array.
function showResult(data) {	
	for (var i = 0; i < data.length; i++) {
		$('#file-list').append('<li class="ms-ListItem">' + 
		'<span class="ms-ListItem-secondaryText">' + 
		  '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' + 
		'</span></li>');
	}
}
