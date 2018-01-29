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
        getDataWithoutAuthChallenge();
    }

    // Called in the first attempt to use the on-behalf-of flow. The assumption
    // is that single factor authentication is all that is needed.
    function getDataWithoutAuthChallenge() {
        Office.context.auth.getAccessTokenAsync({forceConsent: false},
            function (result) {
                if (result.status === "succeeded") {
                    accessToken = result.value;
                    getData("/api/values", accessToken);
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
            // Turn off caching when debugging to force a fetch of data
            // with each call.
            cache: false
        })
        .done(function (result) {
            /*
              If the Microsoft Graph target requests addtional authentication
              factor(s), the result will not be data. It will be a Claims
              JSON telling AAD what addtional factors the user must provide.
              Start a new sign-on that passes this Claims string to AAD so that
              it will provide the needed prompts.
            */

            // If the result contains 'capolids', then it is the Claims string,
            // not the data.
            if (result[0].indexOf('capolids') !== -1) {
                result[0] = JSON.parse(result[0])
                getDataUsingAuthChallenge(result[0]);
            } else {
                showResult(result);
            }
        })
        .fail(function (result) {
          console.log(result.responseJSON.error);
        });
    }

    // Called to trigger a second sign-on in which the user will be prompted
    // to provide additional authentication factor(s). The authChallengeString
    // parameter tells AAD what factor(s) it should prompt for.
    function getDataUsingAuthChallenge(authChallengeString) {
        Office.context.auth.getAccessTokenAsync({authChallenge: authChallengeString},
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

// Displays the data, assumed to be an array.
function showResult(data) {
	for (var i = 0; i < data.length; i++) {
		$('#file-list').append('<li class="ms-ListItem">' +
		'<span class="ms-ListItem-secondaryText">' +
		  '<span class="ms-fontColor-themePrimary">' + data[i] + '</span>' +
		'</span></li>');
	}
}
