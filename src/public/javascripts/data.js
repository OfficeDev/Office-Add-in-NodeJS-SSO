/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to open call Microsoft Graph.
 */

function makeGraphApiCall(accessToken) {
    $.ajax({type: "GET", 
        url: "/getuserdata",
        headers: {"access_token": accessToken },
        cache: false
    }).done(function (response) {

        writeFileNamesToOfficeDocument(response)
        .then(function () { 
            $('.welcome-body').hide();
            $('#message-area').show();   
            $('#message-area').text("Your data has been added to the document."); 
        })
        .catch(function (error) {
            $('.welcome-body').hide();
            $('#message-area').show();     
            $('#message-area').text(JSON.stringify(error.toString()));
        });
    })
    .fail(function (result) {
        console.log("result " + result);
		//handleServerSideErrors(result);
	});
}
