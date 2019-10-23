/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file shows how to use the SSO API to get a bootstrap token.
 */

Office.onReady(function(info) {
    $(document).ready(function() {
        $('#getGraphDataButton').click(getSSOToken);
    });
});

function getSSOToken() {
    Office.context.auth.getAccessTokenAsync(function (result) {
        if (result.status === "succeeded") {
            $('.welcome-body').hide();
            $('#message-area').show();   
            $('#message-area').text(result.value);
            getGraphToken(result.value);
        } else {
            $('.welcome-body').hide();
            $('#message-area').show();   
            $('#message-area').text("Please sign-in in the dialog.");

            $('#message-area').text(JSON.stringify(result));
          //  dialogFallback();
        }
    });
}

function getGraphToken(bootstrapToken) {
    $.ajax({type: "GET", 
		url: "/auth",
        headers: {"Authorization": "Bearer " + bootstrapToken }, 
        cache: false
    })
    .done(function (response) {
       $('.welcome-body').hide();
       $('#message-area').show();   
       $('#message-area').text(JSON.stringify(response));
       
       // If AAD rejects the On-Behalf-Of Flow attempt for any reason, the add-in's
       // server-side relays the error as an ordinary message with HTTP code 200,
       // so this done function executes. So, the code must check here for an error.
       // See comment below for the fail function.
       if (response.error) {
            console.log("result " + response.error);
            //handleServerSideErrors(result);
          //  dialogFallback();
       } else {
        makeGraphApiCall(response.access_token);
       }
    })
    // The fail function runs if the attempt to send the bootstrap token to
    // the add-in's client-side errors. But upstream errors from AAD or Microsoft
    // Graph do not trigger the fail function.
    .fail(function (result) {
        console.log("result " + result);
        dialogFallback();
		//handleServerSideErrors(result);

	});
}

// function forceConsent() {
//     Office.context.auth.getAccessTokenAsync({forceConsent:true}, function (result) {
//         if (result.status === "succeeded") {
//             // Use this token to call Web API
//             var ssoToken = result.value;
//             $('#ssoToken').val(result.value);
//         } else {
//             if (result.error.code === 13003) {
//                 // SSO is not supported for domain user accounts, only
//                 // work or school (Office 365) or Microsoft Account IDs.
//             } else {
//                 // Handle error
//             }
//         }
//     });
// }



// function claimsRequest() {
//     var claimsStr = JSON.parse($("#graphToken").val()).claims;
//     Office.context.auth.getAccessTokenAsync({authChallenge: claimsStr}, function (result) {
//         if (result.status === "succeeded") {
//             $('#ssoToken').val(result.value);
//         } else {
//             $('#ssoToken').val(JSON.stringify(result));
//         }
//     });
// }



