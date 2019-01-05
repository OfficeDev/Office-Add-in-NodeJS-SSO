// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides functionality to get data from OData-compliant endppoints. 
*/

import * as https from 'https';

export class ODataHelper {

    static getData(accessToken: string, 
                   domain: string, 
                   apiURLsegment: string, 
                   apiVersion?: string, 
                   // If any part of queryParamsSegment comes from user input,
                   // be sure that it is sanitized so that it cannot be used in
                   // a Response header injection attack.
                   queryParamsSegment?: string) {

        return new Promise<any>((resolve, reject) => {
            var options = {
                host: domain,
                path: apiVersion + apiURLsegment + queryParamsSegment,
                method: 'GET',
                headers: {
                    'Content-Type': 'application/json',
                    Accept: 'application/json',
                    Authorization: 'Bearer ' + accessToken,
                    'Cache-Control': 'private, no-cache, no-store, must-revalidate',
                    'Expires': '-1',
                    'Pragma': 'no-cache'
                }
            };

            https.get(options, function (response) {
                var body = '';
                response.on('data', function (d) {
                        body += d;
                    });
                response.on('end', function () {
					
                    // TODO: Process the completed response from the OData endpoint
					//       and relay the data (or error) to the caller.
					
                });
            }).on('error', reject);
        });
    }
}