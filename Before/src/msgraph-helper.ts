// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

/* 
    This file provides the provides functionality to get Microsoft Graph data. 
*/

import { ODataHelper} from './odata-helper';

export class MSGraphHelper {

    private static domain: string = "graph.microsoft.com";
    private static versionURLsegment: string = "/v1.0";
    
    // If any part of queryParamsSegment comes from user input,
    // be sure that it is sanitized so that it cannot be used in
    // a Response header injection attack.
    static getGraphData(accessToken: string, apiURLsegment: string, queryParamsSegment?: string) {
        return new Promise<any>(async (resolve, reject) => { 
            const oData = await ODataHelper.getData(accessToken, this.domain, apiURLsegment, this.versionURLsegment, queryParamsSegment);
            resolve(oData);
        })        
    }        
}
